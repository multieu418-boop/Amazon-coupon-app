import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from datetime import datetime, date

# --- 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 侧边栏 ---
st.sidebar.header("📂 必备文件上传指引")
inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report (TXT)", type=['txt'])
template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板 (Excel)", type=['xlsx'])

# --- 1. 动态模板解析逻辑 ---
def get_template_config(file):
    if not file:
        return None, []
    
    try:
        file.seek(0)
        wb = load_workbook(file, data_only=True)
        ws = wb.active
        
        header_row = 7 # 默认第7行
        # 自动定位标题行：扫描前20行，寻找包含关键信息的行
        for r in range(1, 21):
            row_values = [str(ws.cell(row=r, column=c).value) for c in range(1, 10)]
            if any("ASIN" in v.upper() or "折扣" in v or "Discount" in v for v in row_values):
                header_row = r
                break
        
        configs = []
        # 扫描标题行的前 15 列
        for col in range(1, 16):
            cell_val = ws.cell(row=header_row, column=col).value
            if cell_val:
                label = str(cell_val).strip()
                
                # 寻找该列是否有下拉建议（参考标题下方的两行）
                ref_1 = ws.cell(row=header_row+1, column=col).value
                ref_2 = ws.cell(row=header_row+2, column=col).value
                
                options = None
                # 逻辑：如果参考行有值，或者是已知的亚马逊标准选项
                standard_opts = {
                    "折扣类型": ["Percentage", "Money"],
                    "每位买家只能兑换一次": ["Yes", "No"],
                    "优惠券类型": ["Standard", "Subscribe & Save"],
                    "目标买家": ["All Customers", "Amazon Prime Members"],
                    "叠加": ["Yes", "No"]
                }
                
                for k, v in standard_opts.items():
                    if k in label:
                        options = v
                        break
                
                # 如果没匹配到标准选项，但参考行有值，也可以作为参考提示
                configs.append({
                    "col": col,
                    "label": label,
                    "options": options
                })
        
        return header_row, configs
    except Exception as e:
        st.error(f"解析模板失败: {e}")
        return 7, []

# --- 2. 初始化 Session ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []

# --- 核心操作流 ---
header_row_idx, dynamic_configs = get_template_config(template_file)

if template_file:
    st.sidebar.success(f"✅ 模板已识别 (标题行: {header_row_idx})")

# --- 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")
tab1, tab2 = st.tabs(["第一阶段：模板批量填充", "第二阶段：报错纠错修复"])

with tab1:
    if not template_file or not dynamic_configs:
        st.info("💡 请先上传『上传Coupon文件模板』以加载字段。")
    else:
        st.header("1️⃣ 录入优惠券需求")
        with st.form("coupon_form", clear_on_submit=True):
            user_input_data = {}
            grid = st.columns(2)
            for i, cfg in enumerate(dynamic_configs):
                with grid[i % 2]:
                    lbl = cfg['label']
                    k = f"col_{cfg['col']}"
                    
                    if cfg['options']:
                        user_input_data[cfg['col']] = st.selectbox(lbl, options=cfg['options'], key=k)
                    elif any(x in lbl for x in ["ASIN", "列表"]):
                        user_input_data[cfg['col']] = st.text_area(lbl, placeholder="多个ASIN用分号分隔", key=k)
                    elif any(x in lbl for x in ["日期", "Date"]):
                        user_input_data[cfg['col']] = st.date_input(lbl, value=date.today() + timedelta(days=1), key=k)
                    else:
                        user_input_data[cfg['col']] = st.text_input(lbl, key=k)
            
            if st.form_submit_button("➕ 添加至待生成列表"):
                # 统一转为字符串处理，防止 Excel 写入类型错误
                row_dict = {}
                for c_idx, val in user_input_data.items():
                    if isinstance(val, (date, datetime)):
                        row_dict[c_idx] = val.strftime("%m/%d/%Y")
                    else:
                        row_dict[c_idx] = str(val) if val is not None else ""
                st.session_state.coupon_pool.append(row_dict)
                st.toast("需求已记录")

        # 预览与导出
        if st.session_state.coupon_pool:
            st.subheader("📋 待填充列表")
            preview_df = pd.DataFrame(st.session_state.coupon_pool)
            # 重命名列名用于显示
            name_map = {c['col']: c['label'] for c in dynamic_configs}
            st.dataframe(preview_df.rename(columns=name_map), use_container_width=True)
            
            if st.button("🚀 填充并生成 Excel"):
                template_file.seek(0)
                final_wb = load_workbook(template_file)
                final_ws = final_wb.active
                
                # --- 智能寻找真正空白行 ---
                # 从标题行下一行开始找
                target_r = header_row_idx + 1
                while True:
                    # 检查 A 列是否有值，如果有，或者是“示例”字样，继续往下
                    cell_a = final_ws.cell(row=target_r, column=1).value
                    if cell_a is not None and len(str(cell_a).strip()) > 0:
                        target_r += 1
                    else:
                        break
                
                # 写入数据
                for offset, data in enumerate(st.session_state.coupon_pool):
                    write_row = target_r + offset
                    for c_idx, c_val in data.items():
                        final_ws.cell(row=write_row, column=int(c_idx)).value = c_val
                
                out_stream = io.BytesIO()
                final_wb.save(out_stream)
                
                st.success(f"✅ 已避开前 {target_r-1} 行，数据从第 {target_r} 行开始填充。")
                st.download_button(
                    label="💾 下载最终填充版 Excel",
                    data=out_stream.getvalue(),
                    file_name=f"Coupon_Final_{date.today().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
