import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
import datetime  # 统一导入方式，防止 NameError

# --- 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 初始化所有 Session State 变量 (防止 NameError) ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'field_configs' not in st.session_state:
    st.session_state.field_configs = []
if 'final_xlsx_data' not in st.session_state:
    st.session_state.final_xlsx_data = None

# --- 1. 侧边栏：文件上传与指引 ---
st.sidebar.header("📂 必备文件上传指引")
st.sidebar.info("""
1. **All Listing Report**: 获取价格校验（TXT格式）。
2. **上传Coupon文件模板**: 亚马逊下载的空白 Excel 模板。
   - 程序将自动读取第7行作为表单标题。
   - 程序将自动识别下拉选项并生成菜单。
""")

inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])

# --- 2. 动态解析函数 (增加防御性处理) ---
def parse_amazon_template(file):
    if file is None:
        return []
    try:
        # 重置文件指针
        file.seek(0)
        # 必须使用 data_only=True 获取显示值
        wb = load_workbook(file, data_only=True)
        ws = wb.active
        configs = []
        
        # 扫描第7行标题
        for col in range(1, 21): # 扫描前20列确保覆盖全面
            title = ws.cell(row=7, column=col).value
            if title:
                title_str = str(title).strip()
                
                # 预设下拉选项逻辑
                opts = None
                if any(x in title_str for x in ["折扣类型", "Discount Type"]):
                    opts = ["Percentage", "Money"]
                elif any(x in title_str for x in ["兑换一次", "Limit"]):
                    opts = ["Yes", "No"]
                elif any(x in title_str for x in ["目标买家", "Target"]):
                    opts = ["All Customers", "Amazon Prime Members"]
                elif any(x in title_str for x in ["叠加", "Stack"]):
                    opts = ["Yes", "No"]
                elif any(x in title_str for x in ["优惠券类型", "Coupon Type"]):
                    opts = ["Standard", "Subscribe & Save"]
                
                configs.append({
                    "col": col,
                    "label": title_str,
                    "options": opts
                })
        return configs
    except Exception as e:
        st.error(f"解析模板时出错: {e}")
        return []

# 当文件上传且配置为空时，执行解析
if template_file and not st.session_state.field_configs:
    st.session_state.field_configs = parse_amazon_template(template_file)

# --- 3. 主界面逻辑 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：动态表单录入", "第二阶段：报错纠错修复"])

with tab1:
    if not template_file or not st.session_state.field_configs:
        st.warning("👋 请先在左侧上传『上传Coupon文件模板』，我会根据模板第7行自动生成输入框。")
    else:
        st.header("1️⃣ 录入优惠券需求")
        
        # 使用 Form 确保提交逻辑完整
        with st.form("coupon_input_form", clear_on_submit=True):
            st.write("📝 请根据模板字段填写信息：")
            current_inputs = {}
            grid = st.columns(2)
            
            for idx, cfg in enumerate(st.session_state.field_configs):
                with grid[idx % 2]:
                    label = cfg['label']
                    unique_id = f"col_{cfg['col']}"
                    
                    if cfg['options']:
                        current_inputs[cfg['col']] = st.selectbox(label, options=cfg['options'], key=unique_id)
                    elif any(x in label.upper() for x in ["ASIN", "列表"]):
                        current_inputs[cfg['col']] = st.text_area(label, placeholder="多个ASIN请用分号分隔", key=unique_id)
                    elif any(x in label for x in ["日期", "Date"]):
                        # 使用完整的 datetime 路径防止 NameError
                        default_date = datetime.date.today() + datetime.timedelta(days=1)
                        current_inputs[cfg['col']] = st.date_input(label, value=default_date, key=unique_id)
                    else:
                        current_inputs[cfg['col']] = st.text_input(label, key=unique_id)
            
            submit_btn = st.form_submit_button("➕ 添加到待填充列表")
            
            if submit_btn:
                # 数据转换
                entry = {}
                for c_idx, val in current_inputs.items():
                    if isinstance(val, (datetime.date, datetime.datetime)):
                        entry[c_idx] = val.strftime("%m/%d/%Y")
                    else:
                        entry[c_idx] = str(val) if val is not None else ""
                
                st.session_state.coupon_pool.append(entry)
                st.toast("需求已保存到下方列表")

        # --- 预览与最终生成 ---
        if st.session_state.coupon_pool:
            st.subheader("📋 待处理需求预览")
            name_map = {c['col']: c['label'] for c in st.session_state.field_configs}
            st.dataframe(pd.DataFrame(st.session_state.coupon_pool).rename(columns=name_map))
            
            btn_col1, btn_col2 = st.columns(2)
            if btn_col1.button("🗑️ 清空列表"):
                st.session_state.coupon_pool = []
                st.session_state.final_xlsx_data = None
                st.rerun()
            
            if btn_col2.button("🚀 开始填充模板并生成"):
                try:
                    template_file.seek(0)
                    wb = load_workbook(template_file)
                    ws = wb.active 
                    
                    # 智能寻找起始空行 (从第8行起，直到第一列为空)
                    start_row = 8
                    while ws.cell(row=start_row, column=1).value is not None:
                        if not str(ws.cell(row=start_row, column=1).value).strip():
                            break
                        start_row += 1
                    
                    # 写入数据
                    for i, data in enumerate(st.session_state.coupon_pool):
                        row_idx = start_row + i
                        for c_idx, value in data.items():
                            ws.cell(row=row_idx, column=int(c_idx)).value = value
                    
                    # 导出
                    buf = io.BytesIO()
                    wb.save(buf)
                    st.session_state.final_xlsx_data = buf.getvalue()
                    st.success(f"✅ 已成功避开示例行，从第 {start_row} 行开始填充。")
                except Exception as ex:
                    st.error(f"填充失败: {ex}")

    # 下载按钮
    if st.session_state.final_xlsx_data:
        st.download_button(
            label="💾 下载填充好的 Coupon 上传文件",
            data=st.session_state.final_xlsx_data,
            file_name=f"Cupshe_Coupon_{datetime.date.today().strftime('%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
