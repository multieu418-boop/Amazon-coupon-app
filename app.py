import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from datetime import datetime, date

# --- 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 侧边栏指引 ---
st.sidebar.header("📂 必备文件上传指引")
inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])

# --- 初始化 Session State ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'field_configs' not in st.session_state:
    st.session_state.field_configs = []

# --- 1. 动态解析第7行标题及第8/9行选项 ---
if template_file:
    try:
        template_file.seek(0)
        # 使用 data_only=True 以获取单元格显示的文本而非公式
        wb_preview = load_workbook(template_file, data_only=True)
        ws_preview = wb_preview.active
        
        configs = []
        # 扫描第7行（标题行），前15列通常涵盖了所有优惠券字段
        for col in range(1, 16):
            title = ws_preview.cell(row=7, column=col).value
            if title:
                title_str = str(title).strip()
                
                # 检查第8行和第9行是否有预设值作为下拉参考
                ref_vals = []
                for r in [8, 9]:
                    v = ws_preview.cell(row=r, column=col).value
                    if v: ref_vals.append(str(v).strip())
                
                # 预设逻辑：识别常见的亚马逊下拉字段
                options = None
                # 如果标题包含特定关键字，提供标准选项
                if any(x in title_str for x in ["折扣类型", "Discount Type"]):
                    options = ["Percentage", "Money"]
                elif any(x in title_str for x in ["只能兑换一次", "Limit"]):
                    options = ["Yes", "No"]
                elif any(x in title_str for x in ["目标买家", "Target Audience"]):
                    options = ["All Customers", "Amazon Prime Members"]
                elif any(x in title_str for x in ["叠加", "Stack"]):
                    options = ["Yes", "No"]
                elif any(x in title_str for x in ["优惠券类型", "Coupon Type"]):
                    options = ["Standard", "Subscribe & Save"]
                
                configs.append({
                    "col": col,
                    "label": title_str,
                    "options": options
                })
        
        st.session_state.field_configs = configs
        st.sidebar.success(f"✅ 已成功识别第7行标题")
    except Exception as e:
        st.sidebar.error(f"解析模板失败: {e}")

# --- 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：基于第7行动态生成", "第二阶段：报错纠错修复"])

with tab1:
    if not template_file or not st.session_state.field_configs:
        st.info("💡 请先在左侧上传『上传Coupon文件模板』。程序会自动提取第7行的需求字段。")
    else:
        st.header(f"1️⃣ 录入需求（基于模板字段）")
        
        with st.form("dynamic_coupon_form", clear_on_submit=True):
            user_responses = {}
            # 动态生成表单布局
            grid = st.columns(2)
            for idx, cfg in enumerate(st.session_state.field_configs):
                with grid[idx % 2]:
                    label_name = cfg['label']
                    unique_key = f"field_{cfg['col']}"
                    
                    if cfg['options']:
                        # 如果识别到下拉选项
                        user_responses[cfg['col']] = st.selectbox(label_name, options=cfg['options'], key=unique_key)
                    elif any(x in label_name for x in ["ASIN", "列表"]):
                        # 如果是 ASIN 列表，使用大输入框
                        user_responses[cfg['col']] = st.text_area(label_name, placeholder="分号分隔 ASIN", key=unique_key)
                    elif any(x in label_name for x in ["日期", "Date"]):
                        # 如果是日期字段
                        user_responses[cfg['col']] = st.date_input(label_name, value=date.today() + timedelta(days=1), key=unique_key)
                    else:
                        # 普通文本输入
                        user_responses[cfg['col']] = st.text_input(label_name, key=unique_key)
            
            if st.form_submit_button("➕ 添加到需求列表"):
                # 处理数据格式转换
                formatted_row = {}
                for c_idx, val in user_responses.items():
                    if isinstance(val, (date, datetime)):
                        formatted_row[c_idx] = val.strftime("%m/%d/%Y")
                    else:
                        formatted_row[c_idx] = str(val) if val is not None else ""
                st.session_state.coupon_pool.append(formatted_row)
                st.toast("需求已添加")

        # 预览与生成文件
        if st.session_state.coupon_pool:
            st.subheader("📋 待填充的需求池预览")
            # 显示时将列索引换回标题文字
            display_map = {c['col']: c['label'] for c in st.session_state.field_configs}
            st.dataframe(pd.DataFrame(st.session_state.coupon_pool).rename(columns=display_map))
            
            b1, b2 = st.columns(2)
            if b1.button("🗑️ 清空列表"):
                st.session_state.coupon_pool = []
                st.rerun()
            
            if b2.button("🚀 填充模板并下载"):
                template_file.seek(0)
                final_wb = load_workbook(template_file)
                final_ws = final_wb.active
                
                # --- 核心：寻找真正的起始空行 ---
                # 从第8行开始往下扫描，直到发现第一列（ASIN列）为空
                start_fill_row = 8
                while True:
                    cell_val = final_ws.cell(row=start_fill_row, column=1).value
                    if cell_val is not None and len(str(cell_val).strip()) > 0:
                        start_fill_row += 1
                    else:
                        break
                
                # 批量写入用户添加的所有需求
                for offset, data_row in enumerate(st.session_state.coupon_pool):
                    target_row = start_fill_row + offset
                    for col_idx, value in data_row.items():
                        final_ws.cell(row=target_row, column=int(col_idx)).value = value
                
                # 保存并提供下载
                out_io = io.BytesIO()
                final_wb.save(out_io)
                st.success(f"✅ 填充完毕！已自动避开示例，从第 {start_fill_row} 行开始填充。")
                st.download_button(
                    label="💾 点击下载填充后的 Excel",
                    data=out_io.getvalue(),
                    file_name=f"Coupon_Batch_{date.today().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# 第二阶段（报错处理）保持原有逻辑...
