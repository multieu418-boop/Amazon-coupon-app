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
if 'final_xlsx' not in st.session_state:
    st.session_state.final_xlsx = None

# --- 1. 动态解析函数 (独立于 UI) ---
def parse_template(file):
    try:
        file.seek(0)
        # 必须使用 data_only=True
        wb = load_workbook(file, data_only=True)
        ws = wb.active
        configs = []
        # 严格读取第 7 行标题
        for col in range(1, 16):
            title = ws.cell(row=7, column=col).value
            if title:
                title_str = str(title).strip()
                # 预设下拉逻辑
                opts = None
                if any(x in title_str for x in ["折扣类型", "Discount Type"]):
                    opts = ["Percentage", "Money"]
                elif any(x in title_str for x in ["只能兑换一次", "Limit"]):
                    opts = ["Yes", "No"]
                elif any(x in title_str for x in ["目标买家", "Target"]):
                    opts = ["All Customers", "Amazon Prime Members"]
                elif any(x in title_str for x in ["叠加", "Stack"]):
                    opts = ["Yes", "No"]
                
                configs.append({"col": col, "label": title_str, "options": opts})
        return configs
    except:
        return []

# 当模板上传后，立即解析字段
if template_file:
    st.session_state.field_configs = parse_template(template_file)
    st.sidebar.success(f"✅ 模板第7行解析成功")

# --- 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：动态表单生成", "第二阶段：报错纠错修复"])

with tab1:
    if not template_file or not st.session_state.field_configs:
        st.info("💡 请先在左侧上传『上传Coupon文件模板』。系统将自动识别第7行标题并生成输入框。")
    else:
        st.header("1️⃣ 录入优惠券需求")
        
        # 将表单逻辑包裹在完整的 with 块中，确保必须有 submit button
        with st.form("coupon_input_form", clear_on_submit=True):
            user_responses = {}
            grid = st.columns(2)
            
            for idx, cfg in enumerate(st.session_state.field_configs):
                with grid[idx % 2]:
                    lbl = cfg['label']
                    fid = f"f_{cfg['col']}"
                    
                    if cfg['options']:
                        user_responses[cfg['col']] = st.selectbox(lbl, options=cfg['options'], key=fid)
                    elif any(x in lbl for x in ["ASIN", "列表"]):
                        user_responses[cfg['col']] = st.text_area(lbl, placeholder="ASIN用分号分隔", key=fid)
                    elif any(x in lbl for x in ["日期", "Date"]):
                        user_responses[cfg['col']] = st.date_input(lbl, value=date.today() + timedelta(days=1), key=fid)
                    else:
                        user_responses[cfg['col']] = st.text_input(lbl, key=fid)
            
            # 表单内必须有这个按钮
            add_submitted = st.form_submit_button("➕ 添加此条需求至列表")
            
            if add_submitted:
                new_row = {}
                for c_idx, val in user_responses.items():
                    if isinstance(val, (date, datetime)):
                        new_row[c_idx] = val.strftime("%m/%d/%Y")
                    else:
                        new_row[c_idx] = str(val) if val is not None else ""
                st.session_state.coupon_pool.append(new_row)
                st.toast("已记录一项需求！")

        # --- 预览与生成 ---
        if st.session_state.coupon_pool:
            st.subheader("📋 待处理需求池")
            name_map = {c['col']: c['label'] for c in st.session_state.field_configs}
            st.dataframe(pd.DataFrame(st.session_state.coupon_pool).rename(columns=name_map))
            
            c1, c2 = st.columns(2)
            if c1.button("🗑️ 清空所有记录"):
                st.session_state.coupon_pool = []
                st.session_state.final_xlsx = None
                st.rerun()
            
            if c2.button("🚀 填充模板并生成文件"):
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active
                
                # 寻找空行：从第8行起，直到A列为空
                start_row = 8
                while ws.cell(row=start_row, column=1).value is not None:
                    if len(str(ws.cell(row=start_row, column=1).value).strip()) == 0:
                        break
                    start_row += 1
                
                # 写入
                for i, data in enumerate(st.session_state.coupon_pool):
                    target_r = start_row + i
                    for col_idx, value in data.items():
                        ws.cell(row=target_r, column=int(col_idx)).value = value
                
                out = io.BytesIO()
                wb.save(out)
                st.session_state.final_xlsx = out.getvalue()
                st.success(f"✅ 数据已填充，从第 {start_row} 行开始。")

    # 下载按钮放在 Form 之外
    if st.session_state.final_xlsx:
        st.download_button(
            label="💾 下载生成的上传文件",
            data=st.session_state.final_xlsx,
            file_name=f"Coupon_Upload_{date.today().strftime('%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 第二阶段（修复）代码略...
