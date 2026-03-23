import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
# 注意：这里必须是小写的 dataframe_to_rows
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, timedelta

# --- 页面设置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 侧边栏指引 ---
st.sidebar.header("📂 必备文件上传指引")
st.sidebar.markdown("""
1. **All Listing Report**: 从后台导出的 TXT 报告，用于校验价格。
2. **上传Coupon文件模板**: 亚马逊下载的空白 Excel 批量上传模板。
""")

# --- 数据加载函数 ---
@st.cache_data
def load_inventory(file):
    if file:
        try:
            content = file.read()
            for encoding in ['utf-8', 'utf-16', 'cp1252', 'gbk']:
                try:
                    df = pd.read_csv(io.BytesIO(content), sep='\t', encoding=encoding, on_bad_lines='skip')
                    if 'asin1' in df.columns:
                        return df[['asin1', 'price']].drop_duplicates('asin1').set_index('asin1')
                except: continue
            return "ERROR_ENCODING"
        except: return "ERROR_READ"
    return None

# --- 初始化 Session State (存储多条需求) ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'final_excel_data' not in st.session_state:
    st.session_state.final_excel_data = None

# --- 文件上传状态监控 ---
inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
inv_data = load_inventory(inventory_file)
if inv_data is not None:
    if not isinstance(inv_data, str):
        st.sidebar.success("✅ All Listing Report 已就绪")

template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])
if template_file:
    st.sidebar.success("✅ 上传Coupon文件模板 已就绪")

# --- 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：多需求批量填充", "第二阶段：报错纠错修复"])

with tab1:
    st.header("1️⃣ 录入优惠券需求")
    
    # 输入区域
    with st.expander("➕ 添加新需求到列表", expanded=True):
        with st.form("coupon_input_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                asin_raw = st.text_area("子 ASIN 列表 (分号或换行分隔)", height=150)
                d_type = st.selectbox("折扣类型", ["折扣 (Percentage)", "满减 (Money)"])
                d_val = st.number_input("折扣数值", min_value=1.0, value=5.0)
            with col2:
                c_name = st.text_input("优惠券名称后缀", value="Summer Sale")
                c_budget = st.number_input("预算 (€)", min_value=100, value=1000)
                c_range = st.date_input("活动日期范围", [datetime.now() + timedelta(days=1), datetime.now() + timedelta(days=30)])
            
            add_btn = st.form_submit_button("添加需求到下方列表")
            
            if add_btn:
                if not asin_raw:
                    st.error("请填入 ASIN！")
                else:
                    # 清洗 ASIN
                    clean_asins = ";".join([a.strip() for a in re.split(r'[;\n,\s\t]+', asin_raw) if a.strip()])
                    st.session_state.coupon_pool.append({
                        "ASINs": clean_asins,
                        "Type": "Percentage" if "折扣" in d_type else "Money",
                        "Value": d_val,
                        "Name": c_name,
                        "Budget": int(c_budget),
                        "Start": c_range[0].strftime("%m/%d/%Y") if len(c_range)>0 else "",
                        "End": c_range[1].strftime("%m/%d/%Y") if len(c_range)>1 else ""
                    })
                    st.success("已添加！")

    # 列表展示与生成
    if st.session_state.coupon_pool:
        st.subheader("📋 待填充的需求池")
        pool_df = pd.DataFrame(st.session_state.coupon_pool)
        st.dataframe(pool_df, use_container_width=True)
        
        c1, c2 = st.columns(2)
        if c1.button("🗑️ 清空所有需求"):
            st.session_state.coupon_pool = []
            st.rerun()
            
        if c2.button("🚀 填充模板并生成文件"):
            if not template_file:
                st.error("请先在左侧上传『上传Coupon文件模板』！")
            else:
                # 开始操作 Excel 模板
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active # 默认填充第一个工作表
                
                # 核心填充逻辑：从第 7 行开始
                for i, row_data in enumerate(st.session_state.coupon_pool):
                    line = 7 + i
                    ws.cell(row=line, column=1).value = row_data["ASINs"]
                    ws.cell(row=line, column=2).value = row_data["Type"]
                    # 折扣填在 C 列，满减填在 D 列
                    if row_data["Type"] == "Percentage":
                        ws.cell(row=line, column=3).value = row_data["Value"]
                    else:
                        ws.cell(row=line, column=4).value = row_data["Value"]
                    
                    ws.cell(row=line, column=5).value = row_data["Name"]
                    ws.cell(row=line, column=6).value = row_data["Budget"]
                    ws.cell(row=line, column=7).value = row_data["Start"]
                    ws.cell(row=line, column=8).value = row_data["End"]
                    ws.cell(row=line, column=9).value = "Yes"
                    ws.cell(row=line, column=11).value = "All Customers"

                output = io.BytesIO()
                wb.save(output)
                st.session_state.final_excel_data = output.getvalue()
                st.balloons()

    if st.session_state.final_excel_data:
        st.download_button(
            label="💾 下载填充好的 Coupon 上传文件",
            data=st.session_state.final_excel_data,
            file_name=f"Cupshe_Coupon_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 第二阶段逻辑保持原样...
