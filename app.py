import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from openpyxl.utils.dataframe import DataFrame_to_rows
from datetime import datetime, timedelta

# --- 页面设置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 侧边栏指引 ---
st.sidebar.header("📂 必备文件上传指引")
st.sidebar.markdown("""
1. **All Listing Report**: 从亚马逊后台导出的 *All Listing Report* (txt格式)，用于获取实时价格。
2. **上传Coupon文件模板**: 从亚马逊后台下载的 *空白优惠券上传 Excel 模板*。
""")

# --- 1. 数据加载函数 ---
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
        except Exception: return "ERROR_READ"
    return None

# --- 初始化 Session State ---
if 'coupon_list' not in st.session_state:
    st.session_state.coupon_list = []
if 'phase1_final_file' not in st.session_state:
    st.session_state.phase1_final_file = None

# --- 侧边栏文件上传状态 ---
inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
inv_data = load_inventory(inventory_file)
if inv_data is not None:
    if isinstance(inv_data, str): st.sidebar.error(f"❌ All Listing 解析失败")
    else: st.sidebar.success(f"✅ All Listing Report 已就绪")

template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])
if template_file:
    st.sidebar.success(f"✅ 上传Coupon文件模板 已就绪")

# --- UI 界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：多需求批量填充", "第二阶段：报错纠错修复"])

# ==========================================
# 第一阶段：多需求批量填充
# ==========================================
with tab1:
    st.header("1️⃣ 录入优惠券需求")
    
    with st.expander("➕ 添加新的优惠券需求", expanded=True):
        with st.form("single_coupon_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                asin_input = st.text_area("输入子 ASIN 列表 (分号/换行分隔)", height=150)
                disc_type = st.selectbox("折扣类型", ["折扣 (Percentage)", "满减 (Money)"])
                disc_val = st.number_input("折扣数额", min_value=1.0, value=20.0)
            with col2:
                coupon_suffix = st.text_input("优惠券名称后缀", value="Summer Sale")
                budget = st.number_input("预算 (€)", min_value=100, value=1000)
                c_dates = st.date_input("日期范围", [datetime.now() + timedelta(days=1), datetime.now() + timedelta(days=30)])
            
            add_btn = st.form_submit_button("添加至需求列表")
            
            if add_btn:
                if not asin_input:
                    st.error("请输入 ASIN！")
                else:
                    asins = [a.strip() for a in re.split(r'[;\n,\s\t]+', asin_input) if a.strip()]
                    new_req = {
                        "asins": ";".join(asins),
                        "type": "Percentage" if "折扣" in disc_type else "Money",
                        "val": disc_val,
                        "name": coupon_suffix,
                        "budget": budget,
                        "start": c_dates[0].strftime("%m/%d/%Y") if len(c_dates) > 0 else "",
                        "end": c_dates[1].strftime("%m/%d/%Y") if len(c_dates) > 1 else ""
                    }
                    st.session_state.coupon_list.append(new_req)
                    st.toast("已添加一项需求！")

    # 展示已添加的需求列表
    if st.session_state.coupon_list:
        st.subheader("📋 待生成的优惠券列表")
        display_df = pd.DataFrame(st.session_state.coupon_list)
        st.table(display_df)
        
        if st.button("🗑️ 清空列表"):
            st.session_state.coupon_list = []
            st.rerun()

        # 生成逻辑
        if st.button("🚀 填充至模板并生成最终文件"):
            if not template_file:
                st.error("请先在左侧上传『上传Coupon文件模板』！")
            else:
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active # 操作第一个 Sheet
                
                # 从第 7 行开始填充
                start_row = 7
                for i, req in enumerate(st.session_state.coupon_list):
                    curr_row = start_row + i
                    ws.cell(row=curr_row, column=1).value = req["asins"]
                    ws.cell(row=curr_row, column=2).value = req["type"]
                    # 区分百分比列和金额列
                    if req["type"] == "Percentage":
                        ws.cell(row=curr_row, column=3).value = req["val"]
                    else:
                        ws.cell(row=curr_row, column=4).value = req["val"]
                    
                    ws.cell(row=curr_row, column=5).value = req["name"]
                    ws.cell(row=curr_row, column=6).value = req["budget"]
                    ws.cell(row=curr_row, column=7).value = req["start"]
                    ws.cell(row=curr_row, column=8).value = req["end"]
                    ws.cell(row=curr_row, column=9).value = "Yes"
                    ws.cell(row=curr_row, column=11).value = "All Customers"

                output = io.BytesIO()
                wb.save(output)
                st.session_state.phase1_final_file = output.getvalue()
                st.success("✅ 所有需求已成功填入模板！")

    if st.session_state.phase1_final_file:
        st.download_button(
            label="💾 下载最终生成的 Coupon 上传文件",
            data=st.session_state.phase1_final_file,
            file_name=f"Final_Coupon_Upload_{datetime.now().strftime('%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==========================================
# 第二阶段：纠错逻辑 (保持并优化)
# ==========================================
with tab2:
    st.header("2️⃣ 报错自动修复")
    st.write("上传亚马逊报错反馈文件，快速剔除不合格 ASIN。")
    # ... 此处逻辑同前，确保解析 N 列批注 ...
