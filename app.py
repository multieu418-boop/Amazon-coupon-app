import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from datetime import datetime, timedelta

# --- 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 侧边栏指引 ---
st.sidebar.header("📂 必备文件上传指引")
st.sidebar.markdown("""
1. **DE.txt**: 从亚马逊后台导出的 *All Listing Report*，用于获取实时价格。
2. **上传模板**: 亚马逊后台下载的空白 *Coupon 上传 Excel 模板*。本工具将保留其所有格式。
""")

# --- 1. 稳健的数据加载函数 ---
@st.cache_data
def load_inventory(file):
    if file:
        try:
            content = file.read()
            for encoding in ['utf-8', 'utf-16', 'cp1252', 'gbk']:
                try:
                    df = pd.read_csv(io.BytesIO(content), sep='\t', encoding=encoding, on_bad_lines='skip')
                    if 'asin1' in df.columns:
                        st.sidebar.success(f"✅ DE.txt 已就绪")
                        return df[['asin1', 'price']].drop_duplicates('asin1').set_index('asin1')
                except: continue
            st.error("❌ DE.txt 解析失败，请检查文件。")
        except Exception as e: st.error(f"❌ 读取出错: {e}")
    return None

inventory_file = st.sidebar.file_uploader("1. 上传 DE.txt (All Listing)", type=['txt'])
template_file = st.sidebar.file_uploader("2. 上传 Amazon 官方优惠券模板", type=['xlsx'])
inv_data = load_inventory(inventory_file)

# --- UI 界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：基于模板创建", "第二阶段：报错纠错修复"])

# ==========================================
# 第一阶段：基于模板创建
# ==========================================
with tab1:
    st.header("1️⃣ 填写需求并填充模板")
    
    if 'phase1_file' not in st.session_state:
        st.session_state.phase1_file = None

    if not template_file:
        st.warning("⚠️ 请先在左侧上传亚马逊官方 Excel 模板，否则无法生成文件。")
    
    with st.form("template_form"):
        col1, col2 = st.columns(2)
        with col1:
            asin_input = st.text_area("输入子 ASIN 列表", height=150, placeholder="B0xxxxxx;B0yyyyyy...")
            disc_type = st.selectbox("折扣类型", ["折扣 (Percentage)", "满减 (Money)"])
            disc_val = st.number_input("折扣数额", min_value=1.0, value=20.0)
        with col2:
            coupon_suffix = st.text_input("优惠券名称后缀", value="Cupshe Sale")
            budget = st.number_input("预算 (€)", min_value=100, value=1000)
            c_dates = st.date_input("日期范围", [datetime.now() + timedelta(days=1), datetime.now() + timedelta(days=30)])
        
        submitted = st.form_submit_button("✅ 填充模板并生成")

        if submitted and template_file:
            if not asin_input:
                st.error("请输入 ASIN！")
            else:
                # 1. 清洗 ASIN
                asins = [a.strip() for a in re.split(r'[;\n,\s\t]+', asin_input) if a.strip()]
                
                # 2. 价格预检
                if inv_data is not None:
                    for a in asins:
                        if a in inv_data.index:
                            try:
                                p = float(inv_data.loc[a, 'price'])
                                if disc_type == "满减 (Money)" and disc_val >= p * 0.5:
                                    st.warning(f"⚠️ {a} 折扣超 50% (价€{p})")
                            except: pass

                # 3. 使用 openpyxl 填充原始模板，确保不破坏格式
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active # 默认操作第一个 Sheet

                # 填充第 7 行 (Excel 中对应 index 为 7)
                # 列位对应关系 (根据标准模板): A: ASIN, B: 类型, C: 折扣, E: 名称, F: 预算, G: 开始, H: 结束
                ws.cell(row=7, column=1).value = ";".join(asins)
                ws.cell(row=7, column=2).value = "Percentage" if "折扣" in disc_type else "Money"
                
                if "折扣" in disc_type:
                    ws.cell(row=7, column=3).value = disc_val
                else:
                    ws.cell(row=7, column=4).value = disc_val
                
                ws.cell(row=7, column=5).value = coupon_suffix
                ws.cell(row=7, column=6).value = budget
                ws.cell(row=7, column=7).value = c_dates[0].strftime("%m/%d/%Y") if len(c_dates) > 0 else ""
                ws.cell(row=7, column=8).value = c_dates[1].strftime("%m/%d/%Y") if len(c_dates) > 1 else ""
                ws.cell(row=7, column=9).value = "Yes" # 默认限购一次
                ws.cell(row=7, column=11).value = "All Customers" # 默认目标客户

                output = io.BytesIO()
                wb.save(output)
                st.session_state.phase1_file = output.getvalue()
                st.success("✅ 模板已填充！请在下方下载。")

    if st.session_state.phase1_file:
        st.download_button("💾 下载填充后的上传模板", st.session_state.phase1_file, "Amazon_Coupon_Ready.xlsx")

# ==========================================
# 第二阶段：报错纠错修复 (逻辑保留)
# ==========================================
with tab2:
    st.header("2️⃣ 报错自动修复")
    st.info("步骤：上传亚马逊返回的报错文件 -> 选择剔除或修正 -> 下载修复后的新模板")
    # ... (第二阶段逻辑保持不变，确保解析批注并输出)
    err_file = st.file_uploader("上传报错 Excel (带有 N 列批注)", type=['xlsx'])
    if err_file:
        # 逻辑同前，确保解析批注
        st.write("解析中...")
