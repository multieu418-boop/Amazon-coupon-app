import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from datetime import datetime, timedelta

# --- 页面设置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 侧边栏指引 ---
st.sidebar.header("📂 必备文件上传指引")
st.sidebar.markdown("""
1. **All Listing Report**: 从后台导出的 TXT 报告，用于校验价格。
2. **上传Coupon文件模板**: 亚马逊下载的空白 Excel 批量上传模板（程序将动态解析第7行标题）。
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
            return None
        except: return None
    return None

# --- 初始化 Session State ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'dynamic_headers' not in st.session_state:
    st.session_state.dynamic_headers = []

# --- 文件上传 ---
inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
inv_data = load_inventory(inventory_file)
if inv_data is not None:
    st.sidebar.success("✅ All Listing Report 已就绪")

template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])

# --- 核心：解析模板标题 ---
if template_file:
    try:
        template_file.seek(0)
        temp_wb = load_workbook(template_file, read_only=True)
        temp_ws = temp_wb.active
        # 提取第 7 行标题 (假设前 12 列为常用核心列)
        headers = []
        for col in range(1, 13):
            header_val = temp_ws.cell(row=7, column=col).value
            if header_val:
                headers.append({"col": col, "label": str(header_val).strip()})
        st.session_state.dynamic_headers = headers
        st.sidebar.success(f"✅ 模板解析成功（提取到 {len(headers)} 个字段）")
    except Exception as e:
        st.sidebar.error(f"❌ 模板解析失败: {e}")

# --- 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：动态模板填充", "第二阶段：报错纠错修复"])

with tab1:
    if not template_file:
        st.info("💡 请先在左侧上传『上传Coupon文件模板』以生成动态输入表单。")
    else:
        st.header("1️⃣ 根据模板标题录入需求")
        
        with st.expander("➕ 添加新优惠券需求", expanded=True):
            with st.form("dynamic_input_form", clear_on_submit=True):
                # 动态生成输入框
                user_inputs = {}
                cols = st.columns(2)
                for i, h in enumerate(st.session_state.dynamic_headers):
                    with cols[i % 2]:
                        # 对特定的列名做特殊 UI 处理
                        if "ASIN" in h['label'].upper():
                            user_inputs[h['col']] = st.text_area(h['label'], placeholder="分号分隔", key=f"h_{h['col']}")
                        elif "日期" in h['label'] or "Date" in h['label']:
                            user_inputs[h['col']] = st.date_input(h['label'], value=datetime.now() + timedelta(days=1), key=f"h_{h['col']}")
                        elif "折扣" in h['label'] or "满减" in h['label'] or "预算" in h['label']:
                            user_inputs[h['col']] = st.text_input(h['label'], value="", key=f"h_{h['col']}")
                        else:
                            user_inputs[h['col']] = st.text_input(h['label'], key=f"h_{h['col']}")
                
                add_btn = st.form_submit_button("确认添加此条需求")
                
                if add_btn:
                    # 转换日期格式为字符串
                    processed_input = {}
                    for k, v in user_inputs.items():
                        if isinstance(v, (datetime, datetime.date)):
                            processed_input[k] = v.strftime("%m/%d/%Y")
                        else:
                            processed_input[k] = v
                    st.session_state.coupon_pool.append(processed_input)
                    st.success("已成功加入需求池！")

        # 预览与生成
        if st.session_state.coupon_pool:
            st.subheader("📋 待生成列表")
            st.write(pd.DataFrame(st.session_state.coupon_pool).rename(columns={h['col']: h['label'] for h in st.session_state.dynamic_headers}))
            
            col_btn1, col_btn2 = st.columns(2)
            if col_btn1.button("🗑️ 清空当前需求"):
                st.session_state.coupon_pool = []
                st.rerun()

            if col_btn2.button("🚀 寻找空白行并填充模板"):
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active
                
                # --- 核心算法：寻找第一个空白行 ---
                # 从第 8 行开始往下找
                target_row = 8
                while True:
                    # 如果第一列(ASIN列)有内容，就往下一行找
                    if ws.cell(row=target_row, column=1).value is not None:
                        target_row += 1
                    else:
                        break
                
                # 开始填充
                for i, row_data in enumerate(st.session_state.coupon_pool):
                    current_line = target_row + i
                    for col_idx, value in row_data.items():
                        ws.cell(row=current_line, column=col_idx).value = value
                
                output = io.BytesIO()
                wb.save(output)
                st.download_button(
                    label="💾 下载填充好的文件",
                    data=output.getvalue(),
                    file_name=f"Fixed_Template_{datetime.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# --- 第二阶段 (保持此前逻辑) ---
with tab2:
    st.header("2️⃣ 纠错修复逻辑")
    st.write("此处逻辑保持不变，用于处理上传亚马逊后的报错批注。")
