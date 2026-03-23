import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from datetime import datetime, timedelta

# --- 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

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
            return None
        except: return None
    return None

# --- 初始化 Session State ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'dynamic_config' not in st.session_state:
    st.session_state.dynamic_config = []

# --- 侧边栏 ---
st.sidebar.header("📂 必备文件上传指引")
inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
inv_data = load_inventory(inventory_file)
if inv_data is not None:
    st.sidebar.success("✅ All Listing Report 已就绪")

template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])

# --- 核心：解析模板字段与下拉选项 ---
if template_file:
    try:
        template_file.seek(0)
        temp_wb = load_workbook(template_file, data_only=True)
        temp_ws = temp_wb.active
        
        configs = []
        # 读取第 7 行标题，并参考第 8、9 行获取“预设选项”
        for col in range(1, 15): # 扫描前 14 列
            header = temp_ws.cell(row=7, column=col).value
            if header:
                header_str = str(header).strip()
                # 尝试获取第 8 或 9 行的值作为参考
                ref_val_8 = temp_ws.cell(row=8, column=col).value
                ref_val_9 = temp_ws.cell(row=9, column=col).value
                
                # 逻辑：如果参考行里有固定值且不是 ASIN 这种长字符串，我们视其为潜在选项
                options = []
                # 常见的亚马逊固定选项手动加固
                if "折扣类型" in header_str: options = ["Percentage", "Money"]
                elif "限购" in header_str: options = ["Yes", "No"]
                elif "优惠券类型" in header_str: options = ["Standard", "Subscribe & Save"]
                elif "目标买家" in header_str: options = ["All Customers", "Amazon Prime Members"]
                elif "叠加" in header_str: options = ["Yes", "No"]
                
                configs.append({
                    "col": col,
                    "label": header_str,
                    "options": options if options else None
                })
        st.session_state.dynamic_config = configs
        st.sidebar.success(f"✅ 模板解析成功")
    except Exception as e:
        st.sidebar.error(f"❌ 模板解析失败: {e}")

# --- 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：动态模板填充", "第二阶段：纠错修复"])

with tab1:
    if not template_file:
        st.info("💡 请先在左侧上传『上传Coupon文件模板』以激活表单。")
    else:
        st.header("1️⃣ 录入优惠券需求")
        with st.expander("➕ 添加新优惠券（支持多项累加）", expanded=True):
            with st.form("dynamic_form", clear_on_submit=True):
                user_data = {}
                cols = st.columns(2)
                for i, conf in enumerate(st.session_state.dynamic_config):
                    with cols[i % 2]:
                        # 根据是否有预设选项生成 Selectbox 或 Text
                        if conf['options']:
                            user_data[conf['col']] = st.selectbox(conf['label'], options=conf['options'], key=f"c_{conf['col']}")
                        elif "ASIN" in conf['label'].upper():
                            user_data[conf['col']] = st.text_area(conf['label'], placeholder="分号分隔", key=f"c_{conf['col']}")
                        elif "日期" in conf['label'] or "Date" in conf['label']:
                            user_data[conf['col']] = st.date_input(conf['label'], key=f"c_{conf['col']}")
                        else:
                            user_data[conf['col']] = st.text_input(conf['label'], key=f"c_{conf['col']}")
                
                if st.form_submit_button("确认并添加至列表"):
                    # 转换数据格式
                    formatted = {}
                    for k, v in user_data.items():
                        if isinstance(v, (datetime, datetime.date)):
                            formatted[k] = v.strftime("%m/%d/%Y")
                        else:
                            formatted[k] = v
                    st.session_state.coupon_pool.append(formatted)
                    st.toast("已成功添加一条需求！")

        if st.session_state.coupon_pool:
            st.subheader("📋 待填充列表")
            st.write(pd.DataFrame(st.session_state.coupon_pool).rename(columns={c['col']: c['label'] for c in st.session_state.dynamic_config}))
            
            c1, c2 = st.columns(2)
            if c1.button("🗑️ 清空所有需求"):
                st.session_state.coupon_pool = []
                st.rerun()
            
            if c2.button("🚀 寻找空白行并填充至模板"):
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active
                
                # --- 智能寻找空白行 ---
                # 从第 8 行起找，必须满足 A 列（ASIN列）为空
                current_target_row = 8
                while ws.cell(row=current_target_row, column=1).value is not None:
                    current_target_row += 1
                
                # 开始填充
                for i, row_item in enumerate(st.session_state.coupon_pool):
                    write_row = current_target_row + i
                    for col_idx, val in row_item.items():
                        ws.cell(row=write_row, column=col_idx).value = val
                
                output = io.BytesIO()
                wb.save(output)
                st.session_state.phase1_final = output.getvalue()
                st.success(f"✅ 成功从第 {current_target_row} 行开始填充！")

    if st.session_state.get('phase1_final'):
        st.download_button("💾 下载填充好的文件", st.session_state.phase1_final, f"Coupon_Upload_{datetime.now().strftime('%m%d')}.xlsx")

# --- 第二阶段 (解析 N 列批注逻辑保留) ---
with tab2:
    st.header("2️⃣ 报错纠错与自动重做")
    st.info("上传带批注的报错文件，系统将自动从 N 列提取报错 ASIN 并协助你完成修正。")
