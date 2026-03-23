import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from datetime import datetime, date

# --- 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 1. 侧边栏：文件上传与指引 ---
st.sidebar.header("📂 必备文件上传指引")
st.sidebar.info("""
1. **All Listing Report**: 获取价格校验（TXT格式）。
2. **上传Coupon文件模板**: 亚马逊下载的空白 Excel 模板。
   - 程序将自动读取第7行作为表单标题。
   - 程序将自动识别第8、9行的下拉选项。
""")

inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])

# --- 2. 初始化 Session State ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'field_configs' not in st.session_state:
    st.session_state.field_configs = []

# --- 3. 动态解析函数 ---
def parse_amazon_template(file):
    try:
        file.seek(0)
        # data_only=True 确保读取的是值而不是公式
        wb = load_workbook(file, data_only=True)
        ws = wb.active
        configs = []
        
        # 扫描第7行标题（通常1-15列覆盖所有Coupon字段）
        for col in range(1, 16):
            title = ws.cell(row=7, column=col).value
            if title:
                title_str = str(title).strip()
                
                # 预设下拉选项逻辑
                opts = None
                # 方案A：根据标题关键字匹配（最稳妥）
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
    except:
        return []

# 当文件上传时，立即更新配置
if template_file:
    st.session_state.field_configs = parse_amazon_template(template_file)
    if st.session_state.field_configs:
        st.sidebar.success(f"✅ 成功解析 {len(st.session_state.field_configs)} 个模板字段")

# --- 4. 主界面逻辑 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：动态表单录入", "第二阶段：报错纠错修复"])

with tab1:
    if not template_file or not st.session_state.field_configs:
        st.warning("请先在左侧上传『上传Coupon文件模板』以生成动态输入界面。")
    else:
        st.header("1️⃣ 录入优惠券需求")
        
        # 使用 Form 包裹，确保 Submit 按钮逻辑完整
        with st.form("dynamic_coupon_form", clear_on_submit=True):
            st.write("请根据模板要求填写以下信息：")
            user_responses = {}
            # 自动分两列显示，节省空间
            grid = st.columns(2)
            
            for idx, cfg in enumerate(st.session_state.field_configs):
                with grid[idx % 2]:
                    label = cfg['label']
                    f_key = f"input_{cfg['col']}"
                    
                    if cfg['options']:
                        # 生成下拉选择框
                        user_responses[cfg['col']] = st.selectbox(label, options=cfg['options'], key=f_key)
                    elif any(x in label for x in ["ASIN", "列表"]):
                        # 生成多行输入框
                        user_responses[cfg['col']] = st.text_area(label, placeholder="分号分割ASIN", key=f_key)
                    elif any(x in label for x in ["日期", "Date"]):
                        # 生成日期选择器
                        user_responses[cfg['col']] = st.date_input(label, value=date.today() + timedelta(days=1), key=f_key)
                    else:
                        # 生成普通文本框
                        user_responses[cfg['col']] = st.text_input(label, key=f_key)
            
            # 必须包含的提交按钮
            add_trigger = st.form_submit_button("➕ 添加到待填充列表")
            
            if add_trigger:
                # 转换数据类型，统一转为字符串存入 Pool，防止写入 Excel 报错
                formatted_entry = {}
                for c_idx, val in user_responses.items():
                    if isinstance(val, (date, datetime)):
                        formatted_entry[c_idx] = val.strftime("%m/%d/%Y")
                    else:
                        formatted_entry[c_idx] = str(val) if val is not None else ""
                
                st.session_state.coupon_pool.append(formatted_entry)
                st.toast("已成功记录需求！")

        # --- 预览与最终生成 ---
        if st.session_state.coupon_pool:
            st.subheader("📋 待生成的 Coupon 列表")
            # 预览时映射标题名
            name_map = {c['col']: c['label'] for c in st.session_state.field_configs}
            st.dataframe(pd.DataFrame(st.session_state.coupon_pool).rename(columns=name_map))
            
            c1, c2 = st.columns(2)
            if c1.button("🗑️ 清空所有记录"):
                st.session_state.coupon_pool = []
                st.rerun()
            
            if c2.button("🚀 填充模板并下载文件"):
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active # 保持原有 Sheet 不变
                
                # --- 智能寻找起始空行 ---
                # 从第8行开始往下找，直到 A 列（Column 1）为空
                start_row = 8
                while True:
                    cell_val = ws.cell(row=start_row, column=1).value
                    # 如果有内容，或者内容是空格，则视为已被占用
                    if cell_val is not None and len(str(cell_val).strip()) > 0:
                        start_row += 1
                    else:
                        break
                
                # 批量顺序写入
                for i, data_dict in enumerate(st.session_state.coupon_pool):
                    target_r = start_row + i
                    for col_idx, value in data_dict.items():
                        ws.cell(row=target_r, column=int(col_idx)).value = value
                
                # 导出
                out_buffer = io.BytesIO()
                wb.save(out_buffer)
                
                st.success(f"✅ 数据已填充！已跳过示例行，从第 {start_row} 行开始写入。")
                st.download_button(
                    label="💾 点击下载最终上传模板",
                    data=out_buffer.getvalue(),
                    file_name=f"Coupon_Upload_{date.today().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# --- 第二阶段逻辑保持不变 ---
