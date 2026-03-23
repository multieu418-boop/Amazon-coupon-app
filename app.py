import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
import datetime

# --- 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 初始化 Session State ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'field_configs' not in st.session_state:
    st.session_state.field_configs = []
if 'final_xlsx_data' not in st.session_state:
    st.session_state.final_xlsx_data = None

# --- 侧边栏：文件上传与指引 ---
st.sidebar.header("📂 必备文件上传指引")
st.sidebar.markdown("""
1. **All Listing Report**: 用于校验（TXT格式）。
2. **上传Coupon文件模板**: 亚马逊下载的 Excel 模板。
   - **自动识别**: 程序读取第7行标题。
   - **智能下拉**: 自动提取第8、9行的**中文/英文**选项，不进行翻译。
""")

inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])

# --- 动态解析函数 ---
def parse_amazon_template(file):
    if file is None: return []
    try:
        file.seek(0)
        # data_only=True 确保读取显示值
        wb = load_workbook(file, data_only=True)
        ws = wb.active
        configs = []
        
        # 扫描第7行标题
        for col in range(1, 21):
            title = ws.cell(row=7, column=col).value
            if title:
                title_str = str(title).strip()
                
                # --- 核心改进：从第8、9行动态提取选项 ---
                # 不再写死英文，而是看模板里填了什么
                sample_options = []
                for r in [8, 9]:
                    val = ws.cell(row=r, column=col).value
                    if val:
                        s_val = str(val).strip()
                        if s_val not in sample_options:
                            sample_options.append(s_val)
                
                # 如果第8/9行有内容，且内容较短（通常下拉项不长），则视为下拉菜单
                # 或者标题包含关键术语时强制给定标准中文选项
                opts = None
                if len(sample_options) > 0 and all(len(x) < 20 for x in sample_options):
                    opts = sample_options
                
                # 补全常见的中文标准选项（防止模板是空的没有示例）
                if not opts:
                    if "折扣类型" in title_str: opts = ["折扣", "满减", "Percentage", "Money"]
                    elif "兑换一次" in title_str: opts = ["是", "否", "Yes", "No"]
                    elif "目标买家" in title_str: opts = ["所有客户", "Amazon Prime 会员", "All Customers", "Amazon Prime Members"]
                    elif "叠加" in title_str: opts = ["是", "否", "Yes", "No"]

                configs.append({
                    "col": col,
                    "label": title_str,
                    "options": opts
                })
        return configs
    except Exception as e:
        st.error(f"解析模板出错: {e}")
        return []

if template_file and not st.session_state.field_configs:
    st.session_state.field_configs = parse_amazon_template(template_file)

# --- 主界面逻辑 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

tab1, tab2 = st.tabs(["第一阶段：动态表单录入", "第二阶段：报错纠错修复"])

with tab1:
    if not template_file or not st.session_state.field_configs:
        st.warning("请上传『上传Coupon文件模板』以开始。")
    else:
        st.header("1️⃣ 录入优惠券需求")
        
        with st.form("coupon_input_form", clear_on_submit=True):
            user_responses = {}
            grid = st.columns(2)
            
            for idx, cfg in enumerate(st.session_state.field_configs):
                with grid[idx % 2]:
                    label = cfg['label']
                    f_id = f"col_{cfg['col']}"
                    
                    if cfg['options']:
                        # 直接展示模板自带的选项内容，不修改语言
                        user_responses[cfg['col']] = st.selectbox(label, options=cfg['options'], key=f_id)
                    elif any(x in label.upper() for x in ["ASIN", "列表"]):
                        user_responses[cfg['col']] = st.text_area(label, placeholder="分号分割", key=f_id)
                    elif any(x in label for x in ["日期", "Date"]):
                        default_d = datetime.date.today() + datetime.timedelta(days=1)
                        user_responses[cfg['col']] = st.date_input(label, value=default_d, key=f_id)
                    else:
                        user_responses[cfg['col']] = st.text_input(label, key=f_id)
            
            if st.form_submit_button("➕ 添加需求"):
                row = {}
                for c_idx, v in user_responses.items():
                    if isinstance(v, (datetime.date, datetime.datetime)):
                        row[c_idx] = v.strftime("%m/%d/%Y")
                    else:
                        row[c_idx] = str(v) if v is not None else ""
                st.session_state.coupon_pool.append(row)
                st.toast("已记录")

        if st.session_state.coupon_pool:
            st.subheader("📋 预览")
            name_map = {c['col']: c['label'] for c in st.session_state.field_configs}
            st.dataframe(pd.DataFrame(st.session_state.coupon_pool).rename(columns=name_map))
            
            if st.button("🚀 生成并导出"):
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active
                
                # 寻找空行
                start_r = 8
                while ws.cell(row=start_r, column=1).value is not None:
                    if not str(ws.cell(row=start_r, column=1).value).strip(): break
                    start_r += 1
                
                for i, data in enumerate(st.session_state.coupon_pool):
                    curr_r = start_r + i
                    for c_idx, val in data.items():
                        ws.cell(row=curr_r, column=int(c_idx)).value = val
                
                buf = io.BytesIO()
                wb.save(buf)
                st.session_state.final_xlsx_data = buf.getvalue()
                st.success(f"填充成功，起始行：{start_r}")

    if st.session_state.final_xlsx_data:
        st.download_button("💾 下载文件", st.session_state.final_xlsx_data, f"Coupon_{datetime.date.today()}.xlsx")
