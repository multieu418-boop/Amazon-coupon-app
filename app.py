import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import datetime

# --- 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 初始化 Session State ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'field_configs' not in st.session_state:
    st.session_state.field_configs = []

# --- 侧边栏：文件上传与指引 ---
st.sidebar.header("📂 必备文件上传指引")
st.sidebar.markdown("""
1. **All Listing Report**: 获取价格校验（TXT格式）。
2. **上传Coupon文件模板**: 亚马逊下载的 Excel 模板。
   - **标题行**: 锁定第7行。
   - **下拉列**: 自动识别“折扣类型”、“限购”、“买家”等关键字并强制下拉。
   - **自由列**: 其他列全部支持自由手动输入。
""")

inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])

# --- 核心解析函数 ---
def parse_template_v3(file):
    if file is None: return []
    try:
        file.seek(0)
        wb = load_workbook(file, data_only=True)
        ws = wb.active
        configs = []
        
        # 扫描第7行标题
        for col in range(1, 25): # 扫描前24列，确保覆盖所有潜在字段
            title = ws.cell(row=7, column=col).value
            if title:
                title_str = str(title).strip()
                
                # 1. 提取第8, 9行的现有内容作为备选
                samples = []
                for r in [8, 9]:
                    v = ws.cell(row=r, column=col).value
                    if v:
                        s_v = str(v).strip()
                        if s_v not in samples: samples.append(s_v)
                
                # 2. 判断是否为“强制下拉列”
                is_dropdown = False
                dropdown_options = []
                
                # 匹配关键字
                if any(x in title_str for x in ["折扣类型", "Discount Type"]):
                    is_dropdown = True
                    dropdown_options = samples if samples else ["折扣", "满减", "Percentage", "Money"]
                elif any(x in title_str for x in ["兑换一次", "限购", "Limit"]):
                    is_dropdown = True
                    dropdown_options = samples if samples else ["是", "否", "Yes", "No"]
                elif any(x in title_str for x in ["优惠券类型", "Coupon Type"]):
                    is_dropdown = True
                    dropdown_options = samples if samples else ["Standard", "Subscribe & Save"]
                elif any(x in title_str for x in ["目标买家", "Target"]):
                    is_dropdown = True
                    dropdown_options = samples if samples else ["所有客户", "Amazon Prime 会员", "All Customers", "Amazon Prime Members"]
                elif any(x in title_str for x in ["叠加", "Stack"]):
                    is_dropdown = True
                    dropdown_options = samples if samples else ["是", "否", "Yes", "No"]
                
                configs.append({
                    "col": col,
                    "label": title_str,
                    "is_dropdown": is_dropdown,
                    "options": dropdown_options
                })
        return configs
    except Exception as e:
        st.error(f"解析出错: {e}")
        return []

if template_file and not st.session_state.field_configs:
    st.session_state.field_configs = parse_template_v3(template_file)

# --- 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")
tab1, tab2 = st.tabs(["第一阶段：动态表单录入", "第二阶段：报错修复"])

with tab1:
    if not template_file or not st.session_state.field_configs:
        st.warning("👋 请先上传模板，系统将自动识别第7行标题。")
    else:
        st.header("1️⃣ 录入优惠券需求")
        with st.form("coupon_form", clear_on_submit=True):
            user_data = {}
            cols = st.columns(2)
            
            for i, cfg in enumerate(st.session_state.field_configs):
                with cols[i % 2]:
                    label = cfg['label']
                    f_key = f"c_{cfg['col']}"
                    
                    # 如果是强制下拉列
                    if cfg['is_dropdown']:
                        user_data[cfg['col']] = st.selectbox(label, options=cfg['options'], key=f_key)
                    # 如果是日期相关的自由列
                    elif any(x in label for x in ["日期", "Date"]):
                        d_val = datetime.date.today() + datetime.timedelta(days=1)
                        user_data[cfg['col']] = st.date_input(label, value=d_val, key=f_key)
                    # 如果是ASIN等长文本自由列
                    elif any(x in label.upper() for x in ["ASIN", "列表"]):
                        user_data[cfg['col']] = st.text_area(label, key=f_key)
                    # 其他所有列全部设为自由输入文本框
                    else:
                        user_data[cfg['col']] = st.text_input(label, key=f_key)
            
            if st.form_submit_button("➕ 添加到需求池"):
                row = {}
                for c_idx, val in user_data.items():
                    if isinstance(val, (datetime.date, datetime.datetime)):
                        row[c_idx] = val.strftime("%m/%d/%Y")
                    else:
                        row[c_idx] = str(val) if val is not None else ""
                st.session_state.coupon_pool.append(row)
                st.toast("已记录一条需求")

        if st.session_state.coupon_pool:
            st.subheader("📋 预览待生成列表")
            name_map = {c['col']: c['label'] for c in st.session_state.field_configs}
            st.dataframe(pd.DataFrame(st.session_state.coupon_pool).rename(columns=name_map))
            
            if st.button("🚀 寻找空行并导出"):
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active
                
                # 从第8行开始往下找第一个A列为空的行
                start_r = 8
                while ws.cell(row=start_r, column=1).value is not None:
                    if not str(ws.cell(row=start_r, column=1).value).strip(): break
                    start_r += 1
                
                # 写入所有需求
                for idx, data in enumerate(st.session_state.coupon_pool):
                    curr_r = start_r + idx
                    for c_idx, val in data.items():
                        ws.cell(row=curr_r, column=int(c_idx)).value = val
                
                buf = io.BytesIO()
                wb.save(buf)
                st.download_button("💾 下载填充好的文件", buf.getvalue(), f"Coupon_Batch_{datetime.date.today()}.xlsx")
