import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import datetime

# --- 1. 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 2. 初始化 Session State ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'field_configs' not in st.session_state:
    st.session_state.field_configs = []

# --- 3. 侧边栏 ---
st.sidebar.header("📂 文件上传与操作指引")
st.sidebar.markdown("""
1. **上传模板**: 程序将读取**第7行**作为标题，读取**第5行**作为填写规则。
2. **填写校验**: ASIN 列会自动检查数量（15-2900个）及分隔符。
3. **写入逻辑**: 自动避开第8、9行示例，从首个空行开始。
""")

inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])

# --- 4. 增强版解析函数 (提取第5行规则) ---
def parse_template_v5(file):
    if file is None: return []
    try:
        file.seek(0)
        wb = load_workbook(file, data_only=True)
        ws = wb.active
        configs = []
        
        # 扫描标题行（第7行）
        for col in range(1, 26):
            title = ws.cell(row=7, column=col).value
            rule_hint = ws.cell(row=5, column=col).value  # 提取第5行的规则说明
            
            if title:
                title_str = str(title).strip()
                hint_str = str(rule_hint).strip() if rule_hint else "请按照亚马逊要求填写"
                
                # 提取第8, 9行下拉示例
                samples = []
                for r in [8, 9]:
                    v = ws.cell(row=r, column=col).value
                    if v:
                        s_v = str(v).strip()
                        if s_v not in samples: samples.append(s_v)
                
                # 下拉判定逻辑
                is_dropdown = False
                dropdown_options = []
                keywords = ["折扣类型", "兑换一次", "限购", "优惠券类型", "目标买家", "叠加", "Discount Type", "Limit"]
                
                if any(k in title_str for k in keywords):
                    is_dropdown = True
                    dropdown_options = samples if samples else ["折扣", "满减", "是", "否"]

                configs.append({
                    "col": col,
                    "label": title_str,
                    "hint": hint_str,
                    "is_dropdown": is_dropdown,
                    "options": dropdown_options
                })
        return configs
    except Exception as e:
        st.error(f"解析模板失败: {e}")
        return []

if template_file and not st.session_state.field_configs:
    st.session_state.field_configs = parse_template_v5(template_file)

# --- 5. 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")
tab1, tab2 = st.tabs(["第一阶段：动态表单录入", "第二阶段：报错修复"])

with tab1:
    if not template_file or not st.session_state.field_configs:
        st.warning("👋 请先上传模板，系统将自动分析第5行规则与第7行标题。")
    else:
        st.header("1️⃣ 录入优惠券需求")
        
        with st.form("coupon_entry_form", clear_on_submit=True):
            user_input_dict = {}
            grid_cols = st.columns(2)
            
            for i, cfg in enumerate(st.session_state.field_configs):
                with grid_cols[i % 2]:
                    label = cfg['label']
                    hint = cfg['hint']
                    f_key = f"input_col_{cfg['col']}"
                    
                    # 1. 下拉列
                    if cfg['is_dropdown']:
                        user_input_dict[cfg['col']] = st.selectbox(label, options=cfg['options'], help=hint, key=f_key)
                    
                    # 2. ASIN 列 (特殊校验)
                    elif any(x in label.upper() for x in ["ASIN", "列表"]):
                        val = st.text_area(label, help=hint, placeholder="示例: ASIN1;ASIN2;ASIN3", key=f_key)
                        user_input_dict[cfg['col']] = val
                        
                        # 实时校验逻辑
                        if val:
                            asin_list = re.split(r'[;,\s]+', val.strip())
                            asin_list = [a for a in asin_list if a]
                            count = len(asin_list)
                            if count < 15:
                                st.error(f"⚠️ 当前 ASIN 数量: {count}。注意: 亚马逊要求最少 15 个。")
                            elif count > 2900:
                                st.error(f"⚠️ 当前 ASIN 数量: {count}。注意: 亚马逊要求最多 2900 个。")
                            else:
                                st.success(f"✅ 已输入 {count} 个 ASIN，符合数量要求。")
                    
                    # 3. 日期列
                    elif any(x in label for x in ["日期", "Date"]):
                        tomorrow = datetime.date.today() + datetime.timedelta(days=1)
                        user_input_dict[cfg['col']] = st.date_input(label, value=tomorrow, help=hint, key=f_key)
                    
                    # 4. 其他自由列
                    else:
                        user_input_dict[cfg['col']] = st.text_input(label, help=hint, key=f_key)
            
            add_to_list = st.form_submit_button("➕ 添加到待填充列表")
            
            if add_to_list:
                row_data = {}
                for c_idx, val in user_input_dict.items():
                    if isinstance(val, (datetime.date, datetime.datetime)):
                        row_data[c_idx] = val.strftime("%m/%d/%Y")
                    else:
                        row_data[c_idx] = str(val) if val is not None else ""
                st.session_state.coupon_pool.append(row_data)
                st.toast("需求已保存！")

        # --- 预览与生成 ---
        if st.session_state.coupon_pool:
            st.divider()
            st.subheader("📋 待处理需求预览")
            mapping = {c['col']: c['label'] for c in st.session_state.field_configs}
            st.dataframe(pd.DataFrame(st.session_state.coupon_pool).rename(columns=mapping), use_container_width=True)
            
            b_col1, b_col2 = st.columns(2)
            if b_col1.button("🗑️ 清空当前记录"):
                st.session_state.coupon_pool = []
                st.rerun()
            
            if b_col2.button("🚀 寻找空行并导出 Excel"):
                template_file.seek(0)
                wb = load_workbook(template_file)
                ws = wb.active
                
                start_r = 8
                while ws.cell(row=start_r, column=1).value is not None:
                    if not str(ws.cell(row=start_r, column=1).value).strip(): break
                    start_r += 1
                
                for offset, data in enumerate(st.session_state.coupon_pool):
                    write_row = start_r + offset
                    for col_idx, value in data.items():
                        ws.cell(row=write_row, column=int(col_idx)).value = value
                
                output_stream = io.BytesIO()
                wb.save(output_stream)
                st.download_button("💾 下载填充好的 Excel", output_stream.getvalue(), f"Cupshe_Coupon_{datetime.date.today()}.xlsx")
