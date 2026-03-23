import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import datetime

# --- 1. 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 2. 初始化 Session State (防止变量未定义报错) ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'field_configs' not in st.session_state:
    st.session_state.field_configs = []

# --- 3. 侧边栏：文件上传与操作指引 ---
st.sidebar.header("📂 必备文件上传指引")
st.sidebar.markdown("""
1. **All Listing Report**: 获取价格校验（TXT格式）。
2. **上传Coupon文件模板**: 亚马逊下载的 Excel 模板。
   - **标题行**: 锁定第7行。
   - **下拉列**: 自动识别并生成中文/英文下拉菜单。
   - **自由列**: 除特定列外，其余均可自由手动输入。
""")

inventory_file = st.sidebar.file_uploader("1. 上传 All Listing Report", type=['txt'])
template_file = st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx'])

# --- 4. 核心解析函数 (放在表单外部执行) ---
def parse_template_v4(file):
    if file is None: return []
    try:
        file.seek(0)
        wb = load_workbook(file, data_only=True)
        ws = wb.active
        configs = []
        
        # 扫描第7行标题（扫描前25列）
        for col in range(1, 26):
            title = ws.cell(row=7, column=col).value
            if title:
                title_str = str(title).strip()
                
                # 提取第8, 9行的现有内容作为下拉选项参考
                samples = []
                for r in [8, 9]:
                    v = ws.cell(row=r, column=col).value
                    if v:
                        s_v = str(v).strip()
                        if s_v not in samples: samples.append(s_v)
                
                # 定义需要强制下拉的关键词
                is_dropdown = False
                dropdown_options = []
                keywords = ["折扣类型", "兑换一次", "限购", "优惠券类型", "目标买家", "叠加", "Discount Type", "Limit", "Target"]
                
                if any(k in title_str for k in keywords):
                    is_dropdown = True
                    # 优先使用模板自带的选项，如果没有，给定中文标准选项
                    if samples:
                        dropdown_options = samples
                    else:
                        if "折扣类型" in title_str: dropdown_options = ["折扣", "满减", "Percentage", "Money"]
                        elif "限购" in title_str or "兑换" in title_str: dropdown_options = ["是", "否", "Yes", "No"]
                        else: dropdown_options = ["所有客户", "Amazon Prime 会员", "Standard"]

                configs.append({
                    "col": col,
                    "label": title_str,
                    "is_dropdown": is_dropdown,
                    "options": dropdown_options
                })
        return configs
    except Exception as e:
        st.error(f"模板解析失败: {e}")
        return []

# 当模板上传后，立即更新配置（确保在表单渲染前完成）
if template_file:
    # 只有当配置为空时才解析，避免重复刷新
    if not st.session_state.field_configs:
        st.session_state.field_configs = parse_template_v4(template_file)

# --- 5. 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")
tab1, tab2 = st.tabs(["第一阶段：动态表单录入", "第二阶段：报错修复"])

with tab1:
    if not template_file or not st.session_state.field_configs:
        st.warning("👋 请先上传模板，系统将自动识别第7行标题并生成输入框。")
    else:
        st.header("1️⃣ 录入优惠券需求")
        
        # --- 重点：整个表单逻辑必须在 with 缩进内完整闭合 ---
        with st.form("coupon_entry_form", clear_on_submit=True):
            st.write("请根据模板字段填写信息（特定列支持下拉，其余列自由填写）：")
            user_input_dict = {}
            # 自动分两列布局
            grid_cols = st.columns(2)
            
            for i, cfg in enumerate(st.session_state.field_configs):
                with grid_cols[i % 2]:
                    label = cfg['label']
                    unique_key = f"input_col_{cfg['col']}"
                    
                    if cfg['is_dropdown']:
                        # 强制下拉列
                        user_input_dict[cfg['col']] = st.selectbox(label, options=cfg['options'], key=unique_key)
                    elif any(x in label for x in ["日期", "Date"]):
                        # 日期自由列
                        tomorrow = datetime.date.today() + datetime.timedelta(days=1)
                        user_input_dict[cfg['col']] = st.date_input(label, value=tomorrow, key=unique_key)
                    elif any(x in label.upper() for x in ["ASIN", "列表"]):
                        # ASIN长文本自由列
                        user_input_dict[cfg['col']] = st.text_area(label, key=unique_key)
                    else:
                        # 普通文本自由列
                        user_input_dict[cfg['col']] = st.text_input(label, key=unique_key)
            
            # 必须包含的 Submit 按钮，且必须在 with st.form 缩进内
            add_to_list = st.form_submit_button("➕ 添加到待填充列表")
            
            if add_to_list:
                # 转换格式并存入 Session State
                row_data = {}
                for c_idx, val in user_input_dict.items():
                    if isinstance(val, (datetime.date, datetime.datetime)):
                        row_data[c_idx] = val.strftime("%m/%d/%Y")
                    else:
                        row_data[c_idx] = str(val) if val is not None else ""
                st.session_state.coupon_pool.append(row_data)
                st.toast("已记录一条需求，请在下方预览。")

        # --- 预览与文件生成 (在 Form 之外) ---
        if st.session_state.coupon_pool:
            st.divider()
            st.subheader("📋 待生成的需求预览")
            # 预览时映射标题
            mapping = {c['col']: c['label'] for c in st.session_state.field_configs}
            st.dataframe(pd.DataFrame(st.session_state.coupon_pool).rename(columns=mapping), use_container_width=True)
            
            btn_col1, btn_col2 = st.columns(2)
            if btn_col1.button("🗑️ 清空当前所有记录"):
                st.session_state.coupon_pool = []
                st.rerun()
            
            if btn_col2.button("🚀 寻找空行并导出 Excel"):
                try:
                    template_file.seek(0)
                    wb = load_workbook(template_file)
                    ws = wb.active
                    
                    # 寻找空行：从第8行起，直到第一列为空
                    start_r = 8
                    while ws.cell(row=start_r, column=1).value is not None:
                        if not str(ws.cell(row=start_r, column=1).value).strip():
                            break
                        start_r += 1
                    
                    # 写入数据
                    for offset, data in enumerate(st.session_state.coupon_pool):
                        write_row = start_row = start_r + offset
                        for col_idx, value in data.items():
                            ws.cell(row=write_row, column=int(col_idx)).value = value
                    
                    # 导出为二进制流
                    output_stream = io.BytesIO()
                    wb.save(output_stream)
                    st.success(f"✅ 填充成功！已避开示例，从第 {start_r} 行开始填充。")
                    
                    st.download_button(
                        label="💾 下载最终填充版 Excel",
                        data=output_stream.getvalue(),
                        file_name=f"Cupshe_Coupon_{datetime.date.today().strftime('%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"生成文件出错: {e}")
