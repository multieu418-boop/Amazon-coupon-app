import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from datetime import datetime, timedelta

# --- 页面设置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 修复：更稳健的 DE.txt 加载函数 ---
@st.cache_data
def load_inventory(file):
    if file:
        try:
            # 尝试多种编码格式，解决 Amazon 导出的 TXT 乱码问题
            content = file.read()
            # 尝试常用编码
            for encoding in ['utf-8', 'utf-16', 'cp1252', 'gbk']:
                try:
                    df = pd.read_csv(io.BytesIO(content), sep='\t', encoding=encoding, on_bad_lines='skip')
                    if 'asin1' in df.columns:
                        st.sidebar.success(f"✅ 成功读取 (编码: {encoding})")
                        return df[['asin1', 'price']].drop_duplicates('asin1').set_index('asin1')
                except:
                    continue
            st.error("❌ 无法解析文件编码，请确保 DE.txt 是从亚马逊导出的制表符分隔文件。")
        except Exception as e:
            st.error(f"❌ 读取出错: {e}")
    return None

# --- UI 界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

st.sidebar.header("⚙️ 核心数据配置")
inventory_file = st.sidebar.file_uploader("第一步：上传 DE.txt", type=['txt'])
inv_data = load_inventory(inventory_file)

tab1, tab2 = st.tabs(["第一阶段：创建新券", "第二阶段：纠错修复"])

# --- 第一阶段逻辑 ---
with tab1:
    st.header("1️⃣ 初始创建优惠券模板")
    with st.form("create_form"):
        col1, col2 = st.columns(2)
        with col1:
            asin_input = st.text_area("输入子 ASIN 列表", height=150, placeholder="B0xxxxxx;B0yyyyyy...")
            disc_type = st.selectbox("折扣类型", ["折扣 (Percentage)", "满减 (Money)"])
            disc_val = st.number_input("折扣数额", min_value=1.0, value=20.0)
        with col2:
            coupon_suffix = st.text_input("优惠券名称后缀", value="Cupshe Sale")
            budget = st.number_input("预算 (€)", min_value=100, value=1000)
            c_dates = st.date_input("日期范围", [datetime.now() + timedelta(days=1), datetime.now() + timedelta(days=30)])
        
        submitted = st.form_submit_button("生成上传模板")
        
        if submitted:
            if not asin_input:
                st.error("请输入 ASIN！")
            else:
                # 清洗 ASIN，支持空格、分号、换行
                asins = [a.strip() for a in re.split(r'[;\n,\s\t]+', asin_input) if a.strip()]
                
                # 价格预检 (增加空值检查)
                if inv_data is not None:
                    invalid_asins = []
                    for a in asins:
                        if a in inv_data.index:
                            try:
                                p = float(inv_data.loc[a, 'price'])
                                if disc_type == "满减 (Money)" and disc_val >= p * 0.5:
                                    invalid_asins.append(a)
                            except: continue
                    if invalid_asins:
                        st.warning(f"⚠️ 预警：以下 ASIN 折扣超50%: {', '.join(invalid_asins)}")

                # 生成 DataFrame
                df_out = pd.DataFrame([{
                    "ASIN 列表": ";".join(asins),
                    "折扣类型": "Percentage" if "折扣" in disc_type else "Money",
                    "折扣数额": disc_val,
                    "名称": coupon_suffix,
                    "预算": budget,
                    "开始日期": c_dates[0].strftime("%m/%d/%Y"),
                    "结束日期": c_dates[1].strftime("%m/%d/%Y") if len(c_dates)>1 else ""
                }])
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_out.to_excel(writer, index=False)
                st.download_button("💾 下载第一阶段上传文件", output.getvalue(), "Coupon_Phase1.xlsx")

# --- 第二阶段逻辑 ---
with tab2:
    st.header("2️⃣ 报错纠错与自动重做")
    err_file = st.file_uploader("上传带有批注的报错 Excel", type=['xlsx'])
    if err_file:
        wb = load_workbook(err_file)
        ws = wb.active
        error_records = []
        for row_idx in range(8, ws.max_row + 1):
            n_cell = ws.cell(row=row_idx, column=14)
            if n_cell.comment or (n_cell.value and "验证" in str(n_cell.value)):
                comment_text = n_cell.comment.text if n_cell.comment else str(n_cell.value)
                asin_str = str(ws.cell(row=row_idx, column=1).value)
                bug_asins = re.findall(r'[A-Z0-9]{10}', comment_text)
                if bug_asins:
                    error_records.append({
                        "row": row_idx,
                        "coupon_name": ws.cell(row=row_idx, column=5).value,
                        "bug_asins": list(set(bug_asins)),
                        "all_asins": asin_str.split(';'),
                        "orig_data": [ws.cell(row=row_idx, column=c).value for c in range(1, 15)]
                    })
        
        if error_records:
            final_rows = []
            for rec in error_records:
                st.subheader(f"修复券: {rec['coupon_name']}")
                choice = st.radio(f"策略 ({rec['row']})", ["全部剔除报错ASIN", "修改力度重做"], key=f"r_{rec['row']}")
                
                clean = [a for a in rec['all_asins'] if a not in rec['bug_asins']]
                
                if choice == "全部剔除报错ASIN":
                    row_data = list(rec['orig_data'])
                    row_data[0] = ";".join(clean)
                    final_rows.append(row_data)
                else:
                    new_v = st.number_input("新力度", value=5.0, key=f"v_{rec['row']}")
                    r1 = list(rec['orig_data']); r1[0] = ";".join(clean)
                    r2 = list(rec['orig_data']); r2[0] = ";".join(rec['bug_asins']); r2[2] = new_v
                    final_rows.extend([r1, r2])
            
            if st.button("生成纠错后的文件"):
                out_fix = io.BytesIO()
                pd.DataFrame(final_rows).to_excel(out_fix, index=False, header=False)
                st.download_button("💾 下载修复版", out_fix.getvalue(), "Fixed_Coupons.xlsx")
