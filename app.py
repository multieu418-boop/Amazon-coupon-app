import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
import datetime

# --- 1. 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 2. 状态初始化 ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'field_configs' not in st.session_state:
    st.session_state.field_configs = []
if 'inventory_db' not in st.session_state:
    st.session_state.inventory_db = None

# --- 3. 核心工具函数 ---

def clean_asin_format(raw_text):
    """强制转换任何格式的 ASIN 为分号分隔的规范字符串"""
    if not raw_text: return "", 0
    tokens = re.split(r'[;；,，\s\n\r]+', str(raw_text).strip())
    clean_list = [t.strip().upper() for t in tokens if t.strip()]
    # 去重保持顺序
    seen = set()
    final_list = [x for x in clean_list if not (x in seen or seen.add(x))]
    return ";".join(final_list), len(final_list)

def load_inventory(file):
    """加载 All Listing Report 进行价格和在售校验"""
    try:
        content = file.read()
        # 尝试多种编码读取 TXT
        for enc in ['utf-8', 'utf-16', 'gbk', 'cp1252']:
            try:
                df = pd.read_csv(io.BytesIO(content), sep='\t', encoding=enc, on_bad_lines='skip')
                if 'asin1' in df.columns:
                    # 建立索引提高查询速度
                    return df[['asin1', 'seller-sku', 'price', 'quantity']].set_index('asin1')
            except: continue
        return None
    except: return None

# --- 4. 侧边栏：多文件上传流 ---
st.sidebar.header("📁 数据中心")

# 1. 基础库存导入
inv_file = st.sidebar.file_uploader("1. 导入 All Listing Report (TXT)", type=['txt'])
if inv_file and st.session_state.inventory_db is None:
    st.session_state.inventory_db = load_inventory(inv_file)
    if st.session_state.inventory_db is not None:
        st.sidebar.success("✅ 库存数据已挂载")

# 2. 模板导入
template_file = st.sidebar.file_uploader("2. 上传 Coupon 原始模板 (Excel)", type=['xlsx'])

# 3. 纠错流导入
error_archive = st.sidebar.file_uploader("3. (可选) 上传亚马逊报错的文件", type=['xlsx'])

# 解析模板字段
if template_file and not st.session_state.field_configs:
    try:
        template_file.seek(0)
        wb = load_workbook(template_file, data_only=True)
        ws = wb.active
        configs = []
        for col in range(1, 26):
            title = ws.cell(row=7, column=col).value
            hint = ws.cell(row=5, column=col).value
            if title:
                samples = [str(ws.cell(row=r, column=col).value).strip() for r in [8,9] if ws.cell(row=r, column=col).value]
                is_drop = any(k in str(title) for k in ["折扣类型", "兑换一次", "限购", "优惠券类型", "目标买家", "叠加"])
                configs.append({
                    "col": col, "label": str(title).strip(), "hint": str(hint).strip() if hint else "",
                    "is_dropdown": is_drop, "options": samples if samples else ["是", "否"]
                })
        st.session_state.field_configs = configs
    except: st.sidebar.error("模板解析失败")

# --- 5. 主界面三阶段流 ---
st.title("👗 Cupshe 亚马逊优惠助手 (全功能版)")

tab1, tab2, tab3 = st.tabs(["✨ 第一阶段：创建流", "🔍 第二阶段：校验流", "🛠️ 第三阶段：纠错流"])

# --- TAB 1: 创建流 ---
with tab1:
    if not st.session_state.field_configs:
        st.info("请先在左侧上传 Coupon 模板以激活创建表单。")
    else:
        with st.form("main_entry_form", clear_on_submit=True):
            st.subheader("📝 录入新需求")
            user_data = {}
            grid = st.columns(2)
            for i, cfg in enumerate(st.session_state.field_configs):
                with grid[i % 2]:
                    if cfg['is_dropdown']:
                        user_data[cfg['col']] = st.selectbox(cfg['label'], options=cfg['options'], help=cfg['hint'], key=f"in_{cfg['col']}")
                    elif "ASIN" in cfg['label'].upper():
                        user_data[cfg['col']] = st.text_area(cfg['label'], help=cfg['hint'], placeholder="粘贴ASIN，自动转为分号分隔", key=f"in_{cfg['col']}")
                    elif "日期" in cfg['label'] or "Date" in cfg['label']:
                        user_data[cfg['col']] = st.date_input(cfg['label'], value=datetime.date.today()+datetime.timedelta(days=1), key=f"in_{cfg['col']}")
                    else:
                        user_data[cfg['col']] = st.text_input(cfg['label'], help=cfg['hint'], key=f"in_{cfg['col']}")
            
            if st.form_submit_button("➕ 确认并添加"):
                processed_row = {}
                for c_idx, val in user_data.items():
                    # 识别是否为 ASIN 列进行强制格式化
                    lbl = next(c['label'] for c in st.session_state.field_configs if c['col'] == c_idx)
                    if "ASIN" in lbl.upper():
                        clean_str, _ = clean_asin_format(val)
                        processed_row[c_idx] = clean_str
                    elif isinstance(val, (datetime.date, datetime.datetime)):
                        processed_row[c_idx] = val.strftime("%m/%d/%Y")
                    else:
                        processed_row[c_idx] = str(val) if val is not None else ""
                st.session_state.coupon_pool.append(processed_row)
                st.toast("需求已进入需求池")

# --- TAB 2: 校验流 ---
with tab2:
    if not st.session_state.coupon_pool:
        st.write("当前需求池为空。")
    else:
        st.subheader("🧐 自动校验报告")
        pool_df = pd.DataFrame(st.session_state.coupon_pool)
        
        # 如果有库存表，进行交叉比对
        if st.session_state.inventory_db is not None:
            for idx, row in pool_df.iterrows():
                # 寻找 ASIN 列索引
                asin_col = next((c['col'] for c in st.session_state.field_configs if "ASIN" in c['label'].upper()), None)
                if asin_col:
                    asins = row[asin_col].split(';')
                    invalid_asins = [a for a in asins if a not in st.session_state.inventory_db.index]
                    if invalid_asins:
                        st.warning(f"行 {idx+1}: 发现 {len(invalid_asins)} 个 ASIN 不在 All Listing Report 中：{invalid_asins[:3]}...")
                    else:
                        st.success(f"行 {idx+1}: 所有 ASIN 均在库在售。")
        
        # 预览预览
        mapping = {c['col']: c['label'] for c in st.session_state.field_configs}
        st.dataframe(pool_df.rename(columns=mapping), use_container_width=True)

        if st.button("🚀 寻找模板空行并导出 Excel"):
            template_file.seek(0)
            wb = load_workbook(template_file)
            ws = wb.active
            # 定位空行
            curr_r = 8
            while ws.cell(row=curr_r, column=1).value and str(ws.cell(row=curr_r, column=1).value).strip():
                curr_r += 1
            # 写入
            for data in st.session_state.coupon_pool:
                for c_idx, v in data.items():
                    ws.cell(row=curr_r, column=int(c_idx)).value = v
                curr_r += 1
            
            buf = io.BytesIO()
            wb.save(buf)
            st.download_button("💾 点击下载生成的上传文件", buf.getvalue(), f
