import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
import datetime

# --- 1. 页面配置 ---
st.set_page_config(page_title="Cupshe 亚马逊优惠券助手", layout="wide")

# --- 2. 初始化 Session State ---
if 'coupon_pool' not in st.session_state:
    st.session_state.coupon_pool = []
if 'field_configs' not in st.session_state:
    st.session_state.field_configs = []

# --- 3. 辅助函数：ASIN 格式标准化 ---
def clean_asin_format(raw_text):
    """
    将任何形式输入的 ASIN 转化为以英文分号分隔的标准格式
    """
    if not raw_text:
        return ""
    # 匹配所有常见分隔符：换行符 \n, 逗号 ,, 分号 ;, 空格 \s
    # 同时也兼容中文分号 ； 和中文逗号 ，
    tokens = re.split(r'[;；,，\s\n\r]+', raw_text.strip())
    # 过滤空字符串并去重（保持顺序）
    clean_list = []
    for t in tokens:
        t = t.strip().upper() # 亚马逊 ASIN 通常为大写
        if t and t not in clean_list:
            clean_list.append(t)
    
    # 按照第 5 行要求，用英文分号拼接输出
    return ";".join(clean_list), len(clean_list)

# --- 4. 模板解析函数 ---
def parse_template_v6(file):
    if file is None: return []
    try:
        file.seek(0)
        wb = load_workbook(file, data_only=True)
        ws = wb.active
        configs = []
        
        for col in range(1, 26):
            title = ws.cell(row=7, column=col).value
            rule_hint = ws.cell(row=5, column=col).value
            
            if title:
                title_str = str(title).strip()
                hint_str = str(rule_hint).strip() if rule_hint else ""
                
                # 下拉选项提取（第8, 9行）
                samples = []
                for r in [8, 9]:
                    v = ws.cell(row=r, column=col).value
                    if v:
                        s_v = str(v).strip()
                        if s_v not in samples: samples.append(s_v)
                
                is_dropdown = False
                opts = None
                dropdown_keys = ["折扣类型", "兑换一次", "限购", "优惠券类型", "目标买家", "叠加"]
                if any(k in title_str for k in dropdown_keys):
                    is_dropdown = True
                    opts = samples if samples else ["折扣", "满减", "是", "否"]

                configs.append({
                    "col": col,
                    "label": title_str,
                    "hint": hint_str,
                    "is_dropdown": is_dropdown,
                    "options": opts
                })
        return configs
    except Exception as e:
        st.error(f"解析出错: {e}")
        return []

if template_file := st.sidebar.file_uploader("2. 上传 Coupon 文件模板", type=['xlsx']):
    if not st.session_state.field_configs:
        st.session_state.field_configs = parse_template_v6(template_file)

# --- 5. 主界面 ---
st.title("👗 Cupshe 亚马逊优惠券智能管理工具")

if not st.session_state.field_configs:
    st.info("👋 请在侧边栏上传模板，系统将根据第 5 行要求自动格式化 ASIN。")
else:
    with st.form("coupon_form", clear_on_submit=True):
        st.subheader("1️⃣ 录入需求")
        user_input_raw = {}
        grid = st.columns(2)
        
        for i, cfg in enumerate(st.session_state.field_configs):
            with grid[i % 2]:
                label = cfg['label']
                hint = cfg['hint']
                fid = f"field_{cfg['col']}"
                
                if cfg['is_dropdown']:
                    user_input_raw[cfg['col']] = st.selectbox(label, options=cfg['options'], help=hint, key=fid)
                elif any(x in label.upper() for x in ["ASIN", "列表"]):
                    user_input_raw[cfg['col']] = st.text_area(label, help=f"规则参考：{hint}", placeholder="支持换行、空格或逗号输入", key=fid)
                elif any(x in label for x in ["日期", "Date"]):
                    tomorrow = datetime.date.today() + datetime.timedelta(days=1)
                    user_input_raw[cfg['col']] = st.date_input(label, value=tomorrow, help=hint, key=fid)
                else:
                    user_input_raw[cfg['col']] = st.text_input(label, help=hint, key=fid)
        
        if st.form_submit_button("➕ 添加到待处理列表"):
            formatted_entry = {}
            for c_idx, val in user_input_raw.items():
                # 核心功能：识别标题是否为 ASIN 列并进行格式重组
                cfg_item = next(c for c in st.session_state.field_configs if c['col'] == c_idx)
                if any(x in cfg_item['label'].upper() for x in ["ASIN", "列表"]):
                    # 自动转化为分号分隔
                    final_asin_str, count = clean_asin_format(str(val))
                    formatted_entry[c_idx] = final_asin_str
                    if count > 0:
                        st.toast(f"已自动处理 {count} 个 ASIN 为分号分隔格式")
                elif isinstance(val, (datetime.date, datetime.datetime)):
                    formatted_entry[c_idx] = val.strftime("%m/%d/%Y")
                else:
                    formatted_entry[c_idx] = str(val) if val is not None else ""
            
            st.session_state.coupon_pool.append(formatted_entry)

    # --- 预览与生成 ---
    if st.session_state.coupon_pool:
        st.divider()
        st.subheader("📋 待生成列表（ASIN 已自动转换格式）")
        mapping = {c['col']: c['label'] for c in st.session_state.field_configs}
        st.dataframe(pd.DataFrame(st.session_state.coupon_pool).rename(columns=mapping), use_container_width=True)
        
        if st.button("🚀 寻找空行并导出 Excel"):
            template_file.seek(0)
            wb = load_workbook(template_file)
            ws = wb.active
            
            # 寻找首个空行
            start_r = 8
            while ws.cell(row=start_r, column=1).value is not None:
                if not str(ws.cell(row=start_r, column=1).value).strip(): break
                start_r += 1
            
            # 写入
            for offset, data in enumerate(st.session_state.coupon_pool):
                write_r = start_r + offset
                for c_idx, value in data.items():
                    ws.cell(row=write_r, column=int(c_idx)).value = value
            
            buf = io.BytesIO()
            wb.save(buf)
            st.download_button("💾 下载最终文件", buf.getvalue(), f"Coupon_Fixed_{datetime.date.today()}.xlsx")
