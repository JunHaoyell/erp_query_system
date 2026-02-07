import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# ===============================
# Page Config
# ===============================
st.set_page_config("ERP 内部查询系统", layout="wide")

# ===============================
# Password
# ===============================
if "auth" not in st.session_state:
    pwd = st.text_input("🔐 访问密码", type="password")
    if pwd != st.secrets["APP_PASSWORD"]:
        st.stop()
    st.session_state["auth"] = True

# ===============================
# Google Auth
# ===============================
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_dict(
    st.secrets["google_service_account"], scope
)
client = gspread.authorize(creds)

# ===============================
# Load Config
# ===============================
CONFIG_SHEET_ID = st.secrets["CONFIG_SHEET_ID"]

@st.cache_data(ttl=600)
def load_config():
    ws = client.open_by_key(CONFIG_SHEET_ID).worksheet("Config")
    return pd.DataFrame(ws.get_all_records())

@st.cache_data(ttl=600)
def load_all_data(cfg):
    dfs = []
    for _, r in cfg.iterrows():
        if not r["启用"]:
            continue
        ws = client.open_by_key(r["Sheet_ID"]).worksheet(r["Tab_Name"])
        data = ws.get_all_records()
        if not data:
            continue
        df = pd.DataFrame(data)
        df["来源名称"] = r["数据源名称"]
        df["来源Tab"] = r["Tab_Name"]
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

cfg = load_config()
df_all = load_all_data(cfg)

# ===============================
# 强制筛选
# ===============================
st.subheader("🔑 基础筛选（至少选一个）")

c1, c2, c3 = st.columns(3)
with c1:
    month = st.multiselect("月份", sorted(df_all["月份"].dropna().unique()))
with c2:
    bill = st.multiselect("单据编号", sorted(df_all["单据编号"].dropna().unique()))
with c3:
    room = st.multiselect("房间代码", sorted(df_all["房间代码"].dropna().unique()))

if not any([month, bill, room]):
    st.info("请选择：月份 / 单据编号 / 房间代码（至少一个）")
    st.stop()

df = df_all.copy()
if month:
    df = df[df["月份"].isin(month)]
if bill:
    df = df[df["单据编号"].isin(bill)]
if room:
    df = df[df["房间代码"].isin(room)]

# ===============================
# 高级筛选
# ===============================
st.subheader("🔍 高级筛选")

with st.expander("展开所有筛选条件", expanded=True):

    keyword = st.text_input("关键词（全文）")

    for col in [
        "单据类型","币种","借款报销部门","借款报销人","推广部组",
        "报销/冲销","用途","项目","项目类型","费用类型",
        "经办人","费用承担部门","账目类型","分类"
    ]:
        values = st.multiselect(col, sorted(df[col].dropna().unique()))
        if values:
            df = df[df[col].isin(values)]

    # 数值区间
    for col in ["数量","单价","原币金额","本币金额","本币汇率"]:
        if col in df.columns and df[col].notna().any():
            min_v, max_v = float(df[col].min()), float(df[col].max())
            r = st.slider(col, min_v, max_v, (min_v, max_v))
            df = df[df[col].between(r[0], r[1])]

    if keyword:
        text = df.astype(str).agg(" ".join, axis=1)
        df = df[text.str.contains(keyword, case=False, na=False)]

# ===============================
# 表格展示
# ===============================
MAX_ROWS = 5000
st.subheader(f"📋 查询结果：{len(df)} 条")

show_df = df.head(MAX_ROWS)
st.dataframe(show_df, use_container_width=True, height=600)

# ===============================
# 导出
# ===============================
def to_excel(d):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        d.to_excel(w, index=False)
    return buf.getvalue()

st.download_button(
    "⬇️ 导出 Excel",
    to_excel(df),
    "erp_query_result.xlsx"
)
