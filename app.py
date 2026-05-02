import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Lite", layout="wide")

# --- Khởi tạo dữ liệu ---
if "df" not in st.session_state:
    st.session_state.df = None
if "conds" not in st.session_state:
    st.session_state.conds = []

def to_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    return out.getvalue()

# --- Giao diện Sidebar ---
with st.sidebar:
    st.title("📂 Cài đặt")
    file = st.file_uploader("Tải file Excel/CSV", type=['xlsx', 'xls', 'csv'])
    if file:
        file.seek(0)
        if file.name.endswith('.csv'):
            df_input = pd.read_csv(file)
        else:
            sheets = pd.ExcelFile(file).sheet_names
            s = st.selectbox("Chọn sheet", sheets)
            df_input = pd.read_excel(file, sheet_name=s)
        
        if st.button("Làm mới dữ liệu") or st.session_state.df is None:
            st.session_state.df = df_input.copy()
            st.session_state.conds = []
            st.rerun()

# --- Giao diện chính ---
st.title("📊 Bộ lọc Excel Tối giản")

if st.session_state.df is not None:
    df = st.session_state.df
    st.write(f"Dữ liệu hiện tại: **{len(df)} hàng** | **{len(df.columns)} cột**")

    # --- Khu vực tạo bộ lọc ---
    with st.expander("🔍 Thêm điều kiện lọc", expanded=True):
        c1, c2, c3 = st.columns([2, 1, 2])
        col = c1.selectbox("Cột lọc", df.columns)
        op = c2.selectbox("Phép toán", [">", "<", "==", "!=", "chứa", "trống"])
        val = c3.text_input("Giá trị (để trống nếu chọn 'trống')")
        
        if st.button("➕ Thêm"):
            st.session_state.conds.append({"col": col, "op": op, "val": val})
            st.rerun()

    # --- Áp dụng bộ lọc ---
    if st.session_state.conds:
        mask = pd.Series([True] * len(df), index=df.index)
        for c in st.session_state.conds:
            s_col = df[c['col']].astype(str)
            if c['op'] == "==": m = s_col == str(c['val'])
            elif c['op'] == "!=": m = s_col != str(c['val'])
            elif c['op'] == "chứa": m = s_col.str.contains(str(c['val']), case=False, na=False)
            elif c['op'] == "trống": m = df[c['col']].isna()
            elif c['op'] == ">": m = pd.to_numeric(df[c['col']], errors='coerce') > float(c['val'])
            elif c['op'] == "<": m = pd.to_numeric(df[c['col']], errors='coerce') < float(c['val'])
            mask &= m
        
        res = df[mask]
        st.warning(f"Đang chọn {len(res)} hàng dựa trên các điều kiện.")
        
        a1, a2, a3 = st.columns(3)
        if a1.button("🗑️ XÓA hàng đang chọn"):
            st.session_state.df = df[~mask].reset_index(drop=True)
            st.session_state.conds = []
            st.rerun()
        if a2.button("✅ CHỈ GIỮ hàng đang chọn"):
            st.session_state.df = res.reset_index(drop=True)
            st.session_state.conds = []
            st.rerun()
        if a3.button("❌ Xóa các bộ lọc"):
            st.session_state.conds = []
            st.rerun()

    # --- Hiển thị & Xuất file ---
    st.dataframe(df.head(100), use_container_width=True)
    
    st.subheader("📥 Xuất file")
    ex1, ex2 = st.columns(2)
    ex1.download_button("Tải Excel", to_excel(df), "ketqua.xlsx", use_container_width=True)
    ex2.download_button("Tải CSV", df.to_csv(index=False).encode('utf-8-sig'), "ketqua.csv", use_container_width=True)

else:
    st.info("Vui lòng tải file ở cột bên trái.")
