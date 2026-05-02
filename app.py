import streamlit as st
import pandas as pd
import io

# Tối ưu hóa tiêu đề trang
st.set_page_config(page_title="Excel Fast", layout="wide")

# Hàm đọc file có Cache để tăng tốc xử lý
@st.cache_data
def load_data(file, sheet=None):
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    return pd.read_excel(file, sheet_name=sheet)

# Hàm xuất file
def to_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    return out.getvalue()

# GIAO DIỆN
st.title("📊 Bộ Lọc Excel Siêu Tốc")

if "df" not in st.session_state:
    st.session_state.df = None

# Sidebar tải file
with st.sidebar:
    f = st.file_uploader("Tải file", type=['xlsx', 'csv'])
    if f:
        sheet = None
        if f.name.endswith('.xlsx'):
            try:
                sheet = st.selectbox("Chọn sheet", pd.ExcelFile(f).sheet_names)
            except: pass
        
        if st.button("🚀 Xử lý file này"):
            st.session_state.df = load_data(f, sheet)
            st.rerun()

# Xử lý dữ liệu
if st.session_state.df is not None:
    df = st.session_state.df
    st.write(f"Đang có: **{len(df)} hàng**")
    
    # Lọc nhanh
    col = st.selectbox("Chọn cột lọc", df.columns)
    val = st.text_input("Giá trị lọc (Enter để tìm)")
    
    if val:
        mask = df[col].astype(str).str.contains(val, case=False, na=False)
        filtered_df = df[mask]
        st.success(f"Tìm thấy {len(filtered_df)} hàng")
        st.dataframe(filtered_df.head(50))
        
        # Nút thao tác nhanh
        c1, c2 = st.columns(2)
        if c1.button("🗑️ Xóa các hàng này"):
            st.session_state.df = df[~mask].reset_index(drop=True)
            st.rerun()
        if c2.button("📥 Tải về kết quả"):
            st.download_button("Download Excel", to_excel(filtered_df), "ketqua.xlsx")
    else:
        st.dataframe(df.head(100))
        
    if st.button("🔄 Reset toàn bộ"):
        st.session_state.df = None
        st.rerun()
else:
    st.info("Hãy tải file Excel/CSV ở bên trái.")
