import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Fast Excel")

# Hàm xử lý file
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.title("⚡ Trình Lọc Excel Siêu Tốc")

# Tải file
uploaded_file = st.file_uploader("Chọn file", type=["xlsx", "csv"])

if uploaded_file:
    # Đọc dữ liệu nhanh
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    st.write(f"Tổng cộng: {len(df)} hàng")
    
    # Bộ lọc đơn giản
    col = st.selectbox("Cột cần lọc", df.columns)
    search = st.text_input("Từ khóa tìm kiếm")
    
    if search:
        df = df[df[col].astype(str).str.contains(search, case=False, na=False)]
        st.success(f"Còn lại: {len(df)} hàng")
    
    st.dataframe(df.head(100))
    
    # Nút tải về
    st.download_button("📥 Tải file kết quả", to_excel(df), "ketqua.xlsx")
