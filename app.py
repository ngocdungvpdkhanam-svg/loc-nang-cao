import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Multi-Filter Excel", layout="wide")

# Hàm chuyển đổi dữ liệu để tải về
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.title("📊 Bộ Lọc Đa Cột Siêu Tốc")

# 1. Tải file
uploaded_file = st.sidebar.file_uploader("Tải file Excel hoặc CSV", type=["xlsx", "csv"])

if uploaded_file:
    # Đọc dữ liệu (Dùng Cache để chạy cho nhanh)
    @st.cache_data
    def get_data(file):
        if file.name.endswith(".csv"):
            return pd.read_csv(file)
        return pd.read_excel(file)

    df_original = get_data(uploaded_file)
    df_display = df_original.copy()

    st.sidebar.success(f"Đã tải {len(df_original)} hàng")

    # 2. KHU VỰC LỌC NHIỀU CỘT
    st.subheader("🔍 Thiết lập bộ lọc")
    
    # Chọn những cột bạn muốn dùng để lọc
    cols_to_filter = st.multiselect(
        "Chọn các cột bạn muốn lọc dữ liệu:", 
        options=df_original.columns.tolist(),
        default=[]
    )

    # Tạo các ô nhập liệu tương ứng cho mỗi cột đã chọn
    if cols_to_filter:
        filters = {}
        # Chia làm 3 cột trên giao diện cho gọn
        cols_ui = st.columns(3)
        for i, col_name in enumerate(cols_to_filter):
            with cols_ui[i % 3]:
                search_val = st.text_input(f"Lọc cột: {col_name}", key=f"filter_{col_name}")
                if search_val:
                    filters[col_name] = search_val

        # Tiến hành lọc dữ liệu dựa trên tất cả các ô đã nhập
        for col, val in filters.items():
            df_display = df_display[df_display[col].astype(str).str.contains(val, case=False, na=False)]

    st.divider()

    # 3. HIỂN THỊ KẾT QUẢ
    col_stat1, col_stat2 = st.columns(2)
    col_stat1.metric("Tổng số hàng gốc", len(df_original))
    col_stat2.metric("Số hàng sau khi lọc", len(df_display), delta=len(df_display)-len(df_original))

    # Nút tải file
    if len(df_display) > 0:
        st.download_button(
            label="📥 Tải kết quả về máy (Excel)",
            data=to_excel(df_display),
            file_name="ket_qua_loc_da_cot.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    else:
        st.warning("Không có dữ liệu thỏa mãn bộ lọc.")

    # Hiển thị bảng dữ liệu (tối đa 100 hàng cho nhanh)
    st.dataframe(df_display.head(100), use_container_width=True)

else:
    st.info("👋 Vui lòng tải file Excel hoặc CSV ở menu bên trái để bắt đầu.")
