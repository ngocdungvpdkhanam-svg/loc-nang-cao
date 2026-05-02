import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Multi-Filter Pro", layout="wide")

# Hàm chuyển đổi dữ liệu để tải về
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.title("📊 Bộ Lọc Excel Đa Cột & Tùy Chỉnh Tiêu Đề")

# 1. THANH BÊN (SIDEBAR) - CÀI ĐẶT FILE
with st.sidebar:
    st.header("⚙️ Cài đặt file")
    uploaded_file = st.file_uploader("Tải file Excel hoặc CSV", type=["xlsx", "csv"])
    
    header_row = st.number_input(
        "Dòng chứa tiêu đề (Header):", 
        min_value=1, 
        value=1, 
        help="Chọn dòng mà tên cột bắt đầu (Thường là dòng 1)"
    )
    
    # Chuyển đổi số dòng người dùng nhập (1-based) sang index của Pandas (0-based)
    pandas_header = header_row - 1

# 2. XỬ LÝ DỮ LIỆU
if uploaded_file:
    # Hàm đọc dữ liệu (Dùng Cache để chạy nhanh)
    # Lưu ý: Khi đổi header_row hoặc file, cache sẽ tự cập nhật
    @st.cache_data
    def get_data(file, p_header):
        try:
            if file.name.endswith(".csv"):
                file.seek(0)
                return pd.read_csv(file, header=p_header)
            else:
                file.seek(0)
                # Đọc tạm để lấy tên sheet
                xls = pd.ExcelFile(file)
                sheet = xls.sheet_names[0] # Mặc định lấy sheet đầu tiên
                return pd.read_excel(file, sheet_name=sheet, header=p_header)
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")
            return None

    df_original = get_data(uploaded_file, pandas_header)

    if df_original is not None:
        # Xóa các dòng hoàn toàn rỗng để dữ liệu sạch hơn
        df_original = df_original.dropna(how="all").reset_index(drop=True)
        df_display = df_original.copy()

        st.sidebar.success(f"✅ Đã tải {len(df_original)} hàng")

        # 3. KHU VỰC LỌC NHIỀU CỘT
        st.subheader("🔍 Thiết lập bộ lọc đa cột")
        
        # Chọn những cột muốn dùng để lọc
        all_columns = df_original.columns.tolist()
        cols_to_filter = st.multiselect(
            "Chọn các cột bạn muốn nhập từ khóa lọc:", 
            options=all_columns,
            default=[]
        )

        # Tạo các ô nhập liệu tương ứng
        if cols_to_filter:
            filters = {}
            cols_ui = st.columns(3) # Chia làm 3 cột giao diện
            for i, col_name in enumerate(cols_to_filter):
                with cols_ui[i % 3]:
                    search_val = st.text_input(f"Lọc theo: {col_name}", key=f"filter_{col_name}")
                    if search_val:
                        filters[col_name] = search_val

            # Thực hiện lọc cộng dồn (AND)
            for col, val in filters.items():
                df_display = df_display[df_display[col].astype(str).str.contains(str(val), case=False, na=False)]

        st.divider()

        # 4. HIỂN THỊ THỐNG KÊ & TẢI FILE
        c1, c2, c3 = st.columns([1, 1, 2])
        c1.metric("Tổng hàng gốc", len(df_original))
        c2.metric("Kết quả lọc", len(df_display))
        
        with c3:
            if len(df_display) > 0:
                st.download_button(
                    label="📥 Tải kết quả về máy (Excel)",
                    data=to_excel(df_display),
                    file_name="ket_qua_loc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

        # Hiển thị bảng (xem trước 100 dòng)
        st.dataframe(df_display.head(100), use_container_width=True)
        if len(df_display) > 100:
            st.caption(f"Đang hiển thị 100/{len(df_display)} hàng.")

else:
    st.info("👋 Hãy tải file Excel/CSV ở menu bên trái. Nếu file có tiêu đề ở dòng 2 hoặc 3, hãy chỉnh 'Dòng chứa tiêu đề' tương ứng.")

st.divider()
st.caption("Công cụ lọc dữ liệu chuyên nghiệp - Hỗ trợ tùy chỉnh Header")
