import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

st.set_page_config(page_title="Fast Excel Pro", layout="wide")

# --- HÀM TỐI ƯU: ĐỌC MÀU SẮC (Chỉ chạy khi cần) ---
@st.cache_data
def get_row_colors(file_content, header_idx):
    # Dùng BytesIO để đọc từ bộ nhớ
    file_bytes = io.BytesIO(file_content)
    wb = load_workbook(file_bytes, read_only=False, data_only=True)
    ws = wb.active
    
    colors = []
    # Chỉ quét cột A (cột 1) để lấy màu đại diện cho hàng -> Tăng tốc độ cực lớn
    for row in ws.iter_rows(min_row=header_idx + 2, min_col=1, max_col=1):
        cell = row[0]
        color_hex = "No Fill"
        # Kiểm tra màu nền
        if cell.fill and cell.fill.start_color:
            rgb = cell.fill.start_color.rgb
            if isinstance(rgb, str) and len(rgb) == 8:
                color_hex = f"#{rgb[2:]}"
        colors.append(color_hex)
    return colors

# --- HÀM TỐI ƯU: ĐỌC DỮ LIỆU (Dùng Pandas cực nhanh) ---
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    df = pd.read_excel(file_bytes, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all").reset_index(drop=True)

def to_excel(df):
    output = io.BytesIO()
    if '__color__' in df.columns:
        df = df.drop(columns=['__color__'])
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- GIAO DIỆN ---
st.title("⚡ Excel Multi-Filter (Speed Optimized)")

with st.sidebar:
    st.header("⚙️ Nhập dữ liệu")
    uploaded_file = st.file_uploader("Tải file .xlsx", type=["xlsx"])
    h_row = st.number_input("Dòng tiêu đề:", min_value=1, value=1)
    
    st.divider()
    enable_color = st.checkbox("🌈 Bật chế độ lọc màu", help="Chỉ bật khi cần lọc màu để đảm bảo tốc độ nhanh nhất")

if uploaded_file:
    # Bước 1: Đọc nội dung file vào bộ nhớ
    file_content = uploaded_file.getvalue()
    
    # Bước 2: Tải dữ liệu nhanh bằng Pandas
    df_raw = get_clean_data(file_content, h_row - 1)
    
    # Bước 3: Chỉ xử lý màu nếu người dùng yêu cầu
    if enable_color:
        with st.spinner("Đang quét mã màu..."):
            row_colors = get_row_colors(file_content, h_row - 1)
            # Đảm bảo độ dài mảng màu khớp với dữ liệu (tránh lỗi khi có dòng trống cuối file)
            df_raw['__color__'] = row_colors[:len(df_raw)]
    else:
        df_raw['__color__'] = "No Fill"

    # --- KHU VỰC LỌC ---
    st.subheader("🔍 Bộ lọc thông minh")
    
    col_color, col_logic = st.columns([1, 2])
    
    with col_color:
        if enable_color:
            unique_colors = df_raw['__color__'].unique().tolist()
            selected_colors = st.multiselect("Lọc theo màu:", unique_colors, default=unique_colors)
            mask = df_raw['__color__'].isin(selected_colors)
        else:
            st.caption("Chế độ lọc màu đang tắt.")
            mask = pd.Series([True] * len(df_raw))

    with col_logic:
        logic_mode = st.radio("Logic nội dung:", ["VÀ (AND)", "HOẶC (OR)"], horizontal=True)
        sel_cols = st.multiselect("Chọn cột lọc nội dung:", [c for c in df_raw.columns if c != '__color__'])

    # Áp dụng logic nội dung
    if sel_cols:
        sub_masks = []
        for c in sel_cols:
            val = st.text_input(f"Tìm trong {c}:", key=f"v_{c}")
            if val:
                sub_masks.append(df_raw[c].astype(str).str.contains(val, case=False, na=False))
        
        if sub_masks:
            if logic_mode == "VÀ (AND)":
                for sm in sub_masks: mask &= sm
            else:
                or_mask = sub_masks[0]
                for sm in sub_masks[1:]: or_mask |= sm
                mask &= or_mask

    # --- KẾT QUẢ ---
    df_final = df_raw[mask]
    
    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Tổng hàng", len(df_raw))
    c2.metric("Sau lọc", len(df_final))
    
    if not df_final.empty:
        c3.download_button("📥 Tải kết quả", to_excel(df_final), "ket_qua.xlsx", type="primary", use_container_width=True)
        
        # Hiển thị bảng kèm màu sắc mô phỏng (nếu có)
        st.dataframe(df_final, use_container_width=True)
    else:
        st.warning("Không có dữ liệu phù hợp.")

else:
    st.info("Hãy tải file lên để bắt đầu.")
