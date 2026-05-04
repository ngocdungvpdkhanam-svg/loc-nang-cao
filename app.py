import streamlit as st
import pandas as pd
import io

# ── Cấu hình giao diện tối ưu ────────────────────────────────────────────────
st.set_page_config(page_title="Excel Ultra Manager", page_icon="⚡", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 2rem; font-weight: 800; color: #1e3a5f; margin-bottom: 0.5rem; }
    .stat-box { background: #f0f4ff; border-radius: 8px; padding: 10px; text-align: center; border: 1px solid #2e86de; }
    .stat-num { font-size: 1.5rem; font-weight: 700; color: #1e3a5f; }
    .filter-card { background: #ffffff; border: 1px solid #d1d5db; border-radius: 10px; padding: 15px; margin-top: 10px; border-left: 5px solid #2e86de; }
</style>
""", unsafe_allow_html=True)

# ── Hàm đọc dữ liệu siêu tốc ──────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data_ultra(file_content, header_idx):
    # Sử dụng engine calamine để đọc file lớn cực nhanh
    df = pd.read_excel(io.BytesIO(file_content), engine='calamine', header=header_idx)
    
    # Tối ưu bộ nhớ: Làm sạch tên cột
    df.columns = [str(c).strip() for c in df.columns]
    
    # Tối ưu RAM: Ép kiểu dữ liệu thấp nhất có thể
    for col in df.select_dtypes(include=['float64', 'int64']).columns:
        df[col] = pd.to_numeric(df[col], downcast='significant')
    
    return df

def to_excel_fast(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="main-title">⚡ Ultra Fast</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Tải file .xlsx lớn", type=["xlsx"])
    h_row = st.number_input("Dòng tiêu đề:", min_value=1, value=6)
    st.divider()
    st.warning("⚠️ Với file >300k dòng, hãy hạn chế chọn quá nhiều cột lọc cùng lúc.")

# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    with st.spinner("🚀 Đang xử lý dữ liệu khủng (Engine: Calamine)..."):
        file_content = uploaded_file.getvalue()
        df = load_data_ultra(file_content, h_row - 1)
        total_rows = len(df)

    # ── PHẦN 1: THỐNG KÊ NHANH ──────────────────────────────────────────────
    st.subheader("📋 Thông tin file")
    cols = df.columns.tolist()
    
    # Chỉ tính toán thống kê nếu người dùng mở Expander để tiết kiệm CPU
    with st.expander("Xem chi tiết số lượng X và Ô Trống từng cột"):
        if st.button("📊 Chạy thống kê chi tiết (Tốn RAM)"):
            info_cols = st.columns(4)
            for i, col in enumerate(cols):
                with info_cols[i % 4]:
                    n_blank = df[col].isna().sum()
                    n_x = (df[col].astype(str).str.upper() == 'X').sum()
                    st.write(f"**{col}**")
                    st.caption(f"Trống: {n_blank} | X: {n_x}")

    # ── PHẦN 2: BỘ LỌC TỐI ƯU RAM ──────────────────────────────────────────
    st.subheader("🔍 Bộ lọc thông minh")
    sel_cols = st.multiselect("Chọn cột cần lọc:", cols)
    
    # Khởi tạo Mask (Mảng Boolean)
    final_mask = pd.Series([True] * total_rows)

    if sel_cols:
        logic_mode = st.radio("Logic:", ["VÀ (AND)", "HOẶC (OR)"], horizontal=True)
        
        col_masks = []
        for col in sel_cols:
            st.markdown(f'<div class="filter-card">', unsafe_allow_html=True)
            st.write(f"📍 Lọc cột: **{col}**")
            
            c1, c2, c3 = st.columns([1, 1, 2])
            with c1: opt_blank = st.checkbox(f"Ô Trống", key=f"b_{col}")
            with c2: opt_x = st.checkbox(f"Giá trị X", key=f"x_{col}")
            with c3: 
                search_val = st.text_input(f"Tìm từ khóa khác:", key=f"t_{col}", help="Nhập giá trị cụ thể")

            # Xây dựng mask cho từng cột lọc
            m = pd.Series([False] * total_rows)
            if opt_blank: m |= df[col].isna()
            if opt_x: m |= (df[col].astype(str).str.upper() == 'X')
            if search_val: m |= (df[col].astype(str).str.contains(search_val, case=False, na=False))
            
            # Nếu không chọn gì ở cột này coi như lấy hết (True)
            if not opt_blank and not opt_x and not search_val:
                m = pd.Series([True] * total_rows)
            
            col_masks.append(m)
            st.markdown('</div>', unsafe_allow_html=True)

        # Gộp các mask lại theo logic
        if col_masks:
            if logic_mode == "VÀ (AND)":
                for cm in col_masks: final_mask &= cm
            else:
                final_mask = col_masks[0]
                for cm in col_masks[1:]: final_mask |= cm

    # ── PHẦN 3: KẾT QUẢ ────────────────────────────────────────────────────
    df_result = df[final_mask]
    
    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.markdown(f'<div class="stat-box"><div class="stat-num">{total_rows:,}</div><div class="stat-label">Hàng gốc</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_result):,}</div><div class="stat-label">Hàng khớp</div></div>', unsafe_allow_html=True)
    
    if not df_result.empty:
        with c3:
            if st.button("📦 Chuẩn bị file tải về", use_container_width=True, type="primary"):
                with st.spinner("Đang nén dữ liệu..."):
                    excel_data = to_excel_fast(df_result)
                    st.download_button("📥 Tải ngay (.xlsx)", data=excel_data, file_name="ket_qua_khung.xlsx", use_container_width=True)
        
        # Chỉ hiển thị 500 hàng đầu tiên để tránh treo trình duyệt
        st.write("---")
        st.caption(f"Đang hiển thị 500/{len(df_result)} hàng đầu tiên:")
        st.dataframe(df_result.head(500), use_container_width=True)
    else:
        st.warning("Không có kết quả nào.")

else:
    st.info("👋 Hãy tải file Excel (tối đa 500k-1M dòng) để bắt đầu xử lý.")
