import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# ── Cấu hình giao diện ────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel Pro Manager", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 2.2rem; font-weight: 700; background: linear-gradient(135deg, #1e3a5f, #2e86de); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 0.5rem; }
    .stat-box { background: #f0f4ff; border-radius: 10px; padding: 15px; text-align: center; border-left: 5px solid #2e86de; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .stat-num { font-size: 1.8rem; font-weight: 700; color: #1e3a5f; }
    .stat-label { font-size: 0.8rem; color: #555; text-transform: uppercase; font-weight: 600; }
    .filter-card { background: #ffffff; border: 1px solid #d1d5db; border-radius: 10px; padding: 20px; margin-top: 15px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
    .blank-style { color: #cf222e; font-weight: bold; background: #fff1f0; padding: 2px 8px; border-radius: 4px; border: 1px solid #ffa39e; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý dữ liệu ────────────────────────────────────────────────────────
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    # Đọc file và xử lý các ô trống/chuỗi rỗng đồng nhất
    df = pd.read_excel(file_bytes, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    # Chuyển các ô chỉ có khoảng trắng thành thực sự rỗng (NaN)
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    return df.dropna(how="all").reset_index(drop=True)

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Cài đặt dữ liệu")
    uploaded_file = st.file_uploader("Tải file .xlsx", type=["xlsx"])
    h_row = st.number_input("Dòng tiêu đề (Header):", min_value=1, value=1)

# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    file_content = uploaded_file.getvalue()
    df_raw = get_clean_data(file_content, h_row - 1)
    display_cols = df_raw.columns.tolist()

    # ── PHẦN 1: TỔNG QUAN FILE ──────────────────────────────────────────────
    with st.expander("📋 THÔNG TIN CHI TIẾT CÁC CỘT", expanded=True):
        info_cols = st.columns(4)
        for i, col in enumerate(display_cols):
            with info_cols[i % 4]:
                nulls = df_raw[col].isna().sum()
                st.markdown(f'''
                <div style="background:white; padding:10px; border-radius:5px; border:1px solid #ddd; margin-bottom:5px;">
                    <b style="color:#2e86de;">{col}</b><br>
                    <small>Số ô trống: <span class="blank-style">{nulls}</span></small>
                </div>''', unsafe_allow_html=True)

    # ── PHẦN 2: BỘ LỌC CẢI TIẾN ──────────────────────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    
    # Chọn cột để lọc
    sel_cols = st.multiselect("BƯỚC 1: Chọn các cột bạn muốn lọc:", display_cols)
    
    mask = pd.Series([True] * len(df_raw))

    if sel_cols:
        st.write("BƯỚC 2: Cài đặt giá trị cho từng cột:")
        logic_mode = st.radio("Kết hợp điều kiện giữa các cột:", ["VÀ (Thỏa mãn tất cả)", "HOẶC (Chỉ cần thỏa mãn 1 cái)"], horizontal=True)
        
        sub_masks = []
        for col in sel_cols:
            # Tạo một khung trắng riêng cho mỗi cột để bạn không bị nhìn nhầm
            st.markdown(f'<div class="filter-card">', unsafe_allow_html=True)
            st.markdown(f"📍 Đang thiết lập cho cột: **{col}**")
            
            # --- CHỖ BẠN CẦN: CHỌN Ô TRỐNG & CHỌN TẤT CẢ ---
            c1, c2, c3 = st.columns([1.5, 1.5, 2])
            with c1:
                is_only_blank = st.checkbox(f"⚪ Chỉ lấy ô TRỐNG (Blanks)", key=f"blk_{col}")
            with c2:
                is_select_all = st.checkbox(f"✅ Chọn tất cả giá trị", key=f"all_{col}", value=False)
            
            # Lấy danh sách giá trị có trong cột (bỏ qua ô trống)
            val_options = sorted([str(x) for x in df_raw[col].dropna().unique()])
            
            # Xử lý Logic chọn
            if is_only_blank:
                # Nếu tích vào "Chỉ lấy ô trống" -> ẩn ô tìm kiếm và lọc luôn
                st.info(f"Đang lọc: Chỉ lấy những hàng mà cột '{col}' bị TRỐNG.")
                sub_masks.append(df_raw[col].isna())
            else:
                # Nếu không chỉ lấy trống -> hiện ô chọn giá trị
                default_val = val_options if is_select_all else []
                selected_vals = st.multiselect(
                    f"Chọn các giá trị cụ thể của [{col}]:",
                    options=val_options,
                    default=default_val,
                    key=f"ms_{col}"
                )
                
                # Tạo Mask
                if selected_vals:
                    sub_masks.append(df_raw[col].astype(str).isin(selected_vals))
                elif is_select_all:
                    sub_masks.append(pd.Series([True] * len(df_raw)))
                else:
                    # Nếu không chọn gì và không tích ô nào -> Coi như không lọc cột này
                    sub_masks.append(pd.Series([True] * len(df_raw)))
            
            st.markdown('</div>', unsafe_allow_html=True)

        # Kết hợp Mask
        if sub_masks:
            if "VÀ" in logic_mode:
                for sm in sub_masks: mask &= sm
            else:
                or_m = sub_masks[0]
                for sm in sub_masks[1:]: or_m |= sm
                mask &= or_m

    # ── PHẦN 3: KẾT QUẢ XỬ LÝ ──────────────────────────────────────────────
    df_final = df_raw[mask]
    
    st.write("---")
    st.markdown("### 📊 Kết quả sau khi lọc")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_raw):,}</div><div class="stat-label">Hàng gốc</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_final):,}</div><div class="stat-label">Hàng thỏa ĐK</div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#cf222e">{len(df_raw)-len(df_final):,}</div><div class="stat-label">Đã loại bỏ</div></div>', unsafe_allow_html=True)
    ratio = (len(df_final)/len(df_raw)*100) if len(df_raw)>0 else 0
    c4.markdown(f'<div class="stat-box"><div class="stat-num">{ratio:.1f}%</div><div class="stat-label">Tỷ lệ giữ</div></div>', unsafe_allow_html=True)

    if not df_final.empty:
        st.write("")
        st.download_button(
            label=f"📥 Tải file kết quả ({len(df_final)} hàng)",
            data=to_excel(df_final),
            file_name="ket_qua_loc.xlsx",
            type="primary",
            use_container_width=True
        )
        st.dataframe(df_final, use_container_width=True)
    else:
        st.warning("⚠️ Không có hàng nào thỏa mãn điều kiện bạn chọn.")
else:
    st.info("👋 Hãy tải file Excel lên để bắt đầu.")
