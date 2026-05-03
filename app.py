import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# ── Cấu hình giao diện chuẩn ──────────────────────────────────────────────────
st.set_page_config(page_title="Excel Pro Manager", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 2.2rem; font-weight: 700; background: linear-gradient(135deg, #1e3a5f, #2e86de); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 0.5rem; }
    .stat-box { background: #f0f4ff; border-radius: 10px; padding: 15px; text-align: center; border-left: 5px solid #2e86de; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .stat-num { font-size: 1.8rem; font-weight: 700; color: #1e3a5f; }
    .stat-label { font-size: 0.8rem; color: #555; text-transform: uppercase; font-weight: 600; }
    .filter-card { background: #ffffff; border: 1px solid #d1d5db; border-radius: 10px; padding: 20px; margin-top: 10px; border-top: 5px solid #2e86de; }
    .badge-x { background: #e6ffed; color: #1e7e34; padding: 2px 8px; border-radius: 5px; font-weight: bold; border: 1px solid #b7eb8f; }
    .badge-blank { background: #fff1f0; color: #cf222e; padding: 2px 8px; border-radius: 5px; font-weight: bold; border: 1px solid #ffa39e; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý dữ liệu tối ưu ──────────────────────────────────────────────────
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    df = pd.read_excel(file_bytes, header=header_idx)
    # Làm sạch tên cột và đồng bộ hóa dữ liệu
    df.columns = [str(c).strip() if pd.notnull(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]
    return df.dropna(how="all").reset_index(drop=True)

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="main-title" style="font-size:1.5rem;">📁 Cài đặt dữ liệu</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Tải file .xlsx", type=["xlsx"])
    h_row = st.number_input("Dòng tiêu đề (Header):", min_value=1, value=6)
    st.divider()
    st.info("💡 Nếu danh sách giá trị quá lớn, hãy dùng ô tìm kiếm để tránh treo máy.")

# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    file_content = uploaded_file.getvalue()
    df_raw = get_clean_data(file_content, h_row - 1)
    display_cols = df_raw.columns.tolist()

    # ── PHẦN 1: TỔNG QUAN CHI TIẾT ──────────────────────────────────────────
    with st.expander("📊 THÔNG TIN CỘT & ĐẾM GIÁ TRỊ", expanded=True):
        info_cols = st.columns(4)
        for i, col in enumerate(display_cols):
            with info_cols[i % 4]:
                nulls = df_raw[col].isna().sum()
                x_count = df_raw[col].astype(str).str.strip().str.upper().eq('X').sum()
                st.markdown(f'''
                <div style="background:white; padding:10px; border-radius:8px; border:1px solid #eee; margin-bottom:10px;">
                    <b style="color:#2e86de; font-size:0.9rem;">{col}</b><br>
                    <small>Trống: <span style="color:red;">{nulls}</span> | X: <span style="color:green;">{x_count}</span></small>
                </div>''', unsafe_allow_html=True)

    # ── PHẦN 2: BỘ LỌC CẢI TIẾN (CHỐNG TREO MÁY) ──────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    sel_cols = st.multiselect("BƯỚC 1: Chọn các cột bạn muốn lọc:", display_cols)
    
    mask = pd.Series([True] * len(df_raw))

    if sel_cols:
        logic_mode = st.radio("Kết hợp điều kiện:", ["VÀ (AND)", "HOẶC (OR)"], horizontal=True)
        
        sub_masks = []
        for col in sel_cols:
            st.markdown(f'<div class="filter-card">', unsafe_allow_html=True)
            
            # Thống kê nhanh tại chỗ
            n_blank = df_raw[col].isna().sum()
            n_x = df_raw[col].astype(str).str.strip().str.upper().eq('X').sum()
            st.markdown(f"📍 Đang lọc: **{col}**  (Có <span class='badge-blank'>{n_blank} ô Trống</span> và <span class='badge-x'>{n_x} ô X</span>)", unsafe_allow_html=True)
            
            # Lựa chọn kiểu lọc
            c1, c2, c3 = st.columns([1, 1, 1.5])
            with c1: is_blk = st.checkbox(f"Lấy ô TRỐNG", key=f"b_{col}")
            with c2: is_x = st.checkbox(f"Lấy giá trị X", key=f"x_{col}")
            with c3: is_all = st.checkbox(f"Chọn tất cả giá trị khác", key=f"a_{col}")

            # Xử lý danh sách giá trị để xổ ra (Chỉ lấy tối đa 1000 giá trị đầu để tránh treo)
            unique_vals = df_raw[col].dropna()
            # Loại bỏ giá trị X ra khỏi danh sách chọn vì đã có nút riêng
            unique_vals = unique_vals[unique_vals.astype(str).str.strip().str.upper() != 'X']
            unique_vals = sorted(unique_vals.unique().astype(str))
            
            if len(unique_vals) > 500:
                st.warning("⚠️ Cột này có quá nhiều giá trị (>500), hãy nhập từ khóa tìm kiếm bên dưới:")
                search_text = st.text_input(f"Tìm giá trị trong {col}:", key=f"txt_{col}")
                selected_vals = [search_text] if search_text else []
            else:
                default_sel = unique_vals if is_all else []
                selected_vals = st.multiselect(f"Chọn các giá trị khác trong [{col}]:", options=unique_vals, default=default_sel, key=f"ms_{col}")

            # Xây dựng Mask cho từng cột
            col_mask = pd.Series([False] * len(df_raw))
            if is_blk: col_mask |= df_raw[col].isna()
            if is_x: col_mask |= df_raw[col].astype(str).str.strip().str.upper() == 'X'
            if is_all: col_mask |= df_raw[col].notna() # Lấy tất cả những gì không trống
            if selected_vals:
                # Nếu dùng search text
                if len(unique_vals) > 500 and search_text:
                    col_mask |= df_raw[col].astype(str).str.contains(search_text, case=False, na=False)
                else:
                    col_mask |= df_raw[col].astype(str).isin(selected_vals)
            
            # Nếu không tích gì thì coi như không lọc cột này (Hiện tất cả)
            if not is_blk and not is_x and not is_all and not selected_vals:
                col_mask = pd.Series([True] * len(df_raw))
                
            sub_masks.append(col_mask)
            st.markdown('</div>', unsafe_allow_html=True)

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
    st.markdown("### 📊 Kết quả thống kê")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_raw):,}</div><div class="stat-label">Hàng gốc</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_final):,}</div><div class="stat-label">Hàng khớp</div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#cf222e">{len(df_raw)-len(df_final):,}</div><div class="stat-label">Đã loại</div></div>', unsafe_allow_html=True)
    ratio = (len(df_final)/len(df_raw)*100) if len(df_raw)>0 else 0
    c4.markdown(f'<div class="stat-box"><div class="stat-num">{ratio:.1f}%</div><div class="stat-label">Tỷ lệ giữ</div></div>', unsafe_allow_html=True)

    if not df_final.empty:
        st.write("")
        st.download_button(label=f"📥 Tải Excel kết quả ({len(df_final)} hàng)", data=to_excel(df_final), file_name="ket_qua.xlsx", type="primary", use_container_width=True)
        st.dataframe(df_final, use_container_width=True)
    else:
        st.warning("⚠️ Không có hàng nào thỏa mãn điều kiện lọc.")
else:
    st.info("👋 Hãy tải file Excel lên để bắt đầu.")
