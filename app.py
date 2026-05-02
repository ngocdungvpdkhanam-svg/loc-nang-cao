import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# ── Cấu hình trang & Giao diện ──────────────────────────────────────────────
st.set_page_config(page_title="Excel Pro Manager", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 2.2rem; font-weight: 700; background: linear-gradient(135deg, #1e3a5f, #2e86de); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 0.5rem; }
    .stat-box { background: #f0f4ff; border-radius: 10px; padding: 15px; text-align: center; border-left: 5px solid #2e86de; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .stat-num { font-size: 1.8rem; font-weight: 700; color: #1e3a5f; }
    .stat-label { font-size: 0.8rem; color: #555; text-transform: uppercase; font-weight: 600; }
    .filter-card { background: #fdfdfd; border: 1px solid #e1e4e8; border-radius: 10px; padding: 15px; margin-bottom: 20px; }
    .empty-alert { color: #cf222e; font-weight: bold; font-size: 0.9rem; border: 1px solid #cf222e; padding: 2px 5px; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý dữ liệu ──────────────────────────────────────────────────────
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    df = pd.read_excel(file_bytes, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all").reset_index(drop=True)

def to_excel(df):
    output = io.BytesIO()
    if '__color__' in df.columns: df = df.drop(columns=['__color__'])
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="main-title" style="font-size:1.5rem;">📊 Tải file</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Chọn file Excel (.xlsx)", type=["xlsx"])
    h_row = st.number_input("Dòng tiêu đề (Header):", min_value=1, value=1)
    st.info("💡 Mẹo: Nếu muốn chọn tất cả hàng trống, hãy chọn cột cần lọc và tích vào ô 'Chỉ lấy ô TRỐNG'.")

# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    file_content = uploaded_file.getvalue()
    df_raw = get_clean_data(file_content, h_row - 1)
    display_cols = df_raw.columns.tolist()

    # ── PHẦN 1: TỔNG QUAN FILE ──────────────────────────────────────────────
    with st.expander("📋 CHI TIẾT CÁC CỘT & Ô TRỐNG", expanded=True):
        info_cols = st.columns(4)
        for i, col in enumerate(display_cols):
            with info_cols[i % 4]:
                null_count = df_raw[col].isna().sum()
                st.markdown(f"""
                <div style="border: 1px solid #ddd; padding: 8px; border-radius: 5px; margin-bottom: 5px; background: #fff;">
                    <b style="color:#2e86de;">{col}</b><br>
                    <span style="color:#666; font-size:0.8rem;">Ô trống: </span>
                    <span class="{"empty-alert" if null_count > 0 else ""}" style="font-size:0.8rem;">{null_count}</span>
                </div>
                """, unsafe_allow_html=True)

    # ── PHẦN 2: BỘ LỌC CẢI TIẾN ──────────────────────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    
    sel_cols = st.multiselect("Chọn các cột bạn muốn thực hiện lọc:", display_cols)
    
    mask = pd.Series([True] * len(df_raw))

    if sel_cols:
        logic_mode = st.radio("Kết hợp các cột lọc theo kiểu:", ["VÀ (Thỏa mãn tất cả các cột)", "HOẶC (Thỏa mãn 1 trong các cột)"], horizontal=True)
        
        sub_masks = []
        for col in sel_cols:
            # Tạo một khung riêng cho từng cột lọc
            st.markdown(f'<div class="filter-card">', unsafe_allow_html=True)
            st.markdown(f"**📍 ĐANG LỌC CỘT: `{col}`**")
            
            c1, c2 = st.columns([1, 1])
            with c1:
                # TÙY CHỌN CHỌN TẤT CẢ Ô TRỐNG
                only_empty = st.checkbox(f"Chỉ lấy các hàng TRỐNG (Blank) của cột này", key=f"empty_{col}")
            
            # Nếu không chỉ lấy ô trống thì mới hiện danh sách chọn giá trị
            if not only_empty:
                with c2:
                    select_all = st.checkbox(f"Chọn tất cả giá trị có sẵn", key=f"all_{col}", value=False)
                
                # Lấy danh sách giá trị thực tế (loại bỏ NaN)
                u_vals = sorted([str(x) for x in df_raw[col].dropna().unique()])
                
                default_val = u_vals if select_all else []
                selected = st.multiselect(f"Chọn giá trị cụ thể cho [{col}]:", options=u_vals, default=default_val, key=f"ms_{col}")
                
                # Logic tạo Mask cho cột
                if selected:
                    sub_masks.append(df_raw[col].astype(str).isin(selected))
                elif select_all:
                    sub_masks.append(pd.Series([True] * len(df_raw)))
                else:
                    # Nếu không chọn "Chỉ lấy trống" và cũng không chọn giá trị nào -> Mặc định là không lọc (True) hoặc tùy bạn chọn
                    # Ở đây tôi để là True để không làm mất dữ liệu nếu lỡ bấm chọn cột mà chưa kịp thao tác
                    sub_masks.append(pd.Series([True] * len(df_raw)))
            else:
                # NẾU TÍCH VÀO CHỈ LẤY Ô TRỐNG
                sub_masks.append(df_raw[col].isna())
            
            st.markdown('</div>', unsafe_allow_html=True)

        # Kết hợp các Mask
        if sub_masks:
            if "VÀ" in logic_mode:
                for sm in sub_masks: mask &= sm
            else:
                or_m = sub_masks[0]
                for sm in sub_masks[1:]: or_m |= sm
                mask &= or_m

    # ── PHẦN 3: KẾT QUẢ & THỐNG KÊ ──────────────────────────────────────────
    df_final = df_raw[mask]
    
    st.write("---")
    st.markdown("### 📊 Kết quả sau khi lọc")
    
    res_c1, res_c2, res_c3, res_c4 = st.columns(4)
    res_c1.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_raw):,}</div><div class="stat-label">Tổng hàng gốc</div></div>', unsafe_allow_html=True)
    res_c2.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_final):,}</div><div class="stat-label">Hàng thỏa điều kiện</div></div>', unsafe_allow_html=True)
    res_c3.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#cf222e">{len(df_raw)-len(df_final):,}</div><div class="stat-label">Đã bị loại bỏ</div></div>', unsafe_allow_html=True)
    ratio = (len(df_final)/len(df_raw)*100) if len(df_raw)>0 else 0
    res_c4.markdown(f'<div class="stat-box"><div class="stat-num">{ratio:.1f}%</div><div class="stat-label">Tỷ lệ giữ lại</div></div>', unsafe_allow_html=True)

    if not df_final.empty:
        st.write("")
        st.download_button(
            label=f"📥 Tải xuống file Excel đã lọc ({len(df_final)} hàng)",
            data=to_excel(df_final),
            file_name="du_lieu_da_loc.xl
