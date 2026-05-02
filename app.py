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
    .info-card { background: #ffffff; border: 1px solid #e1e4e8; border-radius: 8px; padding: 10px; margin-bottom: 10px; }
    .col-name { color: #2e86de; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý tối ưu ────────────────────────────────────────────────────────
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    df = pd.read_excel(file_bytes, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all").reset_index(drop=True)

@st.cache_data
def get_row_colors(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    wb = load_workbook(file_bytes, read_only=False, data_only=True)
    ws = wb.active
    colors = []
    for row in ws.iter_rows(min_row=header_idx + 2, min_col=1, max_col=1):
        cell = row[0]
        color_hex = "No Fill"
        if cell.fill and cell.fill.start_color:
            rgb = cell.fill.start_color.rgb
            if isinstance(rgb, str) and len(rgb) == 8: color_hex = f"#{rgb[2:]}"
        colors.append(color_hex)
    return colors

def to_excel(df):
    output = io.BytesIO()
    if '__color__' in df.columns: df = df.drop(columns=['__color__'])
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ── Header ──────────────────────────────────────────────────────────────────
st.markdown('<div class="main-title">📊 Excel Filter & Manager Pro</div>', unsafe_allow_html=True)

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Cài đặt dữ liệu")
    uploaded_file = st.file_uploader("Tải file .xlsx", type=["xlsx"])
    h_row = st.number_input("Dòng tiêu đề (Header):", min_value=1, value=1)
    st.divider()
    enable_color = st.checkbox("🌈 Kích hoạt lọc màu")

# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    file_content = uploaded_file.getvalue()
    df_raw = get_clean_data(file_content, h_row - 1)
    
    if enable_color:
        with st.spinner("Đang quét màu..."):
            colors = get_row_colors(file_content, h_row - 1)
            df_raw['__color__'] = colors[:len(df_raw)]
    else:
        df_raw['__color__'] = "No Fill"

    # ── PHẦN MỚI: BẢNG TỔNG QUAN THÔNG TIN FILE ──────────────────────────────
    with st.expander("📋 BẢNG THÔNG TIN CÁC CỘT TRONG FILE", expanded=True):
        st.write("Dựa vào đây để bạn chọn chính xác giá trị cần lọc:")
        
        info_cols = st.columns(3)
        for i, col in enumerate(df_raw.columns):
            if col == "__color__": continue
            
            with info_cols[i % 3]:
                unique_vals = df_raw[col].dropna().unique()
                count_unique = len(unique_vals)
                dtype = "Số" if pd.api.types.is_numeric_dtype(df_raw[col]) else "Chữ"
                
                # Hiển thị tóm tắt từng cột
                st.markdown(f"""
                <div class="info-card">
                    <div class="col-name">📍 {col}</div>
                    <div style="font-size:0.85rem; color:#666;">
                        Kiểu: {dtype} | Duy nhất: {count_unique}<br>
                        Gợi ý: {", ".join(map(str, unique_vals[:5]))}{"..." if count_unique > 5 else ""}
                    </div>
                </div>
                """, unsafe_allow_html=True)

    # ── PHẦN LỌC DỮ LIỆU ─────────────────────────────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    
    t1, t2, t3 = st.tabs(["Lọc Nội dung", "Lọc Màu sắc", "Logic NẾU-THÌ"])
    mask = pd.Series([True] * len(df_raw))

    with t1:
        logic_mode = st.radio("Kết hợp điều kiện:", ["VÀ (AND)", "HOẶC (OR)"], horizontal=True)
        sel_cols = st.multiselect("Chọn cột muốn lọc:", [c for c in df_raw.columns if c != '__color__'])
        
        if sel_cols:
            c_ui = st.columns(2)
            sub_masks = []
            for i, col in enumerate(sel_cols):
                with c_ui[i % 2]:
                    # Nâng cấp: Nếu cột có ít giá trị duy nhất, dùng dropdown cho dễ chọn
                    u_vals = df_raw[col].dropna().unique()
                    if len(u_vals) < 50:
                        val = st.multiselect(f"Chọn giá trị trong [{col}]:", options=u_vals, key=f"sel_{col}")
                        if val:
                            sub_masks.append(df_raw[col].isin(val))
                    else:
                        val = st.text_input(f"Nhập từ khóa tìm trong [{col}]:", key=f"v_{col}")
                        if val:
                            sub_masks.append(df_raw[col].astype(str).str.contains(val, case=False, na=False))
            
            if sub_masks:
                if logic_mode == "VÀ (AND)":
                    for sm in sub_masks: mask &= sm
                else:
                    or_m = sub_masks[0]
                    for sm in sub_masks[1:]: or_m |= sm
                    mask &= or_m

    with t2:
        if enable_color:
            u_colors = df_raw['__color__'].unique().tolist()
            sel_colors = st.multiselect("Chọn màu giữ lại:", u_colors, default=u_colors)
            mask &= df_raw['__color__'].isin(sel_colors)
        else:
            st.info("Hãy tích vào 'Kích hoạt lọc màu' ở thanh bên trái.")

    with t3:
        use_it = st.checkbox("Kích hoạt logic NẾU - THÌ")
        if use_it:
            ic1, ic2, ic3 = st.columns(3)
            if_col = ic1.selectbox("NẾU Cột", df_raw.columns)
            if_val = ic2.text_input("Có giá trị chứa", key="if_v")
            then_col = ic3.selectbox("THÌ Cột đó chứa", df_raw.columns)
            then_val = st.text_input("Giá trị phải thỏa mãn", key="then_v")
            if if_val and then_val:
                cond_a = df_raw[if_col].astype(str).str.contains(if_val, case=False, na=False)
                cond_b = df_raw[then_col].astype(str).str.contains(then_val, case=False, na=False)
                mask &= (~cond_a | cond_b)

    # ── THỐNG KÊ KẾT QUẢ ────────────────────────────────────────────────────
    df_final = df_raw[mask]
    
    st.write("---")
    st.subheader("📊 Kết quả xử lý")
    
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_raw):,}</div><div class="stat-label">Hàng gốc</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_final):,}</div><div class="stat-label">Hàng thỏa ĐK</div></div>', unsafe_allow_html=True)
    with c3:
        diff = len(df_raw) - len(df_final)
        st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#cf222e">{diff:,}</div><div class="stat-label">Đã loại bỏ</div></div>', unsafe_allow_html=True)
    with c4:
        ratio = (len(df_final)/len(df_raw)*100) if len(df_raw)>0 else 0
        st.markdown(f'<div class="stat-box"><div class="stat-num">{ratio:.1f}%</div><div class="stat-label">Tỷ lệ giữ</div></div>', unsafe_allow_html=True)

    if not df_final.empty:
        st.write("")
        st.download_button(
            label=f"📥 Tải file Excel kết quả ({len(df_final)} hàng)",
            data=to_excel(df_final),
            file_name="ket_qua_loc.xlsx",
            type="primary",
            use_container_width=True
        )
        st.dataframe(df_final.drop(columns=['__color__'], errors='ignore'), use_container_width=True)
    else:
        st.warning("Không có dữ liệu thỏa mãn.")

else:
    st.info("👋 Chào mừng! Hãy tải file Excel lên để bắt đầu.")
