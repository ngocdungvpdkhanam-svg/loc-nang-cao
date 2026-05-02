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
    .filter-section { background: #f8f9fa; border: 1px solid #e1e4e8; border-radius: 10px; padding: 15px; margin-top: 10px; }
    .highlight-text { color: #2e86de; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý dữ liệu ────────────────────────────────────────────────────────
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

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Cài đặt dữ liệu")
    uploaded_file = st.file_uploader("Tải file .xlsx", type=["xlsx"])
    h_row = st.number_input("Dòng tiêu đề (Header):", min_value=1, value=1)
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

    # ── PHẦN 1: TỔNG QUAN FILE ──────────────────────────────────────────────
    with st.expander("📋 THÔNG TIN CHI TIẾT CÁC CỘT", expanded=True):
        display_cols = [c for c in df_raw.columns if c != '__color__']
        info_cols = st.columns(4)
        for i, col in enumerate(display_cols):
            with info_cols[i % 4]:
                nulls = df_raw[col].isna().sum()
                st.markdown(f'<div style="background:white; padding:10px; border-radius:5px; border:1px solid #ddd; margin-bottom:5px;"><b style="color:#2e86de;">{col}</b><br><small>Trống: {nulls}</small></div>', unsafe_allow_html=True)

    # ── PHẦN 2: BỘ LỌC CẢI TIẾN ──────────────────────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    t1, t2, t3 = st.tabs(["Lọc Nội dung", "Lọc Màu sắc", "Logic NẾU-THÌ"])
    mask = pd.Series([True] * len(df_raw))

    with t1:
        logic_mode = st.radio("Kết hợp điều kiện:", ["VÀ (AND)", "HOẶC (OR)"], horizontal=True)
        sel_cols = st.multiselect("Chọn các cột muốn lọc:", display_cols)
        
        if sel_cols:
            sub_masks = []
            for col in sel_cols:
                # Tạo một vùng bao quanh mỗi cột để dễ nhìn
                st.markdown(f'<div class="filter-section">', unsafe_allow_html=True)
                st.markdown(f"📍 Đang lọc cột: <span class='highlight-text'>{col}</span>", unsafe_allow_html=True)
                
                # --- NÚT ĐIỀU KHIỂN NHANH ---
                c_btn1, c_btn2, c_btn3 = st.columns([1, 1, 1.5])
                with c_btn1:
                    is_all = st.checkbox("✅ Chọn tất cả", key=f"all_{col}", value=False)
                with c_btn2:
                    is_blanks = st.checkbox("⚪ Chỉ lấy (Blanks)", key=f"blk_{col}", value=False)
                
                # Lấy dữ liệu thực tế
                unique_vals = sorted([str(x) for x in df_raw[col].dropna().unique()])
                options = ["(Blanks)"] + unique_vals if df_raw[col].isna().any() else unique_vals
                
                # Xác định giá trị mặc định dựa trên nút bấm
                if is_blanks:
                    default_sel = ["(Blanks)"]
                elif is_all:
                    default_sel = options
                else:
                    default_sel = []

                # --- Ô CHỌN GIÁ TRỊ ---
                selected = st.multiselect(
                    f"Chọn giá trị cho [{col}]:",
                    options=options,
                    default=default_sel,
                    key=f"ms_{col}"
                )
                
                # Tạo bộ lọc (Mask) cho cột này
                if not selected:
                    sub_masks.append(pd.Series([False] * len(df_raw)))
                elif len(selected) == len(options):
                    sub_masks.append(pd.Series([True] * len(df_raw)))
                else:
                    local_mask = pd.Series([False] * len(df_raw))
                    if "(Blanks)" in selected:
                        local_mask |= df_raw[col].isna()
                        others = [v for v in selected if v != "(Blanks)"]
                        if others: local_mask |= df_raw[col].astype(str).isin(others)
                    else:
                        local_mask |= df_raw[col].astype(str).isin(selected)
                    sub_masks.append(local_mask)
                
                st.markdown('</div>', unsafe_allow_html=True)

            if sub_masks:
                if "VÀ" in logic_mode:
                    for sm in sub_masks: mask &= sm
                else:
                    or_m = sub_masks[0]
                    for sm in sub_masks[1:]: or_m |= sm
                    mask &= or_m

    # ... (Các phần Tab 2, Tab 3 giữ nguyên)
    with t2:
        if enable_color:
            u_colors = df_raw['__color__'].unique().tolist()
            sel_colors = st.multiselect("Chọn màu:", u_colors, default=u_colors)
            mask &= df_raw['__color__'].isin(sel_colors)
        else: st.info("Bật lọc màu ở sidebar.")
        
    with t3:
        if st.checkbox("Dùng Logic NẾU-THÌ"):
            c_if, c_v, c_then, c_tv = st.columns(4)
            if_c = c_if.selectbox("NẾU Cột", display_cols)
            if_v = c_v.text_input("Chứa giá trị", key="ifv")
            then_c = c_then.selectbox("THÌ Cột", display_cols)
            then_v = c_tv.text_input("Giá trị thỏa", key="thenv")
            if if_v and then_v:
                m_a = df_raw[if_c].astype(str).str.contains(if_v, case=False, na=False)
                m_b = df_raw[then_c].astype(str).str.contains(then_v, case=False, na=False)
                mask &= (~m_a | m_b)

    # ── PHẦN 3: KẾT QUẢ XỬ LÝ ──────────────────────────────────────────────
    df_final = df_raw[mask]
    
    st.write("---")
    st.markdown("### 📊 Kết quả xử lý")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_raw):,}</div><div class="stat-label">Gốc</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_final):,}</div><div class="stat-label">Khớp</div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#cf222e">{len(df_raw)-len(df_final):,}</div><div class="stat-label">Đã loại</div></div>', unsafe_allow_html=True)
    ratio = (len(df_final)/len(df_raw)*100) if len(df_raw)>0 else 0
    c4.markdown(f'<div class="stat-box"><div class="stat-num">{ratio:.1f}%</div><div class="stat-label">Tỷ lệ giữ</div></div>', unsafe_allow_html=True)

    if not df_final.empty:
        st.write("")
        st.download_button(label=f"📥 Tải Excel kết quả ({len(df_final)} hàng)", data=to_excel(df_final), file_name="ket_qua.xlsx", type="primary", use_container_width=True)
        st.dataframe(df_final.drop(columns=['__color__'], errors='ignore'), use_container_width=True)
    else:
        st.warning("Không có dữ liệu thỏa mãn.")
else:
    st.info("👋 Hãy tải file Excel lên để bắt đầu.")
