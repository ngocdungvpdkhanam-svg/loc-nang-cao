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
    .info-card { background: #ffffff; border: 1px solid #e1e4e8; border-radius: 8px; padding: 10px; margin-bottom: 10px; min-height: 100px; }
    .col-name { color: #2e86de; font-weight: 700; font-size: 0.9rem; }
    .x-count { color: #1e7e34; font-weight: bold; background: #e6ffed; padding: 2px 5px; border-radius: 4px; }
    .blank-count { color: #d73a49; font-weight: bold; background: #ffeef0; padding: 2px 5px; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý dữ liệu ────────────────────────────────────────────────────────
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    df = pd.read_excel(file_bytes, header=header_idx)
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
    st.write("---")
    st.caption("Phiên bản v4.0 - Thống kê giá trị X")

# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    file_content = uploaded_file.getvalue()
    df_raw = get_clean_data(file_content, h_row - 1)
    display_cols = df_raw.columns.tolist()

    # ── PHẦN 1: TỔNG QUAN CHI TIẾT (Có đếm giá trị X) ──────────────────────────
    with st.expander("📊 BẢNG THÔNG TIN CỘT & ĐẾM GIÁ TRỊ X", expanded=True):
        info_cols = st.columns(4)
        for i, col in enumerate(display_cols):
            with info_cols[i % 4]:
                nulls = df_raw[col].isna().sum()
                # Đếm số lượng X (không phân biệt X hoa hay x thường)
                x_count = df_raw[col].astype(str).str.strip().str.upper().eq('X').sum()
                st.markdown(f'''
                <div class="info-card">
                    <div class="col-name">📍 {col}</div>
                    <div style="margin-top:5px; font-size:0.8rem;">
                        Trống: <span class="blank-count">{nulls}</span><br>
                        Số lượng X: <span class="x-count">{x_count}</span>
                    </div>
                </div>''', unsafe_allow_html=True)

    # ── PHẦN 2: CÀI ĐẶT BỘ LỌC ──────────────────────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    
    t1, t2 = st.tabs(["Lọc Nội dung & X", "Logic NẾU-THÌ"])
    mask = pd.Series([True] * len(df_raw))

    with t1:
        logic_mode = st.radio("Kết hợp điều kiện:", ["VÀ (AND)", "HOẶC (OR)"], horizontal=True)
        sel_cols = st.multiselect("BƯỚC 1: Chọn cột bạn muốn lọc:", display_cols)
        
        if sel_cols:
            st.write("BƯỚC 2: Chọn kiểu lọc nhanh:")
            sub_masks = []
            for col in sel_cols:
                st.write(f"---")
                st.markdown(f"📍 Đang thiết lập cho cột: **{col}**")
                
                # --- LỰA CHỌN LỌC NHANH ---
                c1, c2, c3 = st.columns(3)
                with c1:
                    f_type = st.radio(f"Kiểu lọc [{col}]:", 
                                     ["Tất cả", "Chỉ lấy ô TRỐNG", "Chỉ lấy giá trị X", "Chọn từ danh sách"],
                                     key=f"type_{col}")
                
                if f_type == "Chỉ lấy ô TRỐNG":
                    sub_masks.append(df_raw[col].isna())
                elif f_type == "Chỉ lấy giá trị X":
                    sub_masks.append(df_raw[col].astype(str).str.strip().str.upper() == 'X')
                elif f_type == "Chọn từ danh sách":
                    # Lấy danh sách giá trị thực tế
                    options = sorted([str(x) for x in df_raw[col].dropna().unique()])
                    selected = st.multiselect(f"Chọn giá trị cụ thể trong [{col}]:", options=options, key=f"ms_{col}")
                    if selected:
                        sub_masks.append(df_raw[col].astype(str).isin(selected))
                    else:
                        sub_masks.append(pd.Series([False] * len(df_raw)))
                else:
                    sub_masks.append(pd.Series([True] * len(df_raw)))

            if sub_masks:
                if "VÀ" in logic_mode:
                    for sm in sub_masks: mask &= sm
                else:
                    or_m = sub_masks[0]
                    for sm in sub_masks[1:]: or_m |= sm
                    mask &= or_m

    with t2:
        if st.checkbox("Kích hoạt Logic NẾU-THÌ"):
            ic1, ic2, ic3, ic4 = st.columns(4)
            if_c = ic1.selectbox("NẾU Cột", display_cols)
            if_v = ic2.text_input("Có giá trị là", value="X", key="ifv")
            then_c = ic3.selectbox("THÌ Cột đó chứa", display_cols)
            then_v = ic4.text_input("Giá trị thỏa", key="thenv")
            if if_v and then_v:
                c_a = df_raw[if_c].astype(str).str.strip().str.upper() == str(if_v).upper()
                c_b = df_raw[then_c].astype(str).str.contains(then_v, case=False, na=False)
                mask &= (~c_a | c_b)

    # ── PHẦN 3: KẾT QUẢ XỬ LÝ ──────────────────────────────────────────────
    df_final = df_raw[mask]
    
    st.write("---")
    st.markdown("### 📊 Kết quả sau khi lọc")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_raw):,}</div><div class="stat-label">Tổng hàng gốc</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_final):,}</div><div class="stat-label">Thỏa điều kiện</div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#cf222e">{len(df_raw)-len(df_final):,}</div><div class="stat-label">Đã loại bỏ</div></div>', unsafe_allow_html=True)
    ratio = (len(df_final)/len(df_raw)*100) if len(df_raw)>0 else 0
    c4.markdown(f'<div class="stat-box"><div class="stat-num">{ratio:.1f}%</div><div class="stat-label">Tỷ lệ giữ</div></div>', unsafe_allow_html=True)

    if not df_final.empty:
        st.write("")
        st.download_button(
            label=f"📥 Tải Excel kết quả ({len(df_final)} hàng)",
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
