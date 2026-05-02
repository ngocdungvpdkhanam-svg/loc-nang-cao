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
    .filter-container { background: #ffffff; border: 1px solid #dee2e6; border-radius: 12px; padding: 20px; margin-top: 15px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
    .filter-header { color: #2e86de; font-weight: 800; font-size: 1.1rem; border-bottom: 2px solid #f0f2f6; margin-bottom: 15px; padding-bottom: 5px; }
    .blank-button { background-color: #fff1f0; border: 1px solid #ffa39e; color: #cf222e; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý dữ liệu ────────────────────────────────────────────────────────
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    df = pd.read_excel(file_bytes, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    # Đồng bộ hóa tất cả các kiểu ô trống về chuẩn (None)
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    return df.dropna(how="all").reset_index(drop=True)

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="main-title" style="font-size:1.5rem;">📁 Nhập dữ liệu</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Tải file .xlsx", type=["xlsx"])
    h_row = st.number_input("Dòng tiêu đề (Header):", min_value=1, value=6)

# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    file_content = uploaded_file.getvalue()
    df_raw = get_clean_data(file_content, h_row - 1)
    display_cols = df_raw.columns.tolist()

    # ── PHẦN 1: THÔNG TIN CHI TIẾT CÁC CỘT ──────────────────────────────────
    with st.expander("📋 TỔNG QUAN DỮ LIỆU CÁC CỘT", expanded=True):
        info_cols = st.columns(4)
        for i, col in enumerate(display_cols):
            with info_cols[i % 4]:
                nulls = df_raw[col].isna().sum()
                st.markdown(f'''
                <div style="background:white; padding:10px; border-radius:8px; border:1px solid #eee; margin-bottom:10px;">
                    <small style="color:#888;">Cột:</small><br><b>{col}</b><br>
                    <small style="color:{"red" if nulls > 0 else "green"}">Số ô trống: {nulls}</small>
                </div>''', unsafe_allow_html=True)

    # ── PHẦN 2: BỘ LỌC KIỂU MỚI (CHỐNG TRÔI) ────────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    
    sel_cols = st.multiselect("BƯỚC 1: Chọn các cột bạn muốn thực hiện lọc:", display_cols)
    
    mask = pd.Series([True] * len(df_raw))

    if sel_cols:
        st.info("BƯỚC 2: Tùy chỉnh giá trị cho từng cột đã chọn:")
        logic_mode = st.radio("Kết hợp logic giữa các cột:", ["VÀ (Thỏa mãn tất cả)", "HOẶC (Chỉ cần thỏa 1 cái)"], horizontal=True)
        
        sub_masks = []
        for col in sel_cols:
            with st.container():
                st.markdown(f'<div class="filter-container">', unsafe_allow_html=True)
                st.markdown(f'<div class="filter-header">📍 CỘT: {col}</div>', unsafe_allow_html=True)
                
                # --- PHÂN CHIA CHẾ ĐỘ LỌC RÕ RÀNG ---
                # Dùng radio để người dùng chọn hẳn chế độ, không bị lẫn
                mode = st.radio(
                    f"Chọn chế độ lọc cho [{col}]:",
                    ["Giữ lại tất cả", "CHỈ LẤY Ô TRỐNG (Blanks)", "Chọn theo danh sách cụ thể"],
                    key=f"mode_{col}",
                    horizontal=True
                )

                if mode == "CHỈ LẤY Ô TRỐNG (Blanks)":
                    st.warning(f"Đã kích hoạt: Chỉ lấy các dòng mà cột '{col}' đang để TRỐNG.")
                    sub_masks.append(df_raw[col].isna())
                
                elif mode == "Chọn theo danh sách cụ thể":
                    # Lấy danh sách giá trị (không bao gồm ô trống)
                    unique_vals = sorted([str(x) for x in df_raw[col].dropna().unique()])
                    
                    # Thêm nút bấm hỗ trợ chọn nhanh
                    c1, c2 = st.columns([3, 1])
                    with c2:
                        select_all = st.checkbox("Chọn tất cả", key=f"all_{col}")
                    
                    with c1:
                        default_val = unique_vals if select_all else []
                        selected = st.multiselect(
                            f"Danh sách giá trị của [{col}]:",
                            options=unique_vals,
                            default=default_val,
                            key=f"ms_{col}"
                        )
                    
                    if selected:
                        sub_masks.append(df_raw[col].astype(str).isin(selected))
                    else:
                        # Nếu chọn "Cụ thể" mà không tích gì -> Bị coi là lọc mất hết (giống Excel)
                        sub_masks.append(pd.Series([False] * len(df_raw)))
                
                else: # Giữ lại tất cả
                    sub_masks.append(pd.Series([True] * len(df_raw)))
                
                st.markdown('</div>', unsafe_allow_html=True)

        # Tổng hợp các Mask
        if sub_masks:
            if "VÀ" in logic_mode:
                for sm in sub_masks: mask &= sm
            else:
                or_m = sub_masks[0]
                for sm in sub_masks[1:]: or_m |= sm
                mask &= or_m

    # ── PHẦN 3: KẾT QUẢ XỬ LÝ (STYLE BOXES) ──────────────────────────────────
    df_final = df_raw[mask]
    
    st.write("---")
    st.markdown("### 📊 Kết quả thống kê")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_raw):,}</div><div class="stat-label">Tổng hàng gốc</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_final):,}</div><div class="stat-label">Hàng thỏa điều kiện</div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#cf222e">{len(df_raw)-len(df_final):,}</div><div class="stat-label">Đã loại bỏ</div></div>', unsafe_allow_html=True)
    ratio = (len(df_final)/len(df_raw)*100) if len(df_raw)>0 else 0
    c4.markdown(f'<div class="stat-box"><div class="stat-num">{ratio:.1f}%</div><div class="stat-label">Tỷ lệ giữ lại</div></div>', unsafe_allow_html=True)

    if not df_final.empty:
        st.write("")
        st.download_button(
            label=f"📥 Tải xuống file kết quả ({len(df_final)} hàng)",
            data=to_excel(df_final),
            file_name="ket_qua_loc_du_lieu.xlsx",
            type="primary",
            use_container_width=True
        )
        st.dataframe(df_final, use_container_width=True)
    else:
        st.warning("⚠️ Không có dữ liệu nào khớp với các điều kiện lọc bạn đã chọn.")
else:
    st.info("👋 Chào bạn! Hãy tải file Excel lên để bắt đầu.")
