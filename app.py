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
    .filter-card { background: #ffffff; border: 1px solid #d1d5db; border-radius: 10px; padding: 20px; margin-top: 10px; border-top: 5px solid #2e86de; }
    .blank-tag { background: #ff4b4b; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 0.8rem; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý dữ liệu ────────────────────────────────────────────────────────
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    # Đọc dữ liệu với dòng tiêu đề tùy chỉnh
    df = pd.read_excel(file_bytes, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    # Chuyển đổi các ô trông có vẻ trống thành chuẩn rỗng của hệ thống
    df = df.replace(r'^\s*$', pd.NA, regex=True)
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
    h_row = st.number_input("Dòng tiêu đề (Header):", min_value=1, value=6) # Mặc định dòng 6 như ảnh bạn chụp
    
# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    file_content = uploaded_file.getvalue()
    df_raw = get_clean_data(file_content, h_row - 1)
    display_cols = df_raw.columns.tolist()

    # ── PHẦN 1: THÔNG TIN CHI TIẾT CÁC CỘT ──────────────────────────────────
    with st.expander("📋 XEM THÔNG TIN CHI TIẾT CÁC CỘT", expanded=True):
        info_cols = st.columns(4)
        for i, col in enumerate(display_cols):
            with info_cols[i % 4]:
                null_count = df_raw[col].isna().sum()
                st.markdown(f'''
                <div style="background:white; padding:8px; border-radius:5px; border:1px solid #ddd; margin-bottom:5px;">
                    <b style="color:#2e86de;">{col}</b><br>
                    <small>Số ô trống: <span style="color:{"red" if null_count > 0 else "green"}">{null_count}</span></small>
                </div>''', unsafe_allow_html=True)

    # ── PHẦN 2: BỘ LỌC CẢI TIẾN ──────────────────────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    
    t1, t2, t3 = st.tabs(["Lọc Nội dung", "Lọc Màu sắc", "Logic NẾU-THÌ"])
    mask = pd.Series([True] * len(df_raw))

    with t1:
        logic_mode = st.radio("Kết hợp điều kiện:", ["VÀ (Khớp tất cả)", "HOẶC (Khớp 1 trong các)"], horizontal=True)
        sel_cols = st.multiselect("Chọn các cột muốn lọc:", display_cols)
        
        if sel_cols:
            sub_masks = []
            for col in sel_cols:
                # Tạo một khung trắng riêng cho mỗi cột
                st.markdown(f'<div class="filter-card">', unsafe_allow_html=True)
                st.markdown(f"📍 Đang lọc cột: **{col}**")
                
                # --- CÁCH CHỌN MỚI: CHIA LÀM 3 CHẾ ĐỘ ---
                filter_type = st.radio(
                    f"Kiểu lọc cho cột {col}:",
                    ["Lấy tất cả giá trị", "Chỉ lấy ô TRỐNG (Blanks)", "Chọn giá trị cụ thể"],
                    key=f"type_{col}",
                    horizontal=True
                )

                if filter_type == "Chỉ lấy ô TRỐNG (Blanks)":
                    st.error(f"Đã chọn: Chỉ hiển thị các hàng mà cột [{col}] bị rỗng.")
                    sub_masks.append(df_raw[col].isna())
                
                elif filter_type == "Chọn giá trị cụ thể":
                    # Lấy danh sách giá trị có sẵn (bỏ qua NaN)
                    options = sorted([str(x) for x in df_raw[col].dropna().unique()])
                    
                    c1, c2 = st.columns([3, 1])
                    with c2:
                        all_btn = st.checkbox("Chọn tất cả", key=f"all_{col}")
                    
                    with c1:
                        default_val = options if all_btn else []
                        selected = st.multiselect(f"Chọn từ danh sách [{col}]:", options=options, default=default_val, key=f"ms_{col}")
                    
                    if selected:
                        sub_masks.append(df_raw[col].astype(str).isin(selected))
                    else:
                        # Nếu chọn "Cụ thể" nhưng không chọn gì -> Trả về không có hàng nào (giống bỏ tích hết trong Excel)
                        sub_masks.append(pd.Series([False] * len(df_raw)))
                
                else: # Lấy tất cả
                    sub_masks.append(pd.Series([True] * len(df_raw)))
                
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
    st.markdown("### 📊 Kết quả sau khi lọc")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_raw):,}</div><div class="stat-label">Hàng gốc</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_final):,}</div><div class="stat-label">Khớp điều kiện</div></div>', unsafe_allow_html=True)
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
        st.warning("⚠️ Không có dữ liệu thỏa mãn điều kiện bạn chọn.")
else:
    st.info("👋 Hãy tải file Excel lên để bắt đầu.")
