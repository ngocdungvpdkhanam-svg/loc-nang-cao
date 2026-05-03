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
    .filter-card { background: #ffffff; border: 1px solid #d1d5db; border-radius: 10px; padding: 20px; margin-top: 10px; border-top: 5px solid #ff4b4b; }
    .blank-label { color: #ff4b4b; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý dữ liệu (Cực kỳ quan trọng) ──────────────────────────────────────
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    # 1. Đọc dữ liệu
    df = pd.read_excel(file_bytes, header=header_idx)
    
    # 2. Làm sạch tên cột
    df.columns = [str(c).strip() if pd.notnull(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]
    
    # 3. KỸ THUẬT QUAN TRỌNG: Xử lý ô trống
    # Biến tất cả: NaN, Chuỗi rỗng "", Chuỗi toàn dấu cách "   " thành nhãn "(Trống - Blanks)"
    def clean_blanks(val):
        if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan":
            return "(Trống - Blanks)"
        return str(val).strip()

    for col in df.columns:
        df[col] = df[col].apply(clean_blanks)
        
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
    st.write("---")
    st.info("💡 Lưu ý: Hệ thống đã tự động gộp tất cả ô rỗng/trắng vào nhãn **(Trống - Blanks)** để bạn dễ chọn.")

# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    file_content = uploaded_file.getvalue()
    df_raw = get_clean_data(file_content, h_row - 1)
    display_cols = df_raw.columns.tolist()

    # Kiểm tra nếu file đọc ra bị lỗi cột Unnamed quá nhiều
    if "Unnamed: 0" in display_cols:
        st.warning("⚠️ Dòng tiêu đề có vẻ chưa đúng, bạn hãy thử chỉnh lại 'Dòng tiêu đề' ở bên trái.")

    # ── PHẦN 1: TỔNG QUAN FILE ──────────────────────────────────────────────
    with st.expander("📋 KIỂM TRA CỘT & Ô TRỐNG", expanded=True):
        info_cols = st.columns(4)
        for i, col in enumerate(display_cols):
            with info_cols[i % 4]:
                blank_count = (df_raw[col] == "(Trống - Blanks)").sum()
                st.markdown(f'''
                <div style="background:white; padding:8px; border-radius:5px; border:1px solid #eee; margin-bottom:5px;">
                    <b style="color:#2e86de;">{col}</b><br>
                    <small>Số ô trống: <span class="blank-label">{blank_count}</span></small>
                </div>''', unsafe_allow_html=True)

    # ── PHẦN 2: BỘ LỌC CẢI TIẾN ──────────────────────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    
    t1, t2, t3 = st.tabs(["Lọc Nội dung (Chống lỗi)", "Lọc Màu sắc", "Logic NẾU-THÌ"])
    mask = pd.Series([True] * len(df_raw))

    with t1:
        logic_mode = st.radio("Kết hợp điều kiện:", ["VÀ (AND)", "HOẶC (OR)"], horizontal=True)
        sel_cols = st.multiselect("Chọn các cột bạn muốn lọc:", display_cols)
        
        if sel_cols:
            sub_masks = []
            for col in sel_cols:
                st.markdown(f'<div class="filter-card">', unsafe_allow_html=True)
                st.markdown(f"📍 Đang lọc cột: **{col}**")
                
                # Lấy danh sách giá trị (Nhãn Trống sẽ luôn xuất hiện nếu có)
                options = sorted(df_raw[col].unique().tolist())
                
                # ĐƯA NHÃN TRỐNG LÊN ĐẦU DANH SÁCH CHO DỄ CHỌN
                if "(Trống - Blanks)" in options:
                    options.remove("(Trống - Blanks)")
                    options = ["(Trống - Blanks)"] + options

                c1, c2 = st.columns([3, 1])
                with c2:
                    is_all = st.checkbox("✅ Chọn tất cả", key=f"all_{col}", value=False)
                
                with c1:
                    default_sel = options if is_all else []
                    selected = st.multiselect(
                        f"Chọn giá trị cho [{col}]:",
                        options=options,
                        default=default_sel,
                        key=f"ms_{col}"
                    )
                
                # Tạo bộ lọc (Mask)
                if selected:
                    sub_masks.append(df_raw[col].isin(selected))
                else:
                    # Nếu không chọn gì -> Lọc mất sạch (giống Excel)
                    sub_masks.append(pd.Series([False] * len(df_raw)))
                
                st.markdown('</div>', unsafe_allow_html=True)

            if sub_masks:
                if "VÀ" in logic_mode:
                    for sm in sub_masks: mask &= sm
                else:
                    or_m = sub_masks[0]
                    for sm in sub_masks[1:]: or_m |= sm
                    mask &= or_m

    # ... (Các tab khác giữ nguyên)

    # ── PHẦN 3: KẾT QUẢ XỬ LÝ ──────────────────────────────────────────────
    df_final = df_raw[mask]
    
    # TRƯỚC KHI HIỂN THỊ: Biến ngược nhãn "(Trống - Blanks)" về rỗng để file tải về đẹp
    df_final_export = df_final.replace("(Trống - Blanks)", "")

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
            label=f"📥 Tải Excel kết quả ({len(df_final)} hàng)",
            data=to_excel(df_final_export),
            file_name="ket_qua_loc.xlsx",
            type="primary",
            use_container_width=True
        )
        st.dataframe(df_final_export, use_container_width=True)
    else:
        st.warning("⚠️ Không có dữ liệu thỏa mãn điều kiện.")
else:
    st.info("👋 Hãy tải file Excel lên để bắt đầu.")
