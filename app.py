import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# ── Cấu hình giao diện (Đúng chuẩn Stat Boxes bạn thích) ──────────────────────
st.set_page_config(page_title="Excel Pro Manager", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 2.2rem; font-weight: 700; background: linear-gradient(135deg, #1e3a5f, #2e86de); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 0.5rem; }
    .stat-box { background: #f0f4ff; border-radius: 10px; padding: 15px; text-align: center; border-left: 5px solid #2e86de; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .stat-num { font-size: 1.8rem; font-weight: 700; color: #1e3a5f; }
    .stat-label { font-size: 0.8rem; color: #555; text-transform: uppercase; font-weight: 600; }
    .filter-card { background: #ffffff; border: 1px solid #d1d5db; border-radius: 10px; padding: 20px; margin-top: 15px; border-top: 5px solid #2e86de; }
    .blank-tag { background: #ff4b4b; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

# ── Hàm xử lý dữ liệu (Làm sạch ô trống) ──────────────────────────────────────
@st.cache_data
def get_clean_data(file_content, header_idx):
    file_bytes = io.BytesIO(file_content)
    # Đọc dữ liệu từ dòng tiêu đề được chọn (ví dụ dòng 6)
    df = pd.read_excel(file_bytes, header=header_idx)
    # Làm sạch tên cột
    df.columns = [str(c).strip() if pd.notnull(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]
    # Chuẩn hóa: Biến mọi kiểu "Trống" (NaN, dấu cách) thành một giá trị thực sự dễ lọc
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
    h_row = st.number_input("Dòng tiêu đề (Header):", min_value=1, value=6)
    st.info(f"💡 Đang dùng dòng {h_row} làm tiêu đề.")

# ── Xử lý chính ─────────────────────────────────────────────────────────────
if uploaded_file:
    file_content = uploaded_file.getvalue()
    df_raw = get_clean_data(file_content, h_row - 1)
    display_cols = [c for c in df_raw.columns if not c.startswith("__")]

    # ── PHẦN 1: TỔNG QUAN CHI TIẾT ──────────────────────────────────────────
    with st.expander("📋 XEM THÔNG TIN CHI TIẾT CÁC CỘT", expanded=True):
        info_cols = st.columns(4)
        for i, col in enumerate(display_cols):
            with info_cols[i % 4]:
                nulls = df_raw[col].isna().sum()
                st.markdown(f'''
                <div style="background:white; padding:8px; border-radius:5px; border:1px solid #ddd; margin-bottom:5px;">
                    <b style="color:#2e86de;">{col}</b><br>
                    <small>Số ô trống: <span style="color:{"red" if nulls > 0 else "green"}">{nulls}</span></small>
                </div>''', unsafe_allow_html=True)

    # ── PHẦN 2: BỘ LỌC THÔNG MINH (THIẾT KẾ MỚI) ──────────────────────────────
    st.subheader("🔍 Cài đặt bộ lọc")
    
    t1, t2, t3 = st.tabs(["Lọc Nội dung (Ưu tiên Trống/Tất cả)", "Lọc Màu sắc", "Logic NẾU-THÌ"])
    mask = pd.Series([True] * len(df_raw))

    with t1:
        logic_mode = st.radio("Kết hợp điều kiện giữa các cột:", ["VÀ (AND)", "HOẶC (OR)"], horizontal=True)
        sel_cols = st.multiselect("BƯỚC 1: Chọn các cột bạn muốn lọc:", display_cols)
        
        if sel_cols:
            st.write("BƯỚC 2: Thiết lập giá trị lọc:")
            sub_masks = []
            for col in sel_cols:
                # Tạo khung lọc riêng cho từng cột
                st.markdown(f'<div class="filter-card">', unsafe_allow_html=True)
                st.markdown(f"📍 Đang lọc cột: **{col}**")
                
                # --- NÚT BẤM CHỌN NHANH (KHÔNG CẦN TÌM KIẾM) ---
                c1, c2, c3 = st.columns([1, 1, 1])
                with c1:
                    is_all = st.checkbox(f"✅ Chọn tất cả (Select All)", key=f"all_{col}", value=False)
                with c2:
                    is_blank = st.checkbox(f"⚪ Chỉ lấy hàng TRỐNG (Blanks)", key=f"blk_{col}", value=False)
                
                # Lấy danh sách giá trị thực (bỏ qua NaN)
                options = sorted([str(x) for x in df_raw[col].dropna().unique()])
                
                # Xác định danh sách mặc định dựa trên nút bấm
                default_sel = options if is_all else []
                
                # Ô lọc giá trị (Sẽ tự động điền nếu bấm Chọn tất cả)
                selected = st.multiselect(
                    f"Hoặc chọn giá trị cụ thể của [{col}]:",
                    options=options,
                    default=default_sel,
                    key=f"ms_{col}"
                )
                
                # LOGIC LỌC (MASK)
                local_mask = pd.Series([False] * len(df_raw))
                
                if is_blank:
                    local_mask |= df_raw[col].isna()
                
                if selected:
                    local_mask |= df_raw[col].astype(str).isin(selected)
                
                if not is_blank and not selected and not is_all:
                    # Nếu không chọn gì cả -> Ẩn hết hàng của cột đó (Giống Excel)
                    local_mask = pd.Series([False] * len(df_raw))
                elif is_all:
                    # Nếu chọn tất cả -> Hiện hết (bao gồm cả ô trống nếu có tích)
                    local_mask |= pd.Series([True] * len(df_raw)) if is_blank else df_raw[col].notna()

                sub_masks.append(local_mask)
                st.markdown('</div>', unsafe_allow_html=True)

            if sub_masks:
                if "VÀ" in logic_mode:
                    for sm in sub_masks: mask &= sm
                else:
                    or_m = sub_masks[0]
                    for sm in sub_masks[1:]: or_m |= sm
                    mask &= or_m

    # ... (Các tab Màu sắc và Logic giữ nguyên như bản trước)
    with t2: st.info("Tính năng lọc màu yêu cầu 'Kích hoạt lọc màu' ở Sidebar.")
    with t3: st.caption("Logic NẾU - THÌ dùng để xử lý ràng buộc nâng cao.")

    # ── PHẦN 3: KẾT QUẢ XỬ LÝ (4 Ô STAT BOXES) ──────────────────────────────
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
        st.warning("⚠️ Không có hàng nào thỏa mãn điều kiện bạn đã chọn.")
else:
    st.info("👋 Hãy tải file Excel lên để bắt đầu. Hãy kiểm tra 'Dòng tiêu đề' ở bên trái nhé!")
