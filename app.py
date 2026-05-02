import streamlit as st
import pandas as pd
import io

# ── Cấu hình trang ──────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Excel Filter Pro",
    page_icon="📊",
    layout="wide",
)

# ── CSS Tùy chỉnh (Giao diện hiện đại) ───────────────────────────────────────
st.markdown("""
<style>
    .main-title { font-size: 2.5rem; font-weight: 800; color: #1E3A5F; margin-bottom: 0; }
    .sub-title { color: #666; margin-bottom: 2rem; }
    .stat-card {
        background: #FFFFFF; border: 1px solid #E1E4E8; border-radius: 10px;
        padding: 15px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stat-val { font-size: 1.8rem; font-weight: 700; color: #2E86DE; }
    .stat-label { font-size: 0.8rem; color: #586069; text-transform: uppercase; }
    div[data-testid="stExpander"] { border: 1px solid #E1E4E8; border-radius: 8px; background: #F8F9FA; }
</style>
""", unsafe_allow_html=True)

# ── Khởi tạo Session State ──────────────────────────────────────────────────
if "df_original" not in st.session_state:
    st.session_state.df_original = None
if "df_current" not in st.session_state:
    st.session_state.df_current = None
if "conditions" not in st.session_state:
    st.session_state.conditions = []
if "history" not in st.session_state:
    st.session_state.history = []

# ── Hàm hỗ trợ ──────────────────────────────────────────────────────────────
def reset_app():
    st.session_state.df_original = None
    st.session_state.df_current = None
    st.session_state.conditions = []
    st.session_state.history = []

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def build_mask(df, conditions, logic):
    if not conditions:
        return pd.Series([True] * len(df), index=df.index)
    
    masks = []
    for cond in conditions:
        col, op, val = cond["col"], cond["op"], cond["val"]
        try:
            series = df[col]
            if op == "==":
                m = series.astype(str).isin([str(v) for v in val]) if isinstance(val, list) else series.astype(str) == str(val)
            elif op == "!=":
                m = ~series.astype(str).isin([str(v) for v in val]) if isinstance(val, list) else series.astype(str) != str(val)
            elif op == "chứa":
                m = series.astype(str).str.contains(str(val), case=False, na=False)
            elif op == "không chứa":
                m = ~series.astype(str).str.contains(str(val), case=False, na=False)
            elif op == ">":
                m = pd.to_numeric(series, errors='coerce') > float(val)
            elif op == ">=":
                m = pd.to_numeric(series, errors='coerce') >= float(val)
            elif op == "<":
                m = pd.to_numeric(series, errors='coerce') < float(val)
            elif op == "<=":
                m = pd.to_numeric(series, errors='coerce') <= float(val)
            elif op == "trống":
                m = series.isna() | (series.astype(str).str.strip() == "")
            elif op == "không trống":
                m = ~(series.isna() | (series.astype(str).str.strip() == ""))
            else:
                m = pd.Series([True] * len(df), index=df.index)
            masks.append(m)
        except Exception as e:
            st.error(f"Lỗi lọc tại cột {col}: {e}")
            masks.append(pd.Series([False] * len(df), index=df.index))

    result = masks[0]
    for m in masks[1:]:
        if logic == "VÀ (AND)":
            result = result & m
        else:
            result = result | m
    return result

# ── Giao diện Header ────────────────────────────────────────────────────────
st.markdown('<p class="main-title">📊 Excel Filter & Manager</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Công cụ lọc, xóa và quản lý dữ liệu Excel chuyên nghiệp</p>', unsafe_allow_html=True)

# ── Sidebar: Upload file ───────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Tải lên dữ liệu")
    uploaded_file = st.file_uploader("Chọn file Excel (.xlsx, .xls) hoặc CSV", type=["xlsx", "xls", "csv"])
    
    if uploaded_file:
        try:
            # RESET FILE POINTER ĐỂ ĐẢM BẢO ĐỌC ĐƯỢC FILE
            uploaded_file.seek(0)
            
            if uploaded_file.name.endswith('.csv'):
                df_raw = pd.read_csv(uploaded_file)
                if st.session_state.df_original is None:
                    st.session_state.df_original = df_raw.copy()
                    st.session_state.df_current = df_raw.copy()
            else:
                # Đọc Excel
                xls = pd.ExcelFile(uploaded_file)
                sheet = st.selectbox("Chọn Sheet", xls.sheet_names)
                
                if st.button("🚀 Tải dữ liệu từ Sheet này") or st.session_state.df_original is None:
                    # Tự động tìm hàng header (hàng đầu tiên có ít nhất 2 cột dữ liệu)
                    uploaded_file.seek(0)
                    df_temp = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, nrows=10)
                    header_idx = 0
                    for i, row in df_temp.iterrows():
                        if row.dropna().count() >= 2:
                            header_idx = i
                            break
                    
                    uploaded_file.seek(0)
                    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_idx)
                    # Làm sạch tên cột
                    df_raw.columns = [str(c).strip().replace("\n", " ") for c in df_raw.columns]
                    df_raw = df_raw.dropna(how="all").reset_index(drop=True)
                    
                    st.session_state.df_original = df_raw.copy()
                    st.session_state.df_current = df_raw.copy()
                    st.session_state.conditions = []
                    st.rerun()

            st.success(f"✅ Đã tải: {len(st.session_state.df_original):,} hàng")
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")

    if st.button("🔄 Làm mới ứng dụng (Reset All)", use_container_width=True):
        reset_app()
        st.rerun()

# ── Nội dung chính ─────────────────────────────────────────────────────────
if st.session_state.df_current is not None:
    df = st.session_state.df_current
    
    # Chỉ số thống kê
    s1, s2, s3, s4 = st.columns(4)
    with s1: st.markdown(f'<div class="stat-card"><div class="stat-val">{len(st.session_state.df_original):,}</div><div class="stat-label">Gốc</div></div>', unsafe_allow_html=True)
    with s2: st.markdown(f'<div class="stat-card"><div class="stat-val">{len(df):,}</div><div class="stat-label">Hiện tại</div></div>', unsafe_allow_html=True)
    with s3: st.markdown(f'<div class="stat-card"><div class="stat-val" style="color:#D93025">{len(st.session_state.df_original)-len(df):,}</div><div class="stat-label">Đã xóa</div></div>', unsafe_allow_html=True)
    with s4: st.markdown(f'<div class="stat-card"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Số cột</div></div>', unsafe_allow_html=True)

    st.write("")
    
    # ── KHU VỰC LỌC DỮ LIỆU ────────────────────────────────────────────────
    st.subheader("🔍 Lọc dữ liệu thông minh")
    
    with st.expander("➕ Thêm điều kiện mới", expanded=True):
        c1, c2, c3 = st.columns([2, 1, 2])
        with c1:
            col_target = st.selectbox("Chọn cột lọc", df.columns)
        with c2:
            is_num = pd.api.types.is_numeric_dtype(df[col_target])
            ops = [">", ">=", "<", "<=", "==", "!=", "chứa", "không chứa", "trống", "không trống"]
            op_target = st.selectbox("Phép toán", ops)
        with c3:
            if "trống" in op_target:
                val_target = None
                st.text_input("Giá trị", value="N/A", disabled=True)
            elif op_target in ["==", "!="]:
                unique_vals = sorted(df[col_target].astype(str).unique()[:100])
                val_target = st.multiselect("Chọn giá trị", unique_vals)
            else:
                val_target = st.text_input("Nhập giá trị")

        if st.button("✨ Thêm điều kiện", use_container_width=True):
            if op_target not in ["trống", "không trống"] and not val_target and val_target != 0:
                st.warning("Vui lòng nhập/chọn giá trị lọc!")
            else:
                st.session_state.conditions.append({"col": col_target, "op": op_target, "val": val_target})
                st.rerun()

    # Hiển thị danh sách điều kiện
    if st.session_state.conditions:
        st.info("💡 Các điều kiện đang được áp dụng:")
        logic = st.radio("Sử dụng logic:", ["VÀ (AND)", "HOẶC (OR)"], horizontal=True)
        
        for i, cond in enumerate(st.session_state.conditions):
            cc1, cc2 = st.columns([5, 1])
            cc1.code(f"Điều kiện {i+1}: [{cond['col']}] {cond['op']} {cond['val'] if cond['val'] else ''}")
            if cc2.button("🗑️", key=f"del_{i}"):
                st.session_state.conditions.pop(i)
                st.rerun()

        # Xem trước kết quả lọc
        mask = build_mask(df, st.session_state.conditions, logic)
        matched_df = df[mask]
        
        st.markdown(f"**Kết quả lọc:** Tìm thấy `{len(matched_df)}` hàng thỏa mãn.")
        
        tab_preview, tab_action = st.tabs(["👁️ Xem trước dữ liệu thỏa ĐK", "⚡ Hành động"])
        
        with tab_preview:
            st.dataframe(matched_df.head(100), use_container_width=True)
            
        with tab_action:
            act1, act2, act3 = st.columns(3)
            if act1.button(f"🗑️ XÓA {len(matched_df)} hàng này", use_container_width=True, type="primary"):
                st.session_state.df_current = df[~mask].reset_index(drop=True)
                st.session_state.conditions = []
                st.session_state.history.append(f"Đã xóa {len(matched_df)} hàng")
                st.rerun()
            
            if act2.button(f"✅ CHỈ GIỮ {len(matched_df)} hàng này", use_container_width=True):
                st.session_state.df_current = matched_df.reset_index(drop=True)
                st.session_state.conditions = []
                st.session_state.history.append(f"Đã giữ lại {len(matched_df)} hàng")
                st.rerun()
                
            if act3.button("🧹 Hủy bỏ lọc", use_container_width=True):
                st.session_state.conditions = []
                st.rerun()

    st.divider()

    # ── QUẢN LÝ CỘT & XUẤT FILE ───────────────────────────────────────────
    c_left, c_right = st.columns(2)
    
    with c_left:
        with st.expander("🛠️ Quản lý cột (Xóa/Đổi tên)"):
            col_to_del = st.multiselect("Chọn cột muốn xóa", df.columns)
            if st.button("Xác nhận xóa cột"):
                st.session_state.df_current = df.drop(columns=col_to_del)
                st.rerun()
            
            st.divider()
            old_name = st.selectbox("Chọn cột đổi tên", df.columns)
            new_name = st.text_input("Tên mới")
            if st.button("Đổi tên"):
                st.session_state.df_current = df.rename(columns={old_name: new_name})
                st.rerun()

    with c_right:
        with st.expander("💾 Xuất dữ liệu hiện tại", expanded=True):
            file_format = st.radio("Định dạng file:", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True)
            
            if file_format == "Excel (.xlsx)":
                data_bytes = to_excel(df)
                st.download_button(
                    label="📥 Tải file Excel",
                    data=data_bytes,
                    file_name="ket_qua_loc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                csv_data = df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 Tải file CSV",
                    data=csv_data,
                    file_name="ket_qua_loc.csv",
                    mime="text/csv",
                    use_container_width=True
                )

    # Hiển thị toàn bộ bảng hiện tại
    st.subheader("📄 Bảng dữ liệu hiện tại")
    st.dataframe(df, use_container_width=True, height=400)

else:
    # Màn hình chờ
    st.info("👋 Chào mừng! Vui lòng tải lên file Excel hoặc CSV ở thanh bên trái để bắt đầu xử lý dữ liệu.")
    st.image("https://img.freepik.com/free-vector/data-extraction-concept-illustration_114360-4766.jpg", width=400)

st.caption("Excel Filter Pro v2.0 • Hỗ trợ xử lý dữ liệu lớn • Dev by Streamlit")
