import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Logic Filter Pro", layout="wide")

# --- Hàm hỗ trợ xử lý dữ liệu ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def apply_condition(series, op, val):
    """Trả về một mảng True/False dựa trên phép toán"""
    try:
        s = series.astype(str)
        if op == "==": return s == str(val)
        if op == "!=": return s != str(val)
        if op == "chứa": return s.str.contains(str(val), case=False, na=False)
        if op == "không chứa": return ~s.str.contains(str(val), case=False, na=False)
        if op == "bắt đầu bằng": return s.str.startswith(str(val), na=False)
        
        # Các phép toán số học
        num_s = pd.to_numeric(series, errors='coerce')
        if op == ">": return num_s > float(val)
        if op == ">=": return num_s >= float(val)
        if op == "<": return num_s < float(val)
        if op == "<=": return num_s <= float(val)
        
        if op == "trống": return series.isna() | (s.str.strip() == "")
        if op == "không trống": return ~(series.isna() | (s.str.strip() == ""))
    except:
        return pd.Series([False] * len(series))
    return pd.Series([True] * len(series))

# --- Giao diện ---
st.title("🚀 Trình Quản Lý Logic Excel: VÀ, HOẶC, NẾU-THÌ")

with st.sidebar:
    st.header("⚙️ Nhập dữ liệu")
    uploaded_file = st.file_uploader("Tải file Excel/CSV", type=["xlsx", "csv"])
    header_row = st.number_input("Dòng tiêu đề:", min_value=1, value=1)
    
if uploaded_file:
    # Đọc dữ liệu
    @st.cache_data
    def load_data(file, h):
        file.seek(0)
        return pd.read_csv(file, header=h-1) if file.name.endswith(".csv") else pd.read_excel(file, header=h-1)
    
    df_raw = load_data(uploaded_file, header_row)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    # --- PHẦN 1: LỌC CƠ BẢN (VÀ / HOẶC) ---
    st.subheader("1️⃣ Bộ lọc Cơ bản (VÀ / HOẶC)")
    
    col_logic = st.radio("Kết hợp các điều kiện bên dưới theo kiểu:", ["Tất cả đều đúng (VÀ)", "Chỉ cần một cái đúng (HOẶC)"], horizontal=True)
    
    selected_cols = st.multiselect("Chọn các cột muốn lọc:", options=df_raw.columns.tolist(), key="basic_cols")
    
    basic_masks = []
    if selected_cols:
        for col in selected_cols:
            c1, c2, c3 = st.columns([1, 1, 2])
            with c1: st.info(f"Cột: {col}")
            with c2: 
                op = st.selectbox("Phép toán", [">", ">=", "<", "<=", "==", "!=", "chứa", "trống", "không trống"], key=f"b_op_{col}")
            with c3:
                val = st.text_input("Giá trị", key=f"b_val_{col}") if "trống" not in op else ""
            
            mask = apply_condition(df_raw[col], op, val)
            basic_masks.append(mask)

    # --- PHẦN 2: LOGIC NẾU - THÌ (IF-THEN) ---
    st.subheader("2️⃣ Bộ lọc Quy tắc (NẾU - THÌ)")
    st.caption("Ví dụ: NẾU [Cột A] chứa 'Hà Nội' THÌ [Cột B] phải lớn hơn 1000. (Nếu không thỏa thì hàng bị loại)")
    
    with st.expander("Cài đặt quy tắc NẾU - THÌ"):
        use_if_then = st.checkbox("Kích hoạt logic NẾU - THÌ")
        if_mask = pd.Series([True] * len(df_raw))
        
        if use_if_then:
            ic1, ic2, ic3 = st.columns([1, 1, 1.5])
            with ic1: if_col = st.selectbox("NẾU Cột", df_raw.columns, key="if_col")
            with ic2: if_op = st.selectbox("Có điều kiện", ["==", "chứa", ">", "<", "không trống"], key="if_op")
            with ic3: if_val = st.text_input("Giá trị là", key="if_val")
            
            tc1, tc2, tc3 = st.columns([1, 1, 1.5])
            with tc1: then_col = st.selectbox("THÌ Cột đó (hoặc cột khác)", df_raw.columns, key="then_col")
            with tc2: then_op = st.selectbox("Phải thỏa mãn", ["==", ">", "<", "chứa", "không trống"], key="then_op")
            with tc3: then_val = st.text_input("Giá trị thỏa mãn", key="then_val")
            
            # Logic: IF A THEN B  <=>  (NOT A) OR (B)
            cond_a = apply_condition(df_raw[if_col], if_op, if_val)
            cond_b = apply_condition(df_raw[then_col], then_op, then_val)
            if_mask = (~cond_a) | cond_b

    # --- TỔNG HỢP LOGIC ---
    final_mask = pd.Series([True] * len(df_raw))
    
    # Áp dụng VÀ/HOẶC
    if basic_masks:
        if "VÀ" in col_logic:
            for m in basic_masks: final_mask &= m
        else:
            or_mask = basic_masks[0]
            for m in basic_masks[1:]: or_mask |= m
            final_mask &= or_mask
            
    # Áp dụng NẾU-THÌ
    final_mask &= if_mask

    df_final = df_raw[final_mask]

    # --- KẾT QUẢ ---
    st.divider()
    r1, r2, r3 = st.columns(3)
    r1.metric("Tổng gốc", len(df_raw))
    r2.metric("Sau khi lọc", len(df_final))
    r3.download_button("📥 Tải file kết quả", to_excel(df_final), "ket_qua.xlsx", type="primary", use_container_width=True)

    st.dataframe(df_final, use_container_width=True)

else:
    st.info("👋 Chào mừng! Hãy tải file Excel/CSV lên để bắt đầu lập trình bộ lọc.")
