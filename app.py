import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl import load_workbook

st.set_page_config(page_title="Excel Color & Logic Filter", layout="wide")

# --- Hàm hỗ trợ lấy màu sắc từ file Excel ---
def get_df_with_colors(file, header_idx):
    file.seek(0)
    wb = load_workbook(file, data_only=True)
    ws = wb.active # Lấy sheet đầu tiên
    
    data = []
    colors = []
    
    # Chuyển worksheet thành list các hàng
    rows = list(ws.rows)
    if not rows:
        return pd.DataFrame(), []

    # Lấy tiêu đề
    header_row = [str(cell.value).strip() if cell.value else f"Col{i}" for i, cell in enumerate(rows[header_idx])]
    
    # Duyệt qua các hàng dữ liệu (sau hàng tiêu đề)
    for row in rows[header_idx + 1:]:
        row_values = [cell.value for cell in row]
        data.append(row_values)
        
        # Lấy màu sắc của ô đầu tiên trong hàng làm "Màu hàng" 
        # (Hoặc bạn có thể tùy chỉnh lấy màu của ô cụ thể)
        fill = row[0].fill
        color_hex = "No Fill"
        if fill and fill.start_color and fill.start_color.index != '00000000':
            rgb = fill.start_color.rgb
            if isinstance(rgb, str) and len(rgb) == 8: # ARGB format
                color_hex = f"#{rgb[2:]}" # Chuyển về Hex tiêu chuẩn #RRGGBB
        colors.append(color_hex)
        
    df = pd.DataFrame(data, columns=header_row)
    df['__row_color__'] = colors
    unique_colors = list(set(colors))
    return df, unique_colors

def to_excel(df):
    output = io.BytesIO()
    # Loại bỏ cột màu nội bộ trước khi xuất file
    df_export = df.drop(columns=['__row_color__'], errors='ignore')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False)
    return output.getvalue()

def apply_condition(series, op, val):
    try:
        s = series.astype(str)
        if op == "==": return s == str(val)
        if op == "chứa": return s.str.contains(str(val), case=False, na=False)
        num_s = pd.to_numeric(series, errors='coerce')
        if op == ">": return num_s > float(val)
        if op == "<": return num_s < float(val)
        if op == "trống": return series.isna() | (s.str.strip() == "")
        if op == "không trống": return ~(series.isna() | (s.str.strip() == ""))
    except: return pd.Series([False] * len(series))
    return pd.Series([True] * len(series))

# --- GIAO DIỆN CHÍNH ---
st.title("🎨 Excel Filter: Màu Sắc + Logic Đa Tầng")

with st.sidebar:
    st.header("⚙️ Nhập dữ liệu")
    uploaded_file = st.file_uploader("Tải file Excel (Chỉ hỗ trợ .xlsx)", type=["xlsx"])
    header_row_num = st.number_input("Dòng tiêu đề:", min_value=1, value=1)

if uploaded_file:
    # Đọc dữ liệu kèm màu sắc
    df_raw, unique_colors = get_df_with_colors(uploaded_file, header_row_num - 1)
    
    if not df_raw.empty:
        # --- PHẦN 1: LỌC THEO MÀU SẮC ---
        st.subheader("1️⃣ Lọc theo màu nền (Fill Color)")
        
        c1, c2 = st.columns([1, 3])
        with c1:
            st.write("Bảng màu tìm thấy:")
            # Hiển thị bảng màu nhỏ để người dùng nhận diện
            for c in unique_colors:
                if c != "No Fill":
                    st.markdown(f'<div style="background-color:{c}; width:100%; height:20px; border:1px solid #ccc; margin-bottom:5px; border-radius:3px; text-align:center; font-size:10px;">{c}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div style="background-color:#fff; width:100%; height:20px; border:1px solid #ccc; margin-bottom:5px; border-radius:3px; text-align:center; font-size:10px;">Không màu</div>', unsafe_allow_html=True)

        with c2:
            selected_colors = st.multiselect("Chọn các màu muốn GIỮ LẠI:", options=unique_colors, default=unique_colors)
            color_mask = df_raw['__row_color__'].isin(selected_colors)

        st.divider()

        # --- PHẦN 2: LỌC LOGIC (VÀ/HOẶC/NẾU-THÌ) ---
        st.subheader("2️⃣ Lọc theo nội dung & Logic")
        
        tab1, tab2 = st.tabs(["Bộ lọc Cơ bản (VÀ/HOẶC)", "Bộ lọc Quy tắc (NẾU-THÌ)"])
        
        final_mask = color_mask.copy()

        with tab1:
            logic_type = st.radio("Kiểu kết hợp:", ["VÀ (Khớp tất cả)", "HOẶC (Khớp 1 trong các)"], horizontal=True)
            sel_cols = st.multiselect("Chọn cột cần lọc nội dung:", options=[c for c in df_raw.columns if c != '__row_color__'])
            
            basic_masks = []
            for col in sel_cols:
                cc1, cc2, cc3 = st.columns([1, 1, 1])
                with cc1: st.caption(f"Cột: {col}")
                with cc2: op = st.selectbox("Phép toán", ["==", "chứa", ">", "<", "trống", "không trống"], key=f"b_op_{col}")
                with cc3: val = st.text_input("Giá trị", key=f"b_val_{col}") if "trống" not in op else ""
                basic_masks.append(apply_condition(df_raw[col], op, val))
            
            if basic_masks:
                if "VÀ" in logic_type:
                    for m in basic_masks: final_mask &= m
                else:
                    or_m = basic_masks[0]
                    for m in basic_masks[1:]: or_m |= m
                    final_mask &= or_m

        with tab2:
            use_if_then = st.checkbox("Kích hoạt logic NẾU - THÌ")
            if use_if_then:
                ic1, ic2, ic3 = st.columns(3)
                with ic1: if_col = st.selectbox("NẾU Cột", df_raw.columns, key="if_col")
                with ic2: if_op = st.selectbox("Điều kiện", ["==", "chứa", ">", "<"], key="if_op")
                with ic3: if_val = st.text_input("Giá trị", key="if_val")
                
                tc1, tc2, tc3 = st.columns(3)
                with tc1: then_col = st.selectbox("THÌ Cột", df_raw.columns, key="then_col")
                with tc2: then_op = st.selectbox("Phải là", ["==", "chứa", ">", "<"], key="then_op")
                with tc3: then_val = st.text_input("Giá trị thỏa", key="then_val")
                
                cond_a = apply_condition(df_raw[if_col], if_op, if_val)
                cond_b = apply_condition(df_raw[then_col], then_op, then_val)
                final_mask &= (~cond_a | cond_b)

        # --- KẾT QUẢ ---
        df_final = df_raw[final_mask]

        st.divider()
        res1, res2, res3 = st.columns(3)
        res1.metric("Tổng gốc", len(df_raw))
        res2.metric("Kết quả lọc", len(df_final))
        res3.download_button("📥 Tải file kết quả (Excel)", to_excel(df_final), "result_filtered.xlsx", type="primary", use_container_width=True)

        # Hiển thị bảng kèm màu sắc giả lập ở cột cuối
        st.dataframe(df_final, use_container_width=True)
        
    else:
        st.error("Không thể đọc dữ liệu. Hãy kiểm tra dòng tiêu đề.")
else:
    st.info("👋 Hãy tải file Excel (.xlsx) để trải nghiệm tính năng lọc theo màu sắc và logic.")
