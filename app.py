import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Advanced Filter", layout="wide")

# --- Hàm hỗ trợ ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

@st.cache_data
def get_data(file, p_header):
    try:
        file.seek(0)
        if file.name.endswith(".csv"):
            return pd.read_csv(file, header=p_header)
        else:
            return pd.read_excel(file, header=p_header)
    except Exception as e:
        st.error(f"Lỗi: {e}")
        return None

# --- Giao diện chính ---
st.title("📊 Bộ Lọc Excel Đa Điều Kiện")

with st.sidebar:
    st.header("⚙️ Cấu hình")
    uploaded_file = st.file_uploader("Tải file Excel/CSV", type=["xlsx", "csv"])
    header_row = st.number_input("Dòng chứa tiêu đề:", min_value=1, value=1)
    pandas_header = header_row - 1

if uploaded_file:
    df_raw = get_data(uploaded_file, pandas_header)

    if df_raw is not None:
        df_raw = df_raw.dropna(how="all").reset_index(drop=True)
        # Làm sạch tên cột (xóa khoảng trắng thừa)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        
        st.sidebar.success(f"✅ Đã tải {len(df_raw)} hàng")

        # --- KHU VỰC THIẾT LẬP BỘ LỌC ---
        st.subheader("🔍 Cài đặt điều kiện lọc")
        
        selected_cols = st.multiselect(
            "Chọn các cột muốn áp dụng điều kiện:", 
            options=df_raw.columns.tolist()
        )

        # Lưu trữ các điều kiện lọc
        filters = []

        if selected_cols:
            # Hiển thị mỗi cột đã chọn thành một hàng điều kiện
            for col in selected_cols:
                c1, c2, c3 = st.columns([1.5, 1, 2])
                
                with c1:
                    st.write(f"**Cột: {col}**")
                
                with c2:
                    # Xác định kiểu dữ liệu để gợi ý phép toán
                    is_numeric = pd.api.types.is_numeric_dtype(df_raw[col])
                    if is_numeric:
                        ops = ["==", "!=", ">", ">=", "<", "<=", "trống", "không trống"]
                    else:
                        ops = ["chứa", "không chứa", "==", "!=", "bắt đầu bằng", "kết thúc bằng", "trống", "không trống"]
                    
                    op = st.selectbox(f"Phép toán ({col})", ops, key=f"op_{col}", label_visibility="collapsed")
                
                with c3:
                    if op in ["trống", "không trống"]:
                        val = None
                        st.text_input(f"Giá trị ({col})", value="N/A", disabled=True, label_visibility="collapsed")
                    else:
                        val = st.text_input(f"Nhập giá trị cần lọc cho {col}", key=f"val_{col}", label_visibility="collapsed")
                
                if val or op in ["trống", "không trống"]:
                    filters.append({"col": col, "op": op, "val": val})

        # --- THỰC THI LỌC DỮ LIỆU ---
        df_filtered = df_raw.copy()

        for f in filters:
            col, op, val = f['col'], f['op'], f['val']
            
            try:
                if op == "==":
                    df_filtered = df_filtered[df_filtered[col].astype(str) == str(val)]
                elif op == "!=":
                    df_filtered = df_filtered[df_filtered[col].astype(str) != str(val)]
                elif op == "chứa":
                    df_filtered = df_filtered[df_filtered[col].astype(str).str.contains(str(val), case=False, na=False)]
                elif op == "không chứa":
                    df_filtered = df_filtered[~df_filtered[col].astype(str).str.contains(str(val), case=False, na=False)]
                elif op == "bắt đầu bằng":
                    df_filtered = df_filtered[df_filtered[col].astype(str).str.startswith(str(val), na=False)]
                elif op == "kết thúc bằng":
                    df_filtered = df_filtered[df_filtered[col].astype(str).str.endswith(str(val), na=False)]
                elif op == ">":
                    df_filtered = df_filtered[pd.to_numeric(df_filtered[col], errors='coerce') > float(val)]
                elif op == ">=":
                    df_filtered = df_filtered[pd.to_numeric(df_filtered[col], errors='coerce') >= float(val)]
                elif op == "<":
                    df_filtered = df_filtered[pd.to_numeric(df_filtered[col], errors='coerce') < float(val)]
                elif op == "<=":
                    df_filtered = df_filtered[pd.to_numeric(df_filtered[col], errors='coerce') <= float(val)]
                elif op == "trống":
                    df_filtered = df_filtered[df_filtered[col].isna() | (df_filtered[col].astype(str).str.strip() == "")]
                elif op == "không trống":
                    df_filtered = df_filtered[~(df_filtered[col].isna() | (df_filtered[col].astype(str).str.strip() == ""))]
            except Exception as e:
                st.warning(f"Không thể áp dụng '{op}' cho cột '{col}'. Vui lòng kiểm tra định dạng dữ liệu.")

        st.divider()

        # --- HIỂN THỊ KẾT QUẢ ---
        res_c1, res_c2, res_c3 = st.columns([1, 1, 2])
        res_c1.metric("Gốc", len(df_raw))
        res_c2.metric("Sau lọc", len(df_filtered))
        
        with res_c3:
            if len(df_filtered) > 0:
                st.download_button(
                    label="📥 Tải file kết quả (Excel)",
                    data=to_excel(df_filtered),
                    file_name="ket_qua_loc.xlsx",
                    use_container_width=True,
                    type="primary"
                )

        st.dataframe(df_filtered.head(200), use_container_width=True)
        if len(df_filtered) > 200:
            st.caption(f"Đang hiển thị 200 trên tổng số {len(df_filtered)} hàng.")

else:
    st.info("👋 Chào bạn! Hãy tải file Excel/CSV lên để bắt đầu sử dụng các bộ lọc nâng cao.")
