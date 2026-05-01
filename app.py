import streamlit as st
import pandas as pd
import io
from copy import deepcopy

st.set_page_config(
    page_title="Excel Filter & Manager",
    page_icon="📊",
    layout="wide",
)

# ── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem; font-weight: 700;
        background: linear-gradient(135deg, #1e3a5f, #2e86de);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem;
    }
    .subtitle { color: #666; font-size: 0.95rem; margin-bottom: 1.5rem; }
    .stat-box {
        background: #f0f4ff; border-radius: 10px; padding: 14px 20px;
        text-align: center; border-left: 4px solid #2e86de;
    }
    .stat-num { font-size: 1.8rem; font-weight: 700; color: #1e3a5f; }
    .stat-label { font-size: 0.78rem; color: #555; }
    .cond-card {
        background: #fafbff; border: 1px solid #dde4f0;
        border-radius: 8px; padding: 10px 14px; margin-bottom: 6px;
    }
    .result-keep { color: #1a7f37; font-weight: 600; }
    .result-remove { color: #cf222e; font-weight: 600; }
    div[data-testid="stExpander"] { border: 1px solid #dde4f0 !important; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown('<div class="main-title">📊 Excel Filter & Manager</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Tải lên file Excel → Chọn điều kiện lọc nhiều cột → Giữ lại hoặc xóa dữ liệu → Tải xuống</div>', unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────
for key, default in {
    "df_original": None,
    "df_current": None,
    "conditions": [],
    "history": [],
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


def reset_conditions():
    st.session_state.conditions = []


def get_unique_values(df, col):
    try:
        return sorted(df[col].dropna().unique().tolist(), key=lambda x: str(x))
    except Exception:
        return []


def build_mask(df, conditions, logic):
    """Return a boolean Series: True = row matches ALL/ANY conditions."""
    if not conditions:
        return pd.Series([True] * len(df), index=df.index)

    masks = []
    for cond in conditions:
        col = cond["col"]
        op = cond["op"]
        val = cond["val"]
        dtype = df[col].dtype

        try:
            if op == "==":
                m = df[col].astype(str).isin([str(v) for v in val]) if isinstance(val, list) else df[col] == val
            elif op == "!=":
                m = ~(df[col].astype(str).isin([str(v) for v in val])) if isinstance(val, list) else df[col] != val
            elif op == "chứa":
                m = df[col].astype(str).str.contains(str(val), case=False, na=False)
            elif op == "không chứa":
                m = ~df[col].astype(str).str.contains(str(val), case=False, na=False)
            elif op == ">":
                m = pd.to_numeric(df[col], errors="coerce") > float(val)
            elif op == ">=":
                m = pd.to_numeric(df[col], errors="coerce") >= float(val)
            elif op == "<":
                m = pd.to_numeric(df[col], errors="coerce") < float(val)
            elif op == "<=":
                m = pd.to_numeric(df[col], errors="coerce") <= float(val)
            elif op == "trống":
                m = df[col].isna() | (df[col].astype(str).str.strip() == "")
            elif op == "không trống":
                m = ~(df[col].isna() | (df[col].astype(str).str.strip() == ""))
            else:
                m = pd.Series([True] * len(df), index=df.index)
        except Exception:
            m = pd.Series([False] * len(df), index=df.index)

        masks.append(m)

    if logic == "VÀ (AND)":
        result = masks[0]
        for m in masks[1:]:
            result = result & m
    else:  # OR
        result = masks[0]
        for m in masks[1:]:
            result = result | m
    return result


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Kết quả")
    return buf.getvalue()


# ── File upload ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Tải lên file")
    uploaded = st.file_uploader("Chọn file Excel (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"])

    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df_raw = pd.read_csv(uploaded)
            else:
                sheet_names = pd.ExcelFile(uploaded).sheet_names
                selected_sheet = st.selectbox("Chọn sheet", sheet_names)
                df_raw = pd.read_excel(uploaded, sheet_name=selected_sheet)

            if st.session_state.df_original is None or st.button("🔄 Tải lại file"):
                st.session_state.df_original = df_raw.copy()
                st.session_state.df_current = df_raw.copy()
                st.session_state.conditions = []
                st.session_state.history = []
                st.success(f"✅ Đã tải: {len(df_raw):,} hàng × {len(df_raw.columns)} cột")
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")

    st.divider()
    if st.session_state.df_current is not None:
        st.header("🕐 Lịch sử thao tác")
        if st.session_state.history:
            for i, h in enumerate(reversed(st.session_state.history[-5:])):
                st.caption(f"• {h}")
        else:
            st.caption("Chưa có thao tác nào.")

        if st.button("↩️ Khôi phục dữ liệu gốc", use_container_width=True):
            st.session_state.df_current = st.session_state.df_original.copy()
            st.session_state.conditions = []
            st.session_state.history.append("Khôi phục dữ liệu gốc")
            st.rerun()


# ── Main area ─────────────────────────────────────────────────────────────
if st.session_state.df_current is None:
    st.info("👈 Vui lòng tải lên file Excel ở thanh bên trái để bắt đầu.")
    st.stop()

df = st.session_state.df_current
orig = st.session_state.df_original

# Stats row
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f'<div class="stat-box"><div class="stat-num">{len(orig):,}</div><div class="stat-label">Hàng gốc</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="stat-box"><div class="stat-num">{len(df):,}</div><div class="stat-label">Hàng hiện tại</div></div>', unsafe_allow_html=True)
with c3:
    diff = len(orig) - len(df)
    st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#cf222e">{diff:,}</div><div class="stat-label">Đã xóa</div></div>', unsafe_allow_html=True)
with c4:
    st.markdown(f'<div class="stat-box"><div class="stat-num">{len(df.columns)}</div><div class="stat-label">Số cột</div></div>', unsafe_allow_html=True)

st.divider()

# ── Condition builder ──────────────────────────────────────────────────────
st.subheader("🔍 Xây dựng điều kiện lọc")

col_left, col_right = st.columns([2, 1])

with col_left:
    with st.expander("➕ Thêm điều kiện mới", expanded=True):
        cc1, cc2, cc3 = st.columns([2, 1.5, 2])
        with cc1:
            chosen_col = st.selectbox("Chọn cột", df.columns.tolist(), key="new_col")
        with cc2:
            dtype = df[chosen_col].dtype
            if pd.api.types.is_numeric_dtype(dtype):
                ops = ["==", "!=", ">", ">=", "<", "<=", "trống", "không trống"]
            else:
                ops = ["==", "!=", "chứa", "không chứa", "trống", "không trống"]
            chosen_op = st.selectbox("Phép so sánh", ops, key="new_op")
        with cc3:
            if chosen_op in ("trống", "không trống"):
                st.text_input("Giá trị", value="(không cần)", disabled=True, key="new_val_disabled")
                chosen_val = None
            elif chosen_op in ("==", "!="):
                uniq = get_unique_values(df, chosen_col)
                if len(uniq) <= 100:
                    chosen_val = st.multiselect("Giá trị (chọn nhiều)", uniq, key="new_val_multi")
                else:
                    chosen_val = st.text_input("Giá trị", key="new_val_text")
            else:
                chosen_val = st.text_input("Giá trị", key="new_val_text2")

        if st.button("➕ Thêm điều kiện này", use_container_width=True):
            if chosen_op not in ("trống", "không trống") and not chosen_val and chosen_val != 0:
                st.warning("Vui lòng nhập giá trị.")
            else:
                st.session_state.conditions.append({
                    "col": chosen_col, "op": chosen_op, "val": chosen_val
                })
                st.rerun()

with col_right:
    logic = st.radio("Kết hợp điều kiện", ["VÀ (AND)", "HOẶC (OR)"], horizontal=False)

# Show current conditions
if st.session_state.conditions:
    st.markdown("**📋 Điều kiện hiện tại:**")
    for i, cond in enumerate(st.session_state.conditions):
        rc1, rc2 = st.columns([5, 1])
        with rc1:
            val_display = ", ".join(str(v) for v in cond["val"]) if isinstance(cond["val"], list) else str(cond["val"])
            if cond["op"] in ("trống", "không trống"):
                val_display = ""
            if i > 0:
                lbl = "VÀ" if logic == "VÀ (AND)" else "HOẶC"
                st.markdown(f'<div class="cond-card">🔗 <b>{lbl}</b> → [{cond["col"]}] <b>{cond["op"]}</b> <code>{val_display}</code></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="cond-card">🔹 [{cond["col"]}] <b>{cond["op"]}</b> <code>{val_display}</code></div>', unsafe_allow_html=True)
        with rc2:
            if st.button("🗑️", key=f"del_cond_{i}", help="Xóa điều kiện này"):
                st.session_state.conditions.pop(i)
                st.rerun()
else:
    st.info("Chưa có điều kiện nào. Hãy thêm điều kiện ở trên.")

# ── Preview matched rows ──────────────────────────────────────────────────
if st.session_state.conditions:
    mask = build_mask(df, st.session_state.conditions, logic)
    matched = df[mask]
    not_matched = df[~mask]

    st.divider()
    st.subheader("👁️ Xem trước kết quả")

    tab1, tab2 = st.tabs([
        f"✅ Thỏa điều kiện ({len(matched):,} hàng)",
        f"❌ Không thỏa ({len(not_matched):,} hàng)"
    ])
    with tab1:
        if len(matched):
            st.dataframe(matched.head(200), use_container_width=True, height=280)
            if len(matched) > 200:
                st.caption(f"Hiển thị 200/{len(matched)} hàng đầu tiên.")
        else:
            st.warning("Không có hàng nào thỏa điều kiện.")
    with tab2:
        if len(not_matched):
            st.dataframe(not_matched.head(200), use_container_width=True, height=280)
        else:
            st.warning("Tất cả các hàng đều thỏa điều kiện.")

    # ── Action buttons ────────────────────────────────────────────────────
    st.divider()
    st.subheader("⚡ Thực hiện thao tác")

    a1, a2, a3 = st.columns(3)
    with a1:
        if st.button(f"🗑️ XÓA {len(matched):,} hàng thỏa điều kiện", use_container_width=True, type="primary"):
            if len(matched) == 0:
                st.warning("Không có hàng nào để xóa.")
            else:
                cond_summary = " | ".join(f"[{c['col']}]{c['op']}" for c in st.session_state.conditions)
                st.session_state.history.append(f"Xóa {len(matched)} hàng: {cond_summary}")
                st.session_state.df_current = not_matched.reset_index(drop=True)
                st.session_state.conditions = []
                st.success(f"✅ Đã xóa {len(matched):,} hàng. Còn lại: {len(not_matched):,} hàng.")
                st.rerun()

    with a2:
        if st.button(f"✅ GIỮ {len(matched):,} hàng thỏa điều kiện", use_container_width=True):
            if len(matched) == 0:
                st.warning("Không có hàng nào để giữ.")
            else:
                cond_summary = " | ".join(f"[{c['col']}]{c['op']}" for c in st.session_state.conditions)
                st.session_state.history.append(f"Giữ {len(matched)} hàng: {cond_summary}")
                st.session_state.df_current = matched.reset_index(drop=True)
                st.session_state.conditions = []
                st.success(f"✅ Đã giữ lại {len(matched):,} hàng.")
                st.rerun()

    with a3:
        if st.button("🧹 Xóa tất cả điều kiện", use_container_width=True):
            reset_conditions()
            st.rerun()

# ── Full data view ────────────────────────────────────────────────────────
st.divider()
with st.expander("📄 Xem toàn bộ dữ liệu hiện tại", expanded=False):
    search_term = st.text_input("🔎 Tìm kiếm nhanh (lọc hiển thị)", placeholder="Nhập từ khóa...")
    if search_term:
        mask_search = df.apply(lambda row: row.astype(str).str.contains(search_term, case=False, na=False).any(), axis=1)
        display_df = df[mask_search]
        st.caption(f"Tìm thấy {len(display_df):,} hàng chứa '{search_term}'")
    else:
        display_df = df

    st.dataframe(display_df, use_container_width=True, height=400)

# ── Column manager ─────────────────────────────────────────────────────────
st.divider()
with st.expander("🗂️ Quản lý cột", expanded=False):
    cm1, cm2 = st.columns(2)
    with cm1:
        st.markdown("**Xóa cột không cần thiết**")
        cols_to_drop = st.multiselect("Chọn cột muốn xóa", df.columns.tolist())
        if st.button("🗑️ Xóa cột đã chọn") and cols_to_drop:
            st.session_state.df_current = df.drop(columns=cols_to_drop)
            st.session_state.history.append(f"Xóa cột: {', '.join(cols_to_drop)}")
            st.rerun()
    with cm2:
        st.markdown("**Đổi tên cột**")
        col_rename = st.selectbox("Chọn cột cần đổi tên", df.columns.tolist(), key="rename_col")
        new_name = st.text_input("Tên mới", key="rename_val")
        if st.button("✏️ Đổi tên") and new_name:
            st.session_state.df_current = df.rename(columns={col_rename: new_name})
            st.session_state.history.append(f"Đổi tên cột '{col_rename}' → '{new_name}'")
            st.rerun()

# ── Export ────────────────────────────────────────────────────────────────
st.divider()
st.subheader("💾 Xuất file")

ex1, ex2, ex3 = st.columns(3)
with ex1:
    excel_bytes = to_excel_bytes(df)
    st.download_button(
        label=f"📥 Tải xuống Excel ({len(df):,} hàng)",
        data=excel_bytes,
        file_name="ket_qua_loc.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

with ex2:
    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label=f"📥 Tải xuống CSV ({len(df):,} hàng)",
        data=csv_bytes,
        file_name="ket_qua_loc.csv",
        mime="text/csv",
        use_container_width=True,
    )

with ex3:
    if st.session_state.conditions:
        mask_export = build_mask(df, st.session_state.conditions, logic)
        removed_df = df[mask_export]
        removed_bytes = to_excel_bytes(removed_df)
        st.download_button(
            label=f"📥 Xuất hàng thỏa ĐK ({len(removed_df):,} hàng)",
            data=removed_bytes,
            file_name="hang_thoa_dieu_kien.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.button("📥 Xuất hàng thỏa ĐK", disabled=True, use_container_width=True, help="Thêm điều kiện trước")

st.caption("Made with ❤️ using Streamlit · Hỗ trợ Excel & CSV · Lọc đa điều kiện")
