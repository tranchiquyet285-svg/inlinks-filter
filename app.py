import io
import re
from collections import Counter

import pandas as pd
import streamlit as st

from filter_logic import (
    BUTTON_ANCHORS,
    CONTEXTUAL_PATTERNS,
    EXCLUDE_PATTERNS,
    GROUP_COLORS,
    apply_filter,
    build_excel,
    normalise,
)

# ── Cấu hình trang ─────────────────────────────────────────────
st.set_page_config(
    page_title="Inlinks Filter",
    page_icon="🔗",
    layout="wide",
)

# ── CSS nhỏ để đẹp hơn ─────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background: #F0F7FF;
        border-radius: 10px;
        padding: 16px 20px;
        text-align: center;
    }
    .group-badge {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 13px;
        font-weight: 600;
    }
    div[data-testid="stDownloadButton"] button {
        background-color: #1D4ED8;
        color: white;
        font-weight: bold;
        padding: 10px 28px;
        border-radius: 8px;
        border: none;
        font-size: 15px;
    }
</style>
""", unsafe_allow_html=True)


# ── Header ──────────────────────────────────────────────────────
st.title("🔗 Contextual Inlinks Filter")
st.caption("Upload file **inlinks.csv** từ Screaming Frog → lọc link contextual → tải về Excel")

st.divider()


# ── Upload ──────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Chọn file inlinks.csv",
    type=["csv"],
    help="Export từ Screaming Frog: tab All Inlinks → Export",
)

if not uploaded:
    st.info("Chưa có file. Upload file CSV để bắt đầu.")
    st.stop()


# ── Đọc CSV ─────────────────────────────────────────────────────
@st.cache_data(show_spinner="Đang đọc file...")
def read_csv(file_bytes: bytes) -> pd.DataFrame:
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1258"):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=enc, low_memory=False)
        except UnicodeDecodeError:
            continue
    st.error("Không thể đọc file — thử convert sang UTF-8 rồi upload lại.")
    st.stop()


raw_bytes = uploaded.read()
df_raw = read_csv(raw_bytes)

df, err = normalise(df_raw)
if err:
    st.error(f"**Lỗi cột:** {err}")
    st.stop()


# ── Lọc ─────────────────────────────────────────────────────────
with st.spinner("Đang lọc..."):
    kept, stats = apply_filter(df)


# ── Metrics ─────────────────────────────────────────────────────
pct = stats["kept"] / stats["total"] * 100 if stats["total"] else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("Tổng link input",        f"{stats['total']:,}")
col2.metric("Link contextual giữ lại", f"{stats['kept']:,}")
col3.metric("Tỷ lệ giữ lại",          f"{pct:.1f}%")
col4.metric("Số trang nguồn (From)",  f"{kept['From'].nunique():,}" if len(kept) > 0 else "0")

st.divider()


# ── 2 cột: Lý do loại | Phân bố nhóm ───────────────────────────
left, right = st.columns(2)

with left:
    st.subheader("Lý do loại bỏ")
    reject_data = {
        "Lý do":      ["Type ≠ Hyperlink", "Position ≠ Content", "Anchor rỗng",
                        "Anchor là button/số", "Path marker loại trừ", "Không khớp contextual"],
        "Số lượng":   [stats["non_hyperlink"], stats["non_content"], stats["empty_anchor"],
                        stats["button_anchor"], stats["excluded_path"], stats["no_match"]],
    }
    st.dataframe(
        pd.DataFrame(reject_data),
        hide_index=True,
        use_container_width=True,
    )

with right:
    st.subheader("Phân bố nhóm giữ lại")
    if stats["groups"]:
        group_data = pd.DataFrame(
            sorted(stats["groups"].items(), key=lambda x: -x[1]),
            columns=["Nhóm", "Số lượng"],
        )
        st.dataframe(group_data, hide_index=True, use_container_width=True)
        st.bar_chart(group_data.set_index("Nhóm"))
    else:
        st.warning("Không có link nào được giữ lại.")

st.divider()


# ── Survey: top id/class trong file ─────────────────────────────
if stats["no_match"] > 0:
    with st.expander(f"⚠️ {stats['no_match']:,} rows không khớp pattern — Xem id/class trong file để điều chỉnh"):
        st.caption("Đây là các id và class xuất hiện trong cột Link Path. Dùng để xác định pattern contextual cho website này.")

        if "Link Path" in df.columns:
            all_paths = " ".join(df["Link Path"].dropna())
            ids     = re.findall(r"@id=['\"]([^'\"]+)['\"]",    all_paths)
            classes = re.findall(r"@class=['\"]([^'\"]+)['\"]", all_paths)
            # fallback không có @
            ids     += re.findall(r"(?<![a-z])id=['\"]([^'\"]+)['\"]",    all_paths)
            classes += re.findall(r"(?<![a-z])class=['\"]([^'\"]+)['\"]", all_paths)

            id_df  = pd.DataFrame(Counter(ids).most_common(40),  columns=["id",    "Số lần xuất hiện"])
            cls_df = pd.DataFrame(Counter(classes).most_common(40), columns=["class", "Số lần xuất hiện"])

            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Top 40 `@id`**")
                st.dataframe(id_df,  hide_index=True, use_container_width=True, height=400)
            with c2:
                st.markdown("**Top 40 `@class`**")
                st.dataframe(cls_df, hide_index=True, use_container_width=True, height=400)

            st.info(
                "👉 Tìm id/class tương ứng với **tab mô tả, body bài viết, mô tả category** "
                "→ báo lại để cập nhật pattern cho website của bạn."
            )

st.divider()


# ── Preview kết quả ─────────────────────────────────────────────
if len(kept) > 0:
    st.subheader(f"Xem trước — {min(50, len(kept))} / {len(kept):,} rows")

    preview_cols = [c for c in ["Group", "From", "To", "Anchor Text", "Link Path"] if c in kept.columns]
    st.dataframe(
        kept[preview_cols].head(50),
        hide_index=True,
        use_container_width=True,
        column_config={
            "Group":       st.column_config.TextColumn("Nhóm", width=180),
            "From":        st.column_config.LinkColumn("From", width=220),
            "To":          st.column_config.LinkColumn("To", width=220),
            "Anchor Text": st.column_config.TextColumn("Anchor", width=160),
            "Link Path":   st.column_config.TextColumn("Link Path", width=300),
        },
    )

    st.divider()


# ── Download ─────────────────────────────────────────────────────
st.subheader("Tải về kết quả")

if len(kept) == 0:
    st.warning("Không có link contextual nào — kiểm tra lại file hoặc quy tắc lọc.")
else:
    with st.spinner("Đang tạo file Excel..."):
        excel_bytes = build_excel(kept, stats)

    filename = uploaded.name.replace(".csv", "_contextual.xlsx")
    st.download_button(
        label="📥 Tải về Excel (4 sheet)",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("File Excel gồm: **Contextual Links** · **Theo From** · **Theo To (SEO)** · **Quy tắc & Thống kê**")


# ── Expander: xem quy tắc lọc ───────────────────────────────────
with st.expander("🔍 Xem quy tắc lọc đang áp dụng"):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**✅ Contextual patterns (GIỮ — ít nhất 1 phải khớp)**")
        for p in CONTEXTUAL_PATTERNS:
            st.code(p, language=None)
    with c2:
        st.markdown("**❌ Exclude patterns (LOẠI — không được khớp bất kỳ)**")
        for p, label in EXCLUDE_PATTERNS.items():
            st.code(f"{p}  →  {label}", language=None)

    st.markdown("**🚫 Button anchors bị loại**")
    st.write(", ".join(f"`{a}`" for a in sorted(BUTTON_ANCHORS)))
    st.caption("Anchor chỉ là số nguyên (1, 2, 3...) cũng bị loại tự động.")
