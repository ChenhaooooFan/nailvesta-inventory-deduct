import re
from datetime import date

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="NailVesta 每日水单库存扣减工具",
    page_icon="💅",
    layout="wide",
)

st.title("💅 NailVesta 每日水单库存扣减工具")
st.caption("上传库存表 + 各水单 CSV，选择日期后，程序按【款式名称 + 尺码】汇总扣减库存。")


# =========================
# 基础工具函数
# =========================

def read_csv_safely(uploaded_file):
    """读取 Lark / Excel 导出的 CSV；所有字段按文本读，避免订单号和 tracking 被科学计数法。"""
    if uploaded_file is None:
        return None

    for enc in ["utf-8-sig", "utf-8", "gb18030"]:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc, dtype=str).fillna("")
        except Exception:
            continue

    uploaded_file.seek(0)
    return pd.read_csv(uploaded_file, dtype=str).fillna("")


def clean_col(col):
    return str(col).replace("\ufeff", "").strip()


def clean_dataframe(df):
    if df is None:
        return None
    out = df.copy()
    out.columns = [clean_col(c) for c in out.columns]
    return out.fillna("")


def norm_text(value):
    text = str(value).replace("\ufeff", "").replace("\u200b", "").replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def norm_style(value):
    return norm_text(value).lower()


def norm_size(value):
    text = norm_text(value).upper()
    # 只抓明确的尺码，避免误识别普通单词里的 S/M/L。
    m = re.search(r"\b(XS|XL|S|M|L)\b", text)
    return m.group(1) if m else text


def size_from_sku(sku):
    text = norm_text(sku).upper()
    m = re.search(r"(?:-|_|\s)(XS|XL|S|M|L)$", text)
    if m:
        return m.group(1)
    m = re.search(r"(XS|XL|S|M|L)$", text)
    return m.group(1) if m else ""


def find_col(df, candidates, contains_any=None):
    """先精确匹配列名，再模糊包含匹配。"""
    if df is None or df.empty and len(df.columns) == 0:
        return None

    cols = list(df.columns)
    lower_map = {str(c).strip().lower(): c for c in cols}

    for cand in candidates:
        key = str(cand).strip().lower()
        if key in lower_map:
            return lower_map[key]

    for cand in candidates:
        key = str(cand).strip().lower()
        if not key:
            continue
        for col in cols:
            if key in str(col).strip().lower():
                return col

    if contains_any:
        keys = [str(x).lower() for x in contains_any]
        for col in cols:
            lc = str(col).lower()
            if any(k in lc for k in keys):
                return col
    return None


def first_existing_col(df, candidates):
    for c in candidates:
        col = find_col(df, [c])
        if col:
            return col
    return None


def parse_date_series(series):
    return pd.to_datetime(series.astype(str).str.strip(), errors="coerce").dt.date


def date_col_name(selected_date):
    # 和你库存表里的 05/05 格式保持一致。
    return selected_date.strftime("%m/%d")


def to_csv_bytes(df):
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# =========================
# 款式拆分：重点修正版
# =========================

def split_style_names(raw_value):
    """
    支持一格多个款式，例如：
    - Teal Blossom, Tidal Flower
    - Starlit Rift, Ruby Bloom
    - Cherry Romance ｜ 库位：B-04-11,Aqua Blush ｜ 库位：A-02-07

    一个 cell 里有几个款式，就每个款式各扣 1 件，尺码使用该行的尺码。
    """
    text = norm_text(raw_value)
    if not text:
        return []

    # 如果是“款式 + 库位”字段，优先抓每段“｜ 库位”前面的款式名。
    if "库位" in text and ("｜" in text or "|" in text):
        pieces = re.split(r"[,，\n;；]+", text)
        styles = []
        for piece in pieces:
            piece = norm_text(piece)
            if not piece:
                continue
            piece = re.split(r"\s*[|｜]\s*库位", piece, maxsplit=1)[0]
            piece = norm_text(piece)
            if piece:
                styles.append(piece)
        return styles

    # 普通 Lark 多选字段 CSV 导出一般是英文逗号分隔。
    pieces = re.split(r"[,，\n;；]+", text)
    cleaned = []
    for piece in pieces:
        piece = norm_text(piece)
        if not piece:
            continue
        # 兜底：如果普通列里混进了“款式 ｜ 库位：xxx”，去掉库位部分。
        piece = re.split(r"\s*[|｜]\s*库位", piece, maxsplit=1)[0]
        piece = norm_text(piece)
        if piece:
            cleaned.append(piece)
    return cleaned


def truthy_series(series):
    text = series.astype(str).str.strip().str.lower()
    return (
        text.isin(["1", "true", "yes", "y", "done", "packed", "shipped", "completed"])
        | text.str.contains("已打包|已完成|完成|发货|packed|shipped|completed", case=False, na=False)
    )


def optional_done_filter(df, done_only):
    if not done_only or df is None or df.empty:
        return df

    possible_cols = [
        "是否完成发货",
        "是否打包",
        "状态_ (status)",
        "状态 (status)",
        "status",
        "Status",
    ]
    status_cols = [c for c in possible_cols if c in df.columns]
    if not status_cols:
        return df

    mask = pd.Series(False, index=df.index)
    for col in status_cols:
        mask = mask | truthy_series(df[col])
    return df[mask].copy()


# =========================
# 库存表处理
# =========================

def prepare_inventory(inv_raw):
    inv_original = clean_dataframe(inv_raw)
    inv = inv_original.copy()

    style_col = find_col(inv, ["产品名称", "款式名称", "Product Name", "Style Name", "款式"])
    sku_col = find_col(inv, ["SKU编码", "SKU", "Full SKU", "Variant SKU", "sku"])
    size_col = find_col(inv, ["尺码", "尺码 (size)", "Size", "Size'", "Variant Size"])
    stock_col = find_col(
        inv,
        ["当前库存", "库存量", "库存数量", "剩余库存", "实际库存", "可用库存", "On hand", "Quantity", "Qty", "Stock", "Inventory"],
        contains_any=["当前库存", "库存", "quantity", "qty", "stock", "inventory", "on hand"],
    )

    if style_col:
        # 产品名称在 S/M/L 三行里通常只有第一行有值，自动向下填充仅用于匹配。
        inv["__match_style"] = inv[style_col].replace("", pd.NA).ffill().fillna("").map(norm_text)
    else:
        inv["__match_style"] = ""

    if size_col:
        inv["__match_size"] = inv[size_col].map(norm_size)
    elif sku_col:
        inv["__match_size"] = inv[sku_col].map(size_from_sku)
    else:
        inv["__match_size"] = ""

    inv["__norm_style"] = inv["__match_style"].map(norm_style)
    inv["__norm_size"] = inv["__match_size"].map(norm_size)

    return inv_original, inv, style_col, sku_col, size_col, stock_col


# =========================
# 来源表提取
# =========================

def extract_rows_from_one_table(df_raw, table_name, selected_date, style_cols_priority, size_cols_priority, done_only):
    df = clean_dataframe(df_raw)
    if df is None or (df.empty and len(df.columns) == 0):
        return pd.DataFrame(), [f"{table_name}：空表，已跳过。"]

    date_col = find_col(df, ["日期", "Date", "date"])
    if not date_col:
        return pd.DataFrame(), [f"{table_name}：没有找到日期列，已跳过。"]

    tmp = df.copy()
    tmp["__parsed_date"] = parse_date_series(tmp[date_col])
    tmp = tmp[tmp["__parsed_date"] == selected_date].copy()

    if tmp.empty:
        return pd.DataFrame(), [f"{table_name}：所选日期没有记录。"]

    tmp = optional_done_filter(tmp, done_only)
    if tmp.empty:
        return pd.DataFrame(), [f"{table_name}：所选日期有记录，但没有符合已打包 / 已完成发货条件的记录。"]

    style_col = first_existing_col(tmp, style_cols_priority)
    size_col = first_existing_col(tmp, size_cols_priority)

    if not style_col or not size_col:
        return pd.DataFrame(), [
            f"{table_name}：没有识别到款式列或尺码列，已跳过。当前列名：{', '.join(tmp.columns)}"
        ]

    rows = []
    warnings = []
    for idx, row in tmp.iterrows():
        raw_style = row.get(style_col, "")
        raw_size = row.get(size_col, "")
        styles = split_style_names(raw_style)
        size = norm_size(raw_size)

        if not styles or not size:
            warnings.append(f"{table_name}：第 {int(idx) + 2} 行款式或尺码为空，已跳过。")
            continue

        for style in styles:
            rows.append({
                "款式名称": style,
                "尺码": size,
                "扣减数量": 1,
            })

    return pd.DataFrame(rows), warnings


def build_deduction_summary(uploaded_sources, selected_date, done_only):
    all_rows = []
    all_warnings = []

    for source in uploaded_sources:
        table_name = source["name"]
        uploaded = source["file"]
        if uploaded is None:
            continue

        df_raw = read_csv_safely(uploaded)
        rows, warnings = extract_rows_from_one_table(
            df_raw=df_raw,
            table_name=table_name,
            selected_date=selected_date,
            style_cols_priority=source["style_cols"],
            size_cols_priority=source["size_cols"],
            done_only=done_only,
        )
        if not rows.empty:
            all_rows.append(rows)
        all_warnings.extend(warnings)

    if not all_rows:
        return pd.DataFrame(columns=["款式名称", "尺码", "本次扣减数量"]), all_warnings

    detail = pd.concat(all_rows, ignore_index=True)
    detail["__norm_style"] = detail["款式名称"].map(norm_style)
    detail["__norm_size"] = detail["尺码"].map(norm_size)

    summary = (
        detail.groupby(["__norm_style", "__norm_size"], as_index=False)
        .agg(
            款式名称=("款式名称", "first"),
            尺码=("尺码", "first"),
            本次扣减数量=("扣减数量", "sum"),
        )
        .sort_values(["款式名称", "尺码"])
        .reset_index(drop=True)
    )
    return summary, all_warnings


# =========================
# 页面：上传和设置
# =========================

with st.sidebar:
    st.header("1）上传 CSV")
    inventory_file = st.file_uploader("库存表 CSV（必传）", type=["csv"])

    st.divider()
    b4g1_file = st.file_uploader("水单补寄表 / B4G1_B4 表 CSV", type=["csv"])
    normal_file = st.file_uploader("水单表 - 新普通水单 CSV", type=["csv"])
    influencer_file = st.file_uploader("深达水单表 CSV", type=["csv"])
    exchange_file = st.file_uploader("达人换货表 CSV", type=["csv"])

    st.header("2）选择日期")
    selected_date = st.date_input("只扣减这个日期的记录", value=date.today())
    done_only = st.checkbox("只扣已打包 / 已完成发货记录", value=False)
    update_date_snapshot = st.checkbox("同时新增 / 更新所选日期库存列", value=True)
    floor_zero = st.checkbox("扣减后库存不低于 0", value=False)


if inventory_file is None:
    st.warning("请先在左侧上传库存表 CSV。")
    st.stop()

inv_raw = read_csv_safely(inventory_file)
inv_original, inv_match, inv_style_col, sku_col, inv_size_col, stock_col = prepare_inventory(inv_raw)

if not inv_style_col:
    st.error("库存表没有识别到【产品名称 / 款式名称】列。")
    st.write("当前库存表列名：", list(inv_original.columns))
    st.stop()

if not sku_col and not inv_size_col:
    st.error("库存表没有识别到【SKU编码】或【尺码】列，无法判断 S / M / L。")
    st.write("当前库存表列名：", list(inv_original.columns))
    st.stop()

if not stock_col:
    st.error("库存表没有识别到【当前库存】列。请确认库存列名是否为：当前库存。")
    st.write("当前库存表列名：", list(inv_original.columns))
    st.stop()

# 重复 key 会导致扣错，必须阻止。
valid_key = inv_match["__norm_style"].ne("") & inv_match["__norm_size"].ne("")
dup_mask = inv_match.duplicated(["__norm_style", "__norm_size"], keep=False) & valid_key
if dup_mask.any():
    st.error("库存表存在重复的【款式名称 + 尺码】，为避免扣错，程序已停止。")
    show_cols = [c for c in [inv_style_col, sku_col, inv_size_col, stock_col] if c and c in inv_match.columns]
    st.dataframe(inv_match.loc[dup_mask, show_cols], use_container_width=True)
    st.stop()

sources = [
    {
        "name": "赠送款式表/B4G1",
        "file": b4g1_file,
        "style_cols": ["赠送款式 Style Names", "Product Name", "款式", "款式名称", "Style Names", "款式 + 库位"],
        "size_cols": ["尺码 (size)", "Size'", "Size", "尺码"],
    },
    {
        "name": "新普通水单",
        "file": normal_file,
        "style_cols": ["Product Name", "款式", "款式名称", "Style Names", "赠送款式 Style Names", "款式 + 库位"],
        "size_cols": ["Size'", "Size", "尺码", "尺码 (size)"],
    },
    {
        "name": "深达水单表",
        "file": influencer_file,
        "style_cols": ["Product Name", "Product Name1", "款式", "款式名称", "Style Names", "款式 + 库位"],
        "size_cols": ["Size'", "Size", "尺码", "尺码 (size)"],
    },
    {
        "name": "达人换货表",
        "file": exchange_file,
        "style_cols": ["发货款式", "Product Name", "款式", "款式名称"],
        "size_cols": ["发货尺码", "Size'", "Size", "尺码", "尺码 (size)"],
    },
]

summary, warnings = build_deduction_summary(sources, selected_date, done_only)

if summary.empty:
    st.warning("所选日期没有可扣减记录。")
    if warnings:
        with st.expander("查看提示"):
            for w in warnings:
                st.write("-", w)
    st.stop()

# =========================
# 扣减库存
# =========================

result = inv_original.copy()
work = inv_match.copy()

work["__stock_before"] = pd.to_numeric(
    work[stock_col].astype(str).str.replace(",", "", regex=False).str.strip(),
    errors="coerce",
).fillna(0)

summary_for_merge = summary[["__norm_style", "__norm_size", "本次扣减数量"]].copy()
work = work.merge(summary_for_merge, on=["__norm_style", "__norm_size"], how="left")
work["本次扣减数量"] = work["本次扣减数量"].fillna(0).astype(int)
work["__stock_after"] = work["__stock_before"] - work["本次扣减数量"]
negative_mask = work["__stock_after"] < 0

if floor_zero:
    work["__stock_after"] = work["__stock_after"].clip(lower=0)

# 只把最终库存写回原始表，不添加任何诊断列。
result[stock_col] = work["__stock_after"].round(0).astype(int)

if update_date_snapshot:
    day_col = date_col_name(selected_date)
    result[day_col] = work["__stock_after"].round(0).astype(int)

matched_keys = set(zip(work.loc[work["本次扣减数量"] > 0, "__norm_style"], work.loc[work["本次扣减数量"] > 0, "__norm_size"]))
summary["__matched"] = summary.apply(lambda r: (r["__norm_style"], r["__norm_size"]) in matched_keys, axis=1)
matched_summary = summary[summary["__matched"]].copy()
unmatched_summary = summary[~summary["__matched"]].copy()

show_summary = matched_summary[["款式名称", "尺码", "本次扣减数量"]].sort_values(["款式名称", "尺码"]).reset_index(drop=True)
all_summary_clean = summary[["款式名称", "尺码", "本次扣减数量"]].sort_values(["款式名称", "尺码"]).reset_index(drop=True)

# =========================
# 页面展示：只展示用户需要的两部分
# =========================

st.subheader("1）本次扣减汇总")
st.caption("一个 cell 里有多个款式时，程序会拆成多个款式分别扣减；同一行的尺码会应用到该行所有款式。")

m1, m2, m3 = st.columns(3)
with m1:
    st.metric("本次扣减总件数", int(show_summary["本次扣减数量"].sum()) if not show_summary.empty else 0)
with m2:
    st.metric("扣减款式 + 尺码数", len(show_summary))
with m3:
    st.metric("未匹配项", len(unmatched_summary))

st.dataframe(show_summary, use_container_width=True, hide_index=True)

if not unmatched_summary.empty:
    st.error("以下款式 + 尺码没有在库存表中匹配到，所以没有扣减。请检查水单款式名 / 尺码是否和库存表一致。")
    st.dataframe(
        unmatched_summary[["款式名称", "尺码", "本次扣减数量"]].sort_values(["款式名称", "尺码"]),
        use_container_width=True,
        hide_index=True,
    )

if negative_mask.any():
    negative_rows = work.loc[negative_mask, ["__match_style", "__match_size", "__stock_before", "本次扣减数量", "__stock_after"]].copy()
    negative_rows.columns = ["款式名称", "尺码", "扣减前库存", "本次扣减数量", "扣减后库存"]
    st.warning("有库存扣减后小于 0，请重点核对。")
    st.dataframe(negative_rows, use_container_width=True, hide_index=True)

if warnings:
    with st.expander("查看跳过 / 识别提示"):
        for w in warnings:
            st.write("-", w)

st.subheader("2）扣完库存后的最新全部库存")
st.caption("这里保留库存表原本的格式；只更新【当前库存】，并按需新增 / 更新所选日期库存列。")
st.dataframe(result, use_container_width=True, hide_index=True)

st.subheader("下载结果")
col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "下载：扣完库存后的最新库存 CSV",
        data=to_csv_bytes(result),
        file_name=f"NailVesta_库存扣减后_{selected_date.strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )
with col2:
    st.download_button(
        "下载：本次扣减汇总 CSV",
        data=to_csv_bytes(all_summary_clean),
        file_name=f"NailVesta_本次扣减汇总_{selected_date.strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )
