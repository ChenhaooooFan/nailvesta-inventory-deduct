import hashlib
import re
from datetime import date, datetime

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="NailVesta 水单库存扣减工具",
    page_icon="💅",
    layout="wide",
)

st.title("💅 NailVesta 水单库存扣减工具")
st.caption("上传库存表 + 各水单总表 CSV，手动选择扣减日期区间后，程序按【款式名称 + 尺码】汇总扣减库存；支持上传扣减日志，避免同一天旧记录重复扣减。")


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
    if df is None or (df.empty and len(df.columns) == 0):
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
    # 兼容 Lark/Excel 导出的 2026/05/06、2026-05-06、带时间戳等格式。
    return pd.to_datetime(series.astype(str).str.strip(), errors="coerce").dt.date


def format_date_distribution(parsed_dates):
    valid = parsed_dates.dropna()
    if valid.empty:
        return "没有可识别日期"
    counts = valid.value_counts().sort_index()
    return "；".join([f"{d.strftime('%Y/%m/%d')}：{int(n)}条" for d, n in counts.items()])


def get_date_distribution_row(df_raw, table_name):
    df = clean_dataframe(df_raw)
    if df is None or (df.empty and len(df.columns) == 0):
        return {"来源表": table_name, "总行数": 0, "日期列": "未找到", "可识别日期分布": "空表"}

    date_col = find_col(df, ["日期", "Date", "date"])
    if not date_col:
        return {"来源表": table_name, "总行数": len(df), "日期列": "未找到", "可识别日期分布": "无法检查"}

    parsed = parse_date_series(df[date_col]) if len(df) else pd.Series([], dtype=object)
    return {
        "来源表": table_name,
        "总行数": len(df),
        "日期列": date_col,
        "可识别日期分布": format_date_distribution(parsed),
    }


def get_available_dates_from_one_file(uploaded_file):
    if uploaded_file is None:
        return []
    df_raw = read_csv_safely(uploaded_file)
    df = clean_dataframe(df_raw)
    if df is None or (df.empty and len(df.columns) == 0):
        return []
    date_col = find_col(df, ["日期", "Date", "date"])
    if not date_col:
        return []
    parsed = parse_date_series(df[date_col]).dropna()
    return sorted(set(parsed.tolist()))


def date_col_name(selected_date):
    # 和你库存表里的 05/05 格式保持一致。
    return selected_date.strftime("%m/%d")


def format_dates_label(selected_dates):
    if not selected_dates:
        return "未选择"
    return "、".join([d.strftime("%Y/%m/%d") for d in selected_dates])


def file_date_label(selected_dates):
    if not selected_dates:
        return datetime.now().strftime("%Y%m%d_%H%M")
    dates = sorted(selected_dates)
    if len(dates) == 1:
        return dates[0].strftime("%Y%m%d")
    return f"{dates[0].strftime('%Y%m%d')}-{dates[-1].strftime('%Y%m%d')}_{len(dates)}dates"


def to_csv_bytes(df):
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def stable_hash(text):
    return hashlib.sha1(str(text).encode("utf-8", errors="ignore")).hexdigest()[:20]


def find_existing_cols(df, candidates):
    """返回候选列中实际存在的列，尽量保留候选顺序。"""
    if df is None:
        return []
    found = []
    for cand in candidates:
        col = find_col(df, [cand])
        if col and col not in found:
            found.append(col)
    return found


def build_source_record_base(row, table_name, row_index, date_col, style_col, size_col, id_cols):
    """
    生成来源行的稳定身份。
    优先使用 Order ID / Tracking / Handle 等业务字段；如果这些字段都空，
    再用日期 + 款式 + 尺码 + 行号兜底。
    """
    parts = [f"source={table_name}"]

    date_value = norm_text(row.get(date_col, "")) if date_col else ""
    if date_value:
        parts.append(f"date={date_value}")

    has_business_id = False
    for col in id_cols:
        value = norm_text(row.get(col, ""))
        if value:
            parts.append(f"{col}={value}")
            has_business_id = True

    # 有业务 ID 时，行身份尽量不包含款式 cell。这样同一行之后新增款式时，
    # 旧款式的扣减 ID 不会因为 cell 内容变化而重新生成，避免重复扣。
    # 没有业务 ID 时，才用款式 + 尺码 + 行号兜底。
    if not has_business_id:
        style_value = norm_text(row.get(style_col, "")) if style_col else ""
        size_value = norm_text(row.get(size_col, "")) if size_col else ""
        if style_value:
            parts.append(f"style_cell={style_value}")
        if size_value:
            parts.append(f"size_cell={size_value}")
        parts.append(f"fallback_row={int(row_index) + 2}")

    return "|".join(parts)


def make_deduction_item_id(table_name, source_record_base, style, size):
    # 同一行如果有多个款式，每个款式 + 尺码生成一个扣减 ID；之后新增款式不会被旧记录挡掉。
    raw = f"{table_name}|{source_record_base}|style={norm_style(style)}|size={norm_size(size)}"
    return stable_hash(raw)


def read_previous_deduction_log(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame(columns=["扣减记录ID"])
    log = clean_dataframe(read_csv_safely(uploaded_file))
    if log is None or log.empty:
        return pd.DataFrame(columns=["扣减记录ID"])
    if "扣减记录ID" not in log.columns:
        st.warning("你上传了扣减日志，但里面没有【扣减记录ID】列；程序无法用它防重复，本次会当作没有历史日志。")
        return pd.DataFrame(columns=["扣减记录ID"])
    log["扣减记录ID"] = log["扣减记录ID"].astype(str).map(norm_text)
    log = log[log["扣减记录ID"].ne("")].copy()
    return log.drop_duplicates(subset=["扣减记录ID"], keep="last")


def summarize_detail(detail_df):
    empty_summary_cols = ["__norm_style", "__norm_size", "款式名称", "尺码", "本次扣减数量"]
    if detail_df is None or detail_df.empty:
        return pd.DataFrame(columns=empty_summary_cols)
    summary = (
        detail_df.groupby(["__norm_style", "__norm_size"], as_index=False)
        .agg(
            款式名称=("款式名称", "first"),
            尺码=("尺码", "first"),
            本次扣减数量=("扣减数量", "sum"),
        )
        .sort_values(["款式名称", "尺码"])
        .reset_index(drop=True)
    )
    summary["本次扣减数量"] = summary["本次扣减数量"].astype(int)
    return summary


def build_updated_deduction_log(previous_log, new_detail, selected_dates):
    now_label = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    batch_label = file_date_label(selected_dates)

    new_log = pd.DataFrame(columns=[
        "扣减记录ID", "扣减时间", "扣减批次", "日期", "来源表", "来源记录ID", "款式名称", "尺码", "扣减数量"
    ])
    if new_detail is not None and not new_detail.empty:
        new_log = new_detail[["扣减记录ID", "日期", "来源表", "来源记录ID", "款式名称", "尺码", "扣减数量"]].copy()
        new_log.insert(1, "扣减时间", now_label)
        new_log.insert(2, "扣减批次", batch_label)
        new_log["日期"] = new_log["日期"].apply(lambda x: x.strftime("%Y/%m/%d") if hasattr(x, "strftime") else str(x))

    if previous_log is None or previous_log.empty:
        combined = new_log.copy()
    else:
        combined = pd.concat([previous_log, new_log], ignore_index=True, sort=False)

    if not combined.empty and "扣减记录ID" in combined.columns:
        combined["扣减记录ID"] = combined["扣减记录ID"].astype(str).map(norm_text)
        combined = combined[combined["扣减记录ID"].ne("")].drop_duplicates(subset=["扣减记录ID"], keep="last")

    return combined


# =========================
# 款式拆分
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

def extract_rows_from_one_table(df_raw, table_name, selected_dates, style_cols_priority, size_cols_priority, id_cols_priority, done_only):
    df = clean_dataframe(df_raw)
    if df is None or (df.empty and len(df.columns) == 0):
        return pd.DataFrame(), [f"{table_name}：空表，已跳过。"]

    if df.empty:
        return pd.DataFrame(), [f"{table_name}：CSV 只有表头，0 条记录，已跳过。"]

    date_col = find_col(df, ["日期", "Date", "date"])
    if not date_col:
        return pd.DataFrame(), [f"{table_name}：没有找到日期列，已跳过。当前列名：{', '.join(df.columns)}"]

    selected_dates = sorted(set(selected_dates))
    selected_dates_set = set(selected_dates)
    selected_label = format_dates_label(selected_dates)

    tmp = df.copy()
    tmp["__parsed_date"] = parse_date_series(tmp[date_col])
    date_distribution = format_date_distribution(tmp["__parsed_date"])
    tmp = tmp[tmp["__parsed_date"].isin(selected_dates_set)].copy()

    if tmp.empty:
        return pd.DataFrame(), [
            f"{table_name}：所选日期 {selected_label} 没有记录。这个 CSV 里可识别日期是：{date_distribution}。"
        ]

    tmp = optional_done_filter(tmp, done_only)
    if tmp.empty:
        return pd.DataFrame(), [f"{table_name}：所选日期有记录，但没有符合已打包 / 已完成发货条件的记录。"]

    style_col = first_existing_col(tmp, style_cols_priority)
    size_col = first_existing_col(tmp, size_cols_priority)
    id_cols = find_existing_cols(tmp, id_cols_priority)

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
        row_date = row.get("__parsed_date")

        if not styles or not size:
            warnings.append(f"{table_name}：第 {int(idx) + 2} 行款式或尺码为空，已跳过。")
            continue

        source_record_base = build_source_record_base(
            row=row,
            table_name=table_name,
            row_index=idx,
            date_col=date_col,
            style_col=style_col,
            size_col=size_col,
            id_cols=id_cols,
        )
        source_record_id = stable_hash(source_record_base)

        for style in styles:
            deduction_item_id = make_deduction_item_id(table_name, source_record_base, style, size)
            rows.append({
                "扣减记录ID": deduction_item_id,
                "来源记录ID": source_record_id,
                "日期": row_date,
                "款式名称": style,
                "尺码": size,
                "扣减数量": 1,
                "来源表": table_name,
            })

    return pd.DataFrame(rows), warnings


def build_deduction_summary(uploaded_sources, selected_dates, done_only):
    all_rows = []
    all_warnings = []
    date_check_rows = []

    for source in uploaded_sources:
        table_name = source["name"]
        uploaded = source["file"]
        if uploaded is None:
            continue

        df_raw = read_csv_safely(uploaded)
        date_check_rows.append(get_date_distribution_row(df_raw, table_name))
        rows, warnings = extract_rows_from_one_table(
            df_raw=df_raw,
            table_name=table_name,
            selected_dates=selected_dates,
            style_cols_priority=source["style_cols"],
            size_cols_priority=source["size_cols"],
            id_cols_priority=source.get("id_cols", []),
            done_only=done_only,
        )
        if not rows.empty:
            all_rows.append(rows)
        all_warnings.extend(warnings)

    date_check = pd.DataFrame(date_check_rows, columns=["来源表", "总行数", "日期列", "可识别日期分布"])

    empty_summary_cols = ["__norm_style", "__norm_size", "款式名称", "尺码", "本次扣减数量"]
    empty_detail_cols = ["扣减记录ID", "来源记录ID", "日期", "款式名称", "尺码", "扣减数量", "来源表", "__norm_style", "__norm_size"]
    if not all_rows:
        return pd.DataFrame(columns=empty_summary_cols), all_warnings, date_check, pd.DataFrame(columns=empty_detail_cols)

    detail = pd.concat(all_rows, ignore_index=True)
    detail["__norm_style"] = detail["款式名称"].map(norm_style)
    detail["__norm_size"] = detail["尺码"].map(norm_size)

    summary = summarize_detail(detail)
    return summary, all_warnings, date_check, detail


# =========================
# 页面：上传和设置
# =========================

with st.sidebar:
    st.header("1）上传 CSV")
    inventory_file = st.file_uploader("库存表 CSV（必传）", type=["csv"])
    deduction_log_file = st.file_uploader(
        "上次下载的扣减日志 CSV（防重复，可选但建议必传）",
        type=["csv"],
        help="第一次使用可以不传；从第二次开始，上传上次下载的扣减日志，程序就会自动跳过以前扣过的记录。",
    )

    st.divider()
    b4g1_file = st.file_uploader("赠送款式 / B4 表 CSV", type=["csv"])
    normal_file = st.file_uploader("水单表 - 新普通水单 CSV", type=["csv"])
    influencer_file = st.file_uploader("深达水单表 CSV", type=["csv"])
    exchange_file = st.file_uploader("达人换货表 CSV", type=["csv"])

    available_dates = sorted(set(
        get_available_dates_from_one_file(b4g1_file)
        + get_available_dates_from_one_file(normal_file)
        + get_available_dates_from_one_file(influencer_file)
        + get_available_dates_from_one_file(exchange_file)
    ))

    st.header("2）选择日期区间")
    st.caption("以后上传的是完整总表也没问题：这里只按你手动选择的日期区间扣减，区间首尾日期都包含。")

    # 上传的是完整总表时，默认不能选“最早日期到最新日期”，否则容易误扣很多天。
    # 所以默认只选 CSV 里能识别到的最新一天；你需要补扣周五-周一时，再手动拉开区间。
    if available_dates:
        max_available_date = max(available_dates)
        default_start = max_available_date
        default_end = max_available_date
    else:
        default_start = date.today()
        default_end = date.today()

    date_range = st.date_input(
        "选择扣减日期区间",
        value=(default_start, default_end),
        help="例如周一补扣时，开始日期选上周五，结束日期选本周一。程序会扣这个区间内所有记录。",
    )

    if isinstance(date_range, tuple) and len(date_range) == 2:
        range_start, range_end = date_range
    else:
        range_start = date_range
        range_end = date_range

    if range_start > range_end:
        range_start, range_end = range_end, range_start

    selected_dates = [d.date() for d in pd.date_range(range_start, range_end, freq="D")]

    st.caption(f"本次将扣减：{range_start.strftime('%Y/%m/%d')} 至 {range_end.strftime('%Y/%m/%d')}，共 {len(selected_dates)} 天。")
    if available_dates:
        st.caption("上传 CSV 实际包含日期：" + "、".join([d.strftime("%Y/%m/%d") for d in available_dates]))
    else:
        st.caption("还没有从上传的水单 CSV 里识别到日期。")

    done_only = st.checkbox("只扣已打包 / 已完成发货记录", value=False)
    update_date_snapshot = st.checkbox("同时新增 / 更新日期库存列", value=True)
    floor_zero = st.checkbox("扣减后库存不低于 0", value=False)

selected_dates = sorted(set(selected_dates))

if not selected_dates:
    st.warning("请选择要扣减的日期区间。")
    st.stop()

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
        "name": "赠送款式/B4表",
        "file": b4g1_file,
        "style_cols": ["赠送款式 Style Names", "Product Name", "款式", "款式名称", "Style Names", "款式 + 库位"],
        "size_cols": ["尺码 (size)", "Size'", "Size", "尺码"],
        "id_cols": ["Order ID", "订单ID", "订单编号", "订单号", "Tracking ID (Last 4 digital)", "Tracking ID", "顾客Handle （customer name in tiktokshop)", "顾客Handle", "customer name"],
    },
    {
        "name": "水单表-新普通水单",
        "file": normal_file,
        "style_cols": ["Product Name", "款式", "款式名称", "Style Names", "赠送款式 Style Names", "款式 + 库位"],
        "size_cols": ["Size'", "Size", "尺码", "尺码 (size)"],
        "id_cols": ["Order ID", "订单ID", "订单编号", "订单号", "查物流 Tracking No.", "Tracking No.", "Tracking ID", "Phone", "Shipping Info"],
    },
    {
        "name": "深达水单表",
        "file": influencer_file,
        "style_cols": ["Product Name", "Product Name1", "款式", "款式名称", "Style Names", "款式 + 库位"],
        "size_cols": ["Size'", "Size", "尺码", "尺码 (size)"],
        "id_cols": ["Order ID", "订单ID", "订单编号", "订单号", "查物流 Tracking No.", "打包 Tracking No.", "Tracking No.", "Handle", "达人Name", "email", "Phone"],
    },
    {
        "name": "达人换货表",
        "file": exchange_file,
        "style_cols": ["发货款式", "Product Name", "款式", "款式名称"],
        "size_cols": ["发货尺码", "Size'", "Size", "尺码", "尺码 (size)"],
        "id_cols": ["Order ID", "订单ID", "订单编号", "订单号", "Tracking ID (Last 4 digital)", "Tracking ID", "达人Handle", "Handle", "原款 SKU", "原款式名称", "原款式尺码"],
    },
]

candidate_summary, warnings, date_check, candidate_detail = build_deduction_summary(sources, selected_dates, done_only)

previous_log = read_previous_deduction_log(deduction_log_file)
processed_ids = set(previous_log["扣减记录ID"].astype(str).map(norm_text).tolist()) if not previous_log.empty else set()

if candidate_detail.empty:
    detail = candidate_detail.copy()
    already_processed_detail = pd.DataFrame(columns=candidate_detail.columns)
else:
    candidate_detail["扣减记录ID"] = candidate_detail["扣减记录ID"].astype(str).map(norm_text)
    already_processed_detail = candidate_detail[candidate_detail["扣减记录ID"].isin(processed_ids)].copy()
    detail = candidate_detail[~candidate_detail["扣减记录ID"].isin(processed_ids)].copy()

summary = summarize_detail(detail)
already_processed_summary = summarize_detail(already_processed_detail)
updated_log = build_updated_deduction_log(previous_log, detail, selected_dates)

st.info(f"本次选择扣减日期：{format_dates_label(selected_dates)}")

if deduction_log_file is None:
    st.warning("你本次没有上传【扣减日志】。第一次使用没关系；但从下一次开始，请上传这次下载的扣减日志，否则同一日期旧记录可能会被重复扣减。")

if summary.empty:
    if not candidate_detail.empty and not already_processed_detail.empty:
        st.success("所选日期区间内的记录都已经在扣减日志里，程序没有重复扣减。")
        if not already_processed_summary.empty:
            st.subheader("已跳过的历史扣减记录")
            st.dataframe(already_processed_summary[["款式名称", "尺码", "本次扣减数量"]], use_container_width=True, hide_index=True)
        st.download_button(
            "下载：更新后的扣减日志 CSV",
            data=to_csv_bytes(updated_log),
            file_name=f"NailVesta_扣减日志_{file_date_label(selected_dates)}.csv",
            mime="text/csv",
        )
    else:
        st.warning("所选日期没有可扣减记录。")

    if not date_check.empty:
        st.subheader("上传文件日期检查")
        st.dataframe(date_check, use_container_width=True, hide_index=True)
    if warnings:
        with st.expander("查看提示", expanded=True):
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

# 用总扣减数量计算最终库存。
summary_for_merge = summary[["__norm_style", "__norm_size", "本次扣减数量"]].copy()
work = work.merge(summary_for_merge, on=["__norm_style", "__norm_size"], how="left")
work["本次扣减数量"] = work["本次扣减数量"].fillna(0).astype(int)
work["__stock_after_raw"] = work["__stock_before"] - work["本次扣减数量"]
negative_mask = work["__stock_after_raw"] < 0

# 如勾选日期库存列，则按日期顺序逐日扣减并写入 05/xx 列。
work["__running_stock"] = work["__stock_before"].copy()
if update_date_snapshot:
    daily_summary = (
        detail.groupby(["日期", "__norm_style", "__norm_size"], as_index=False)
        .agg(当日扣减数量=("扣减数量", "sum"))
    )

    for d in selected_dates:
        one_day = daily_summary[daily_summary["日期"] == d][["__norm_style", "__norm_size", "当日扣减数量"]].copy()
        tmp_qty = work[["__norm_style", "__norm_size"]].merge(
            one_day,
            on=["__norm_style", "__norm_size"],
            how="left",
        )["当日扣减数量"].fillna(0).astype(int)
        work["__running_stock"] = work["__running_stock"] - tmp_qty.values
        if floor_zero:
            work["__running_stock"] = work["__running_stock"].clip(lower=0)
        result[date_col_name(d)] = work["__running_stock"].round(0).astype(int)
else:
    work["__running_stock"] = work["__stock_after_raw"]
    if floor_zero:
        work["__running_stock"] = work["__running_stock"].clip(lower=0)

# 只把最终库存写回原始表，不添加任何诊断列。
result[stock_col] = work["__running_stock"].round(0).astype(int)

available_keys = set(zip(work["__norm_style"], work["__norm_size"]))
summary["__matched"] = summary.apply(lambda r: (r["__norm_style"], r["__norm_size"]) in available_keys, axis=1)
matched_summary = summary[summary["__matched"]].copy()
unmatched_summary = summary[~summary["__matched"]].copy()

show_summary = matched_summary[["款式名称", "尺码", "本次扣减数量"]].sort_values(["款式名称", "尺码"]).reset_index(drop=True)
all_summary_clean = summary[["款式名称", "尺码", "本次扣减数量"]].sort_values(["款式名称", "尺码"]).reset_index(drop=True)

# =========================
# 页面展示：只展示用户需要的两部分
# =========================

st.subheader("1）本次扣减汇总")
st.caption("一个 cell 里有多个款式时，程序会拆成多个款式分别扣减；同一行的尺码会应用到该行所有款式。下面只显示【扣减日志里没有出现过的新记录】。")

m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("本次新增扣减件数", int(show_summary["本次扣减数量"].sum()) if not show_summary.empty else 0)
with m2:
    st.metric("新增款式 + 尺码数", len(show_summary))
with m3:
    st.metric("已跳过旧记录件数", int(already_processed_detail["扣减数量"].sum()) if not already_processed_detail.empty else 0)
with m4:
    st.metric("未匹配项", len(unmatched_summary))

if not show_summary.empty:
    show_summary = show_summary.copy()
    show_summary["本次扣减数量"] = show_summary["本次扣减数量"].astype(int)
    st.table(show_summary)
else:
    st.info("没有成功匹配到库存表的新增扣减项。")

if not already_processed_summary.empty:
    with st.expander("查看已跳过的历史扣减记录（不会重复扣）", expanded=False):
        st.dataframe(
            already_processed_summary[["款式名称", "尺码", "本次扣减数量"]].sort_values(["款式名称", "尺码"]),
            use_container_width=True,
            hide_index=True,
        )

if not unmatched_summary.empty:
    st.error("以下款式 + 尺码没有在库存表中匹配到，所以没有扣减。请检查水单款式名 / 尺码是否和库存表一致。")
    st.dataframe(
        unmatched_summary[["款式名称", "尺码", "本次扣减数量"]].sort_values(["款式名称", "尺码"]),
        use_container_width=True,
        hide_index=True,
    )

if negative_mask.any():
    negative_rows = work.loc[negative_mask, ["__match_style", "__match_size", "__stock_before", "本次扣减数量", "__stock_after_raw"]].copy()
    negative_rows.columns = ["款式名称", "尺码", "扣减前库存", "本次扣减数量", "扣减后库存"]
    st.warning("有库存扣减后小于 0，请重点核对。")
    st.dataframe(negative_rows, use_container_width=True, hide_index=True)

with st.expander("上传文件日期检查"):
    if not date_check.empty:
        st.dataframe(date_check, use_container_width=True, hide_index=True)
    else:
        st.write("没有上传任何水单来源表。")

if warnings:
    with st.expander("查看跳过 / 识别提示", expanded=False):
        for w in warnings:
            st.write("-", w)

st.subheader("2）扣完库存后的最新全部库存")
st.caption("这里保留库存表原本的格式；只更新【当前库存】。如果勾选日期库存列，程序会按日期区间顺序逐日新增 / 更新 05/xx 库存列。")
st.dataframe(result, use_container_width=True, hide_index=True)

st.subheader("下载结果")
col1, col2, col3 = st.columns(3)
with col1:
    st.download_button(
        "下载：扣完库存后的最新库存 CSV",
        data=to_csv_bytes(result),
        file_name=f"NailVesta_库存扣减后_{file_date_label(selected_dates)}.csv",
        mime="text/csv",
    )
with col2:
    st.download_button(
        "下载：本次新增扣减汇总 CSV",
        data=to_csv_bytes(all_summary_clean),
        file_name=f"NailVesta_本次新增扣减汇总_{file_date_label(selected_dates)}.csv",
        mime="text/csv",
    )
with col3:
    st.download_button(
        "下载：更新后的扣减日志 CSV",
        data=to_csv_bytes(updated_log),
        file_name=f"NailVesta_扣减日志_{file_date_label(selected_dates)}.csv",
        mime="text/csv",
    )
