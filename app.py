import re
from datetime import date

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="NailVesta 每日水单库存扣减工具",
    page_icon="💅",
    layout="wide",
)

st.title("💅 NailVesta 每日水单 / 补寄 / 换货库存自动扣减工具")
st.caption("上传水单表 CSV + 库存表 CSV，选择日期后，程序会按【款式名称 + 尺码】汇总并从库存表自动扣减。")


# =========================
# 基础工具函数
# =========================

def read_csv_safely(uploaded_file):
    """Read CSV exported from Lark / Excel. Keep every field as text."""
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


def normalize_col_name(col):
    return str(col).strip().replace("\ufeff", "")


def clean_dataframe(df):
    if df is None:
        return None
    out = df.copy()
    out.columns = [normalize_col_name(c) for c in out.columns]
    return out.fillna("")


def find_col(df, candidates, contains_any=None):
    """Find one column by exact match first, then fuzzy match."""
    if df is None or len(df.columns) == 0:
        return None

    cols = list(df.columns)
    lower_map = {str(c).strip().lower(): c for c in cols}

    for cand in candidates:
        key = cand.strip().lower()
        if key in lower_map:
            return lower_map[key]

    for cand in candidates:
        key = cand.strip().lower()
        for c in cols:
            if key and key in str(c).strip().lower():
                return c

    if contains_any:
        keys = [str(x).lower() for x in contains_any]
        for c in cols:
            lc = str(c).lower()
            if any(k in lc for k in keys):
                return c
    return None


def candidate_cols(df, candidates, contains_any=None):
    """Return possible columns in priority order, deduped."""
    if df is None or len(df.columns) == 0:
        return []
    found = []
    for cand in candidates:
        col = find_col(df, [cand])
        if col and col not in found:
            found.append(col)
    if contains_any:
        keys = [str(x).lower() for x in contains_any]
        for c in df.columns:
            lc = str(c).lower()
            if any(k in lc for k in keys) and c not in found:
                found.append(c)
    return found


def norm_text(x):
    x = str(x).replace("\ufeff", "").replace("\u200b", "").replace("\xa0", " ").strip()
    x = re.sub(r"\s+", " ", x)
    return x


def norm_style(x):
    return norm_text(x).lower()


def norm_size(x):
    x = norm_text(x).upper()
    m = re.search(r"\b(XS|XL|S|M|L)\b", x)
    return m.group(1) if m else x


def size_from_sku(sku):
    sku = norm_text(sku).upper()
    m = re.search(r"(?:-|_|\s)(XS|XL|S|M|L)$", sku)
    if m:
        return m.group(1)
    m = re.search(r"(XS|XL|S|M|L)$", sku)
    return m.group(1) if m else ""


def split_styles(value):
    """Split Lark multi-select cells, e.g. 'Aqua Blush, Cherry Romance'."""
    value = norm_text(value)
    if not value:
        return []

    # 兼容“款式 + 库位”字段：Cherry Romance ｜ 库位：B-04-11,Aqua Blush ｜ 库位：A-02-07
    if "库位" in value and ("｜" in value or "|" in value):
        parts = re.findall(r"([^,，\n;；|｜]+?)\s*[|｜]\s*库位", value)
        if parts:
            return [norm_text(p) for p in parts if norm_text(p)]

    pieces = re.split(r"[,，\n;；]+", value)
    return [norm_text(p) for p in pieces if norm_text(p)]


def parse_date_series(s):
    return pd.to_datetime(s.astype(str).str.strip(), errors="coerce").dt.date


def truthy_series(s):
    text = s.astype(str).str.strip().str.lower()
    return (
        text.isin(["1", "true", "yes", "y", "done", "packed", "shipped", "completed"])
        | text.str.contains("已打包|已完成|完成|发货|packed|shipped|completed", case=False, na=False)
    )


def to_csv_bytes(df):
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def selected_date_col_name(selected_date):
    return selected_date.strftime("%-m/%-d") if hasattr(selected_date, "strftime") else "日期库存"


# =========================
# 库存表处理
# =========================


def prepare_inventory(inv_df):
    """
    支持你的库存表格式：
    产品名称 | 甲型 | SKU编码 | 当前库存 | 05/05
    - 产品名称只有第一行有值，下面 S/M/L 为空：程序会自动向下填充用于匹配。
    - 尺码不单独成列：程序会从 SKU 编码末尾的 -S / -M / -L 自动识别。
    """
    inv_df = clean_dataframe(inv_df)

    sku_col = find_col(inv_df, ["SKU编码", "SKU", "Full SKU", "Variant SKU", "sku"])
    style_col = find_col(inv_df, ["产品名称", "款式名称", "款式", "Product Name", "Style Name", "Style Names", "Name"])
    size_col = find_col(inv_df, ["尺码", "尺码 (size)", "Size", "Size'", "Variant Size"])

    inv_df["__match_style"] = ""
    inv_df["__match_size"] = ""

    if style_col:
        # Lark / Excel 合并单元格导出后，M/L 行的产品名称可能为空；这里自动向下填充。
        inv_df["__match_style"] = inv_df[style_col].replace("", pd.NA).ffill().fillna("").map(norm_text)

    if size_col:
        inv_df["__match_size"] = inv_df[size_col].map(norm_size)
    elif sku_col:
        inv_df["__match_size"] = inv_df[sku_col].map(size_from_sku)

    inv_df["__norm_style"] = inv_df["__match_style"].map(norm_style)
    inv_df["__norm_size"] = inv_df["__match_size"].map(norm_size)

    return inv_df, sku_col, style_col, size_col


# =========================
# 来源表提取
# =========================


def filter_by_selected_date(df, selected_date):
    date_col = find_col(df, ["日期", "date", "Date"])
    if not date_col:
        return df.iloc[0:0].copy(), "没有找到日期列"
    temp = df.copy()
    temp["__parsed_date"] = parse_date_series(temp[date_col])
    return temp[temp["__parsed_date"] == selected_date].copy(), ""


def filter_done_records(df, enabled=True):
    if not enabled or df is None or df.empty:
        return df

    # 有“是否打包 / 是否完成发货 / 状态”才过滤；没有则保留。
    status_candidates = [
        "是否完成发货",
        "是否打包",
        "状态_ (status)",
        "状态 (status)",
        "status",
    ]
    status_cols = [c for c in status_candidates if c in df.columns]
    if not status_cols:
        return df

    mask = pd.Series(False, index=df.index)
    for col in status_cols:
        mask = mask | truthy_series(df[col])
    return df[mask].copy()


def extract_deduction_rows(df, table_name, selected_date, style_candidates, size_candidates, done_only=False):
    df = clean_dataframe(df)
    if df is None or (df.empty and len(df.columns) == 0):
        return pd.DataFrame(), pd.DataFrame()

    dated, date_warning = filter_by_selected_date(df, selected_date)
    if dated.empty:
        return pd.DataFrame(), pd.DataFrame([{
            "表格": table_name,
            "问题": date_warning or "所选日期没有记录",
            "行号": "",
            "款式原值": "",
            "尺码原值": "",
        }])

    dated = filter_done_records(dated, done_only)
    if dated.empty:
        return pd.DataFrame(), pd.DataFrame([{
            "表格": table_name,
            "问题": "所选日期有记录，但没有符合“已打包/已完成发货”条件的记录",
            "行号": "",
            "款式原值": "",
            "尺码原值": "",
        }])

    style_cols = candidate_cols(dated, style_candidates, contains_any=["product", "款式", "style"])
    size_cols = candidate_cols(dated, size_candidates, contains_any=["size", "尺码"])

    warnings = []
    if not style_cols or not size_cols:
        warnings.append({
            "表格": table_name,
            "问题": f"没有找到款式列或尺码列。识别到款式列={style_cols}，尺码列={size_cols}",
            "行号": "",
            "款式原值": "",
            "尺码原值": "",
        })
        return pd.DataFrame(), pd.DataFrame(warnings)

    out = []
    for idx, row in dated.iterrows():
        raw_style = ""
        used_style_col = ""
        for c in style_cols:
            if norm_text(row.get(c, "")):
                raw_style = row.get(c, "")
                used_style_col = c
                break

        raw_size = ""
        used_size_col = ""
        for c in size_cols:
            if norm_text(row.get(c, "")):
                raw_size = row.get(c, "")
                used_size_col = c
                break

        styles = split_styles(raw_style)
        size = norm_size(raw_size)

        if not styles or not size:
            warnings.append({
                "表格": table_name,
                "问题": "款式或尺码为空，已跳过",
                "行号": int(idx) + 2,
                "款式原值": raw_style,
                "尺码原值": raw_size,
            })
            continue

        for style in styles:
            out.append({
                "日期": selected_date.strftime("%Y-%m-%d"),
                "来源表": table_name,
                "源文件行号": int(idx) + 2,
                "款式名称": style,
                "尺码": size,
                "扣减数量": 1,
                "款式列": used_style_col,
                "尺码列": used_size_col,
            })

    return pd.DataFrame(out), pd.DataFrame(warnings)


# =========================
# 页面：上传与设置
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
    floor_zero = st.checkbox("扣减后库存不低于 0（低于 0 显示为 0）", value=True)
    update_date_snapshot = st.checkbox("同时新增 / 更新所选日期库存列", value=True)

st.info("你的库存表可以直接用：程序会自动把空白的【产品名称】向下填充，并从【SKU编码】末尾的 -S / -M / -L 识别尺码。")

if inventory_file is None:
    st.warning("请先在左侧上传库存表 CSV。")
    st.stop()

inventory_raw = read_csv_safely(inventory_file)
inventory, sku_col, inv_style_col, inv_size_col = prepare_inventory(inventory_raw)

if not inv_style_col:
    st.error("库存表没有识别到【产品名称 / 款式名称】列。")
    st.write("库存表当前列名：", list(inventory_raw.columns))
    st.stop()

if not inv_size_col and not sku_col:
    st.error("库存表没有识别到【尺码】列，也没有识别到【SKU编码】列，所以无法判断 S/M/L。")
    st.write("库存表当前列名：", list(inventory_raw.columns))
    st.stop()

if inventory["__norm_style"].eq("").any() or inventory["__norm_size"].eq("").any():
    bad_inv = inventory[inventory["__norm_style"].eq("") | inventory["__norm_size"].eq("")].copy()
    st.warning("库存表有部分行没有识别到产品名称或尺码，这些行不会参与扣减。")
    st.dataframe(bad_inv.head(100), use_container_width=True)

stock_candidates = [
    "当前库存", "库存量", "库存数量", "剩余库存", "实际库存", "可用库存",
    "Available", "On hand", "Quantity", "Qty", "Stock", "Inventory"
]
stock_default = find_col(inventory, stock_candidates, contains_any=["库存", "quantity", "qty", "stock", "inventory", "on hand", "available"])
stock_options = [c for c in inventory.columns if not c.startswith("__")]
default_idx = stock_options.index(stock_default) if stock_default in stock_options else 0

col_a, col_b, col_c, col_d = st.columns(4)
with col_a:
    st.metric("库存表行数", len(inventory))
with col_b:
    st.write("产品名称列：", inv_style_col)
with col_c:
    st.write("SKU列：", sku_col or "未识别")
with col_d:
    st.write("尺码来源：", inv_size_col if inv_size_col else "从 SKU 末尾识别")

stock_col = st.selectbox("请选择要扣减的库存数量列", stock_options, index=default_idx)

# 检查库存 key 是否重复
key_valid = inventory["__norm_style"].ne("") & inventory["__norm_size"].ne("")
key_dup_mask = inventory.duplicated(["__norm_style", "__norm_size"], keep=False) & key_valid
if key_dup_mask.any():
    st.error("库存表中存在重复的【款式名称 + 尺码】。为避免扣错，请先处理重复项后再运行。")
    show_cols = [c for c in [inv_style_col, sku_col, stock_col, "__match_style", "__match_size"] if c and c in inventory.columns]
    st.dataframe(inventory.loc[key_dup_mask, show_cols], use_container_width=True)
    st.stop()

# =========================
# 提取扣减记录
# =========================

source_configs = [
    (
        "水单补寄表/B4G1_B4",
        b4g1_file,
        ["赠送款式 Style Names", "Product Name", "款式", "款式名称", "Style Names", "款式 + 库位"],
        ["尺码 (size)", "Size", "Size'", "尺码"],
    ),
    (
        "水单表-新普通水单",
        normal_file,
        ["Product Name", "款式", "款式名称", "Style Names", "赠送款式 Style Names", "款式 + 库位"],
        ["Size'", "Size", "尺码", "尺码 (size)"],
    ),
    (
        "深达水单表",
        influencer_file,
        ["Product Name", "Product Name1", "款式", "款式名称", "Style Names", "款式 + 库位"],
        ["Size", "Size'", "尺码", "尺码 (size)"],
    ),
    (
        "达人换货表",
        exchange_file,
        ["发货款式", "Product Name", "款式", "款式名称"],
        ["发货尺码", "Size", "Size'", "尺码", "尺码 (size)"],
    ),
]

all_rows = []
all_warnings = []
for table_name, uploaded, style_candidates, size_candidates in source_configs:
    if uploaded is None:
        continue
    source_df = read_csv_safely(uploaded)
    rows, warnings = extract_deduction_rows(
        df=source_df,
        table_name=table_name,
        selected_date=selected_date,
        style_candidates=style_candidates,
        size_candidates=size_candidates,
        done_only=done_only,
    )
    if not rows.empty:
        all_rows.append(rows)
    if not warnings.empty:
        all_warnings.append(warnings)

if not all_rows:
    st.warning("所选日期没有可扣减记录。请检查日期是否正确，或取消勾选“只扣已打包/已完成发货记录”。")
    if all_warnings:
        st.subheader("提示 / 跳过原因")
        st.dataframe(pd.concat(all_warnings, ignore_index=True), use_container_width=True)
    st.stop()

raw_deductions = pd.concat(all_rows, ignore_index=True)
raw_deductions["__norm_style"] = raw_deductions["款式名称"].map(norm_style)
raw_deductions["__norm_size"] = raw_deductions["尺码"].map(norm_size)

summary = (
    raw_deductions
    .groupby(["__norm_style", "__norm_size", "款式名称", "尺码"], as_index=False)
    .agg(
        本次扣减数量=("扣减数量", "sum"),
        来源明细=("来源表", lambda x: ", ".join(sorted(set(x)))),
        来源行号=("源文件行号", lambda x: ", ".join(map(str, sorted(set(x))))),
    )
)

# =========================
# 匹配库存并扣减
# =========================

update = inventory.copy()
update["__stock_before_numeric"] = pd.to_numeric(
    update[stock_col].astype(str).str.replace(",", "", regex=False).str.strip(),
    errors="coerce",
).fillna(0)

summary_for_merge = summary[["__norm_style", "__norm_size", "本次扣减数量", "来源明细"]]
update = update.merge(summary_for_merge, on=["__norm_style", "__norm_size"], how="left")
update["本次扣减数量"] = update["本次扣减数量"].fillna(0).astype(int)
update["扣减前库存"] = update["__stock_before_numeric"]
update["扣减后库存"] = update["扣减前库存"] - update["本次扣减数量"]
update["是否扣减后为负"] = update["扣减后库存"] < 0
if floor_zero:
    update["扣减后库存"] = update["扣减后库存"].clip(lower=0)

# 覆盖用户选择的库存列
update[stock_col] = update["扣减后库存"].round(0).astype(int)

# 可选：新增 / 更新日期库存列，例如 5/6
if update_date_snapshot:
    day_col = selected_date_col_name(selected_date)
    update[day_col] = update["扣减后库存"].round(0).astype(int)

matched_keys = set(zip(
    update.loc[update["本次扣减数量"] > 0, "__norm_style"],
    update.loc[update["本次扣减数量"] > 0, "__norm_size"],
))
summary["是否匹配到库存表"] = summary.apply(lambda r: (r["__norm_style"], r["__norm_size"]) in matched_keys, axis=1)
unmatched = summary[~summary["是否匹配到库存表"]].copy()
matched_summary = summary[summary["是否匹配到库存表"]].copy()

# 导出时删除内部字段，保留审核字段
internal_cols = [c for c in update.columns if c.startswith("__")]
updated_inventory_export = update.drop(columns=internal_cols, errors="ignore")
summary_export = summary.drop(columns=["__norm_style", "__norm_size"], errors="ignore")
raw_export = raw_deductions.drop(columns=["__norm_style", "__norm_size"], errors="ignore")
unmatched_export = unmatched.drop(columns=["__norm_style", "__norm_size"], errors="ignore")

# =========================
# 页面展示
# =========================

st.subheader("扣减结果总览")
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("来源明细行数", len(raw_deductions))
with col2:
    st.metric("扣减总件数", int(summary["本次扣减数量"].sum()))
with col3:
    st.metric("匹配成功款式+尺码", int(matched_summary.shape[0]))
with col4:
    st.metric("未匹配款式+尺码", int(unmatched.shape[0]))

st.subheader("扣减汇总")
st.dataframe(
    summary_export.sort_values(["是否匹配到库存表", "款式名称", "尺码"], ascending=[True, True, True]),
    use_container_width=True,
)

if all_warnings:
    with st.expander("查看跳过 / 识别提示"):
        st.dataframe(pd.concat(all_warnings, ignore_index=True), use_container_width=True)

if not unmatched_export.empty:
    st.error("有扣减项没有匹配到库存表，以下项目不会被扣减。请检查款式名 / 尺码是否和库存表一致。")
    st.dataframe(unmatched_export, use_container_width=True)

neg_rows = updated_inventory_export[updated_inventory_export["是否扣减后为负"] == True]
if not neg_rows.empty:
    st.warning("有库存扣减后会小于 0，已在结果中标记“是否扣减后为负”。")
    show_cols = [c for c in [inv_style_col, sku_col, stock_col, "本次扣减数量", "扣减前库存", "扣减后库存", "是否扣减后为负"] if c and c in neg_rows.columns]
    st.dataframe(neg_rows[show_cols], use_container_width=True)

st.subheader("更新后的库存表预览")
st.dataframe(updated_inventory_export.head(300), use_container_width=True)

st.subheader("下载结果")
col_d1, col_d2, col_d3 = st.columns(3)
with col_d1:
    st.download_button(
        "下载：已扣减库存表 CSV",
        data=to_csv_bytes(updated_inventory_export),
        file_name=f"updated_inventory_{selected_date.strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )
with col_d2:
    st.download_button(
        "下载：扣减明细 CSV",
        data=to_csv_bytes(raw_export),
        file_name=f"deduction_detail_{selected_date.strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )
with col_d3:
    st.download_button(
        "下载：未匹配项目 CSV",
        data=to_csv_bytes(unmatched_export),
        file_name=f"unmatched_deductions_{selected_date.strftime('%Y%m%d')}.csv",
        mime="text/csv",
        disabled=unmatched_export.empty,
    )
