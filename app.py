import io
from pathlib import Path

import pandas as pd
import streamlit as st

# ---- Constants ----

DEFAULT_EFFORTS_PATH = Path(__file__).parent / "merged_efforts.xlsx"
DEFAULT_LEDGER_PATH = Path(__file__).parent / "merged_ledger.xlsx"

EFFORT_COL = "作業時間(h)"
PROJECT_COL = "WBS要素(代入)"
UF1_COL = "USER_FIELD_01"
UF2_COL = "USER_FIELD_02"
UF1_TARGET = "不具合対応"
UF2_ELEC = "電気設計要因"
UF2_SUPPLIER = "サプライヤー要因"

# Source column names in ledger → output column names
ENTITY_COL_MAP = {
    "Deleted Entities": "削除要素数",
    "Added Entities":   "追加要素数",
    "Diff Entities":    "差分要素数",
    "Unchanged Entities": "不変要素数",
    "Total Entities":   "合計要素数",
}

OUT_COLS = [
    "指番",
    "電気設計要因不具合対応工数[h]",
    "サプライヤー要因不具合対応工数[h]",
    "不具合対応工数[h]",
    "削除要素数",
    "追加要素数",
    "差分要素数",
    "不変要素数",
    "合計要素数",
    "差分要素割合[%]",
    "差分要素数[個]／不具合対応工数[h]",
]

COL_WIDTHS = {
    "指番": 18,
    "電気設計要因不具合対応工数[h]": 28,
    "サプライヤー要因不具合対応工数[h]": 30,
    "不具合対応工数[h]": 20,
    "削除要素数": 12,
    "追加要素数": 12,
    "差分要素数": 12,
    "不変要素数": 12,
    "合計要素数": 12,
    "差分要素割合[%]": 18,
    "差分要素数[個]／不具合対応工数[h]": 30,
}

# column name → Python format string for browser display
DISPLAY_FORMATS = {
    "電気設計要因不具合対応工数[h]":     "{:,.2f}",
    "サプライヤー要因不具合対応工数[h]": "{:,.2f}",
    "不具合対応工数[h]":               "{:,.2f}",
    "削除要素数":                      "{:,}",
    "追加要素数":                      "{:,}",
    "差分要素数":                      "{:,}",
    "不変要素数":                      "{:,}",
    "合計要素数":                      "{:,}",
    "差分要素割合[%]":                 "{:,.2f}%",
    "差分要素数[個]／不具合対応工数[h]":    "{:,.2f}",
}

# column name → Excel number format
COL_FORMATS = {
    "電気設計要因不具合対応工数[h]":     "#,##0.00",
    "サプライヤー要因不具合対応工数[h]": "#,##0.00",
    "不具合対応工数[h]":               "#,##0.00",
    "削除要素数":                      "#,##0",
    "追加要素数":                      "#,##0",
    "差分要素数":                      "#,##0",
    "不変要素数":                      "#,##0",
    "合計要素数":                      "#,##0",
    "差分要素割合[%]":                 '#,##0.00"%"',
    "差分要素数[個]／不具合対応工数[h]":    "#,##0.00",
}


# ---- Processing ----

def load_efforts(source) -> pd.DataFrame:
    """Load all sheets and concat from effort Excel."""
    excel_data = pd.read_excel(source, sheet_name=None)
    df = pd.concat(excel_data.values(), ignore_index=True)
    df.columns = df.columns.str.strip()
    return df


def calc_effort_by_project(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate 電気設計要因 and サプライヤー要因 hours per project (USER_FIELD_01=不具合対応)."""
    sub = df[df[UF1_COL] == UF1_TARGET].copy()

    if sub.empty:
        return pd.DataFrame(columns=[
            "指番", "電気設計要因不具合対応工数[h]",
            "サプライヤー要因不具合対応工数[h]", "不具合対応工数[h]",
        ])

    pivot = (
        sub.groupby([PROJECT_COL, UF2_COL])[EFFORT_COL]
        .sum()
        .unstack(fill_value=0.0)
        .reset_index()
    )

    for col in [UF2_ELEC, UF2_SUPPLIER]:
        if col not in pivot.columns:
            pivot[col] = 0.0

    result = pivot[[PROJECT_COL, UF2_ELEC, UF2_SUPPLIER]].copy()
    result.columns = ["指番", "電気設計要因不具合対応工数[h]", "サプライヤー要因不具合対応工数[h]"]
    result["不具合対応工数[h]"] = (
        result["電気設計要因不具合対応工数[h]"] + result["サプライヤー要因不具合対応工数[h]"]
    )
    return result


def load_ledger(source) -> pd.DataFrame:
    """Load Merged Data sheet from ledger Excel."""
    try:
        return pd.read_excel(source, sheet_name="Merged Data")
    except Exception:
        return pd.read_excel(source, sheet_name=0)


def calc_ledger_by_project(df: pd.DataFrame) -> pd.DataFrame:
    """Deduplicate cross-project duplicates (keep newest Recorded Date) and aggregate per project."""
    df = df.copy()
    df["Recorded Date"] = pd.to_datetime(df["Recorded Date"], errors="coerce")
    df = df[df["Child"].notna() & df["Parent"].notna()]

    # For the same (Child, Parent) across projects, keep the row with the newest Recorded Date
    dedup = (
        df.sort_values("Recorded Date", ascending=False, na_position="last")
        .drop_duplicates(subset=["Child", "Parent"], keep="first")
    )

    summary = (
        dedup.groupby("Project Number", sort=True)
        .agg(**{dst: (src, "sum") for src, dst in ENTITY_COL_MAP.items()})
        .reset_index()
        .rename(columns={"Project Number": "指番"})
    )
    for col in ENTITY_COL_MAP.values():
        summary[col] = summary[col].astype(int)
    return summary


def build_output(effort_df: pd.DataFrame, ledger_df: pd.DataFrame) -> pd.DataFrame:
    """Outer join effort and ledger data and compute derived columns."""
    merged = pd.merge(effort_df, ledger_df, on="指番", how="outer")

    for col in ["電気設計要因不具合対応工数[h]", "サプライヤー要因不具合対応工数[h]", "不具合対応工数[h]"]:
        merged[col] = merged[col].fillna(0.0)
    for col in ENTITY_COL_MAP.values():
        merged[col] = merged[col].fillna(0).astype(int)

    # 差分要素割合[%] = 差分要素数 / 合計要素数 × 100
    def _pct(row):
        total = row["合計要素数"]
        return float(row["差分要素数"]) / total * 100 if total != 0 else float("nan")

    # 差分要素数[個]／不具合対応工数[h]
    def _rate(row):
        effort = row["不具合対応工数[h]"]
        return float(row["差分要素数"]) / effort if effort != 0 else float("nan")

    merged["差分要素割合[%]"] = merged.apply(_pct, axis=1)
    merged["差分要素数[個]／不具合対応工数[h]"] = merged.apply(_rate, axis=1)

    return merged[OUT_COLS].sort_values("指番").reset_index(drop=True)


def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="生産性分析", index=False)
        wb = writer.book
        ws = writer.sheets["生産性分析"]

        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#4472C4", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter",
        })

        ws.set_row(0, 20, header_fmt)
        ws.freeze_panes(1, 0)

        for i, col in enumerate(df.columns):
            num_fmt = COL_FORMATS.get(col)
            cell_fmt = wb.add_format({"num_format": num_fmt}) if num_fmt else None
            ws.set_column(i, i, COL_WIDTHS.get(col, 15), cell_fmt)

    return buf.getvalue()


# ---- UI ----

st.set_page_config(page_title="設計生産性分析", page_icon="📊", layout="wide")
st.title("設計生産性分析")
st.caption("プロジェクトごとの差分要素数と不具合対応工数を集計します。")

col_l, col_r = st.columns(2)

with col_l:
    st.subheader("工数データ（merged_efforts.xlsx）")
    if DEFAULT_EFFORTS_PATH.exists():
        st.info(f"デフォルトファイルを使用: `{DEFAULT_EFFORTS_PATH.name}`")
    efforts_upload = st.file_uploader(
        "上書きする場合はここにドロップ" if DEFAULT_EFFORTS_PATH.exists() else "ファイルを選択（必須）",
        type=["xlsx"],
        key="efforts",
    )
    efforts_source = efforts_upload if efforts_upload else (DEFAULT_EFFORTS_PATH if DEFAULT_EFFORTS_PATH.exists() else None)

with col_r:
    st.subheader("台帳データ（merged_ledger.xlsx）")
    if DEFAULT_LEDGER_PATH.exists():
        st.info(f"デフォルトファイルを使用: `{DEFAULT_LEDGER_PATH.name}`")
    ledger_upload = st.file_uploader(
        "上書きする場合はここにドロップ" if DEFAULT_LEDGER_PATH.exists() else "ファイルを選択（必須）",
        type=["xlsx"],
        key="ledger",
    )
    ledger_source = ledger_upload if ledger_upload else (DEFAULT_LEDGER_PATH if DEFAULT_LEDGER_PATH.exists() else None)

if not efforts_source or not ledger_source:
    missing = []
    if not efforts_source:
        missing.append("工数データ")
    if not ledger_source:
        missing.append("台帳データ")
    st.warning(f"ファイルが必要です: {', '.join(missing)}")
    st.stop()

if st.button("分析実行", type="primary", width="stretch"):
    with st.spinner("処理中..."):
        try:
            effort_raw = load_efforts(efforts_source)
            effort_df = calc_effort_by_project(effort_raw)
        except Exception as e:
            st.error(f"工数データの読み込みに失敗しました: {e}")
            st.stop()

        try:
            ledger_raw = load_ledger(ledger_source)
            ledger_df = calc_ledger_by_project(ledger_raw)
        except Exception as e:
            st.error(f"台帳データの読み込みに失敗しました: {e}")
            st.stop()

        result = build_output(effort_df, ledger_df)
        excel_bytes = to_excel(result)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("指番数", len(result))
    c2.metric("電気設計要因不具合対応工数 合計[h]", f"{result['電気設計要因不具合対応工数[h]'].sum():.2f}")
    c3.metric("サプライヤー要因不具合対応工数 合計[h]", f"{result['サプライヤー要因不具合対応工数[h]'].sum():.2f}")
    c4.metric("不具合対応工数 合計[h]", f"{result['不具合対応工数[h]'].sum():.2f}")

    st.download_button(
        "📥 Excel をダウンロード",
        data=excel_bytes,
        file_name="design_productivity.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

    st.subheader("分析結果")
    st.dataframe(
        result.style.format(DISPLAY_FORMATS, na_rep=""),
        width="stretch",
        hide_index=True,
    )
