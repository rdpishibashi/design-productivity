import io
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ---- Constants ----

DEFAULT_EFFORTS_PATH = Path(__file__).parent / "merged_efforts.xlsx"
DEFAULT_LEDGER_PATH = Path(__file__).parent / "統合図面管理台帳.xlsx"

EFFORT_COL = "作業時間(h)"
PROJECT_COL = "WBS要素(代入)"
UF1_COL = "USER_FIELD_01"
UF2_COL = "USER_FIELD_02"
UF1_TARGET = "不具合対応"
UF2_ELEC = "電気設計要因"
UF2_SUPPLIER = "サプライヤー要因"

# Source column names in the 台帳データ Summary sheet (already aggregated per 指番)
LEDGER_SUMMARY_COLS = [
    "指番",
    "削除図形総数",
    "追加図形総数",
    "変更図形総数",
    "図形総数",
    "図形変更率 [%]",
    "差分ペア総数",
    "完全新規図面数",
    "指番図面総数",
    "流用率 [%]",
    "新規作成率 [%]",
]

OUT_COLS = [
    "指番",
    "削除図形総数",
    "追加図形総数",
    "変更図形総数",
    "図形総数",
    "図形変更率[%]",
    "差分ペア総数",
    "完全新規図面数",
    "指番図面総数",
    "流用率[%]",
    "新規作成率[%]",
    "電気設計要因不具合対応工数[h]",
    "サプライヤー要因不具合対応工数[h]",
    "不具合対応工数[h]",
    "図形変更効率[図形変更数/h]",
    "図面変更作業効率[図番数/h]",
]

INT_COLS = ["削除図形総数", "追加図形総数", "変更図形総数", "図形総数", "差分ペア総数", "完全新規図面数", "指番図面総数"]
HOUR_COLS = ["電気設計要因不具合対応工数[h]", "サプライヤー要因不具合対応工数[h]", "不具合対応工数[h]"]
EFFICIENCY_COLS = ["図形変更効率[図形変更数/h]", "図面変更作業効率[図番数/h]"]
PERCENT_COLS = ["図形変更率[%]", "流用率[%]", "新規作成率[%]"]

COL_WIDTHS = {
    "指番": 18,
    "削除図形総数": 12,
    "追加図形総数": 12,
    "変更図形総数": 12,
    "図形総数": 12,
    "図形変更率[%]": 14,
    "差分ペア総数": 12,
    "完全新規図面数": 14,
    "指番図面総数": 12,
    "流用率[%]": 12,
    "新規作成率[%]": 14,
    "電気設計要因不具合対応工数[h]": 28,
    "サプライヤー要因不具合対応工数[h]": 30,
    "不具合対応工数[h]": 20,
    "図形変更効率[図形変更数/h]": 24,
    "図面変更作業効率[図番数/h]": 24,
}

# column name -> Python format string for browser display
DISPLAY_FORMATS = {
    **{col: "{:,}" for col in INT_COLS},
    **{col: "{:,.2f}" for col in HOUR_COLS},
    **{col: "{:,.2f}" for col in EFFICIENCY_COLS},
    **{col: "{:.2%}" for col in PERCENT_COLS},
}

# column name -> Excel number format
COL_FORMATS = {
    **{col: "#,##0" for col in INT_COLS},
    **{col: "#,##0.00" for col in HOUR_COLS},
    **{col: "#,##0.00" for col in EFFICIENCY_COLS},
    **{col: "0.00%" for col in PERCENT_COLS},
}


# ---- Processing ----

def load_efforts(source) -> pd.DataFrame:
    """Load 工数データ (single sheet, WBS要素(代入)/USER_FIELD_01/02/作業時間(h))."""
    excel_data = pd.read_excel(source, sheet_name=None)
    df = pd.concat(excel_data.values(), ignore_index=True)
    df.columns = df.columns.str.strip()
    mask = df[PROJECT_COL].notna()
    df.loc[mask, PROJECT_COL] = df.loc[mask, PROJECT_COL].astype(str).str.strip()
    return df


def calc_effort_by_project(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate 電気設計要因 and サプライヤー要因 hours per 指番 (USER_FIELD_01=不具合対応)."""
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
    """Load the pre-aggregated Summary sheet from 台帳データ (指番ごとの指標が既に集計済み)."""
    df = pd.read_excel(source, sheet_name="Summary")
    df = df[LEDGER_SUMMARY_COLS].copy()
    df.columns = [
        "指番",
        "削除図形総数", "追加図形総数", "変更図形総数", "図形総数",
        "図形変更率[%]",
        "差分ペア総数", "完全新規図面数", "指番図面総数",
        "流用率[%]", "新規作成率[%]",
    ]
    df["指番"] = df["指番"].astype(str).str.strip()
    for col in INT_COLS:
        df[col] = df[col].astype(int)
    return df


def build_output(effort_df: pd.DataFrame, ledger_df: pd.DataFrame) -> pd.DataFrame:
    """Inner join 台帳データ(Summary) and 工数データ on 指番, keep only 指番 present in both, compute derived columns."""
    merged = pd.merge(ledger_df, effort_df, on="指番", how="inner")

    # 図形変更効率[図形変更数/h] = 変更図形総数 / 不具合対応工数[h]
    def _shape_efficiency(row):
        effort = row["不具合対応工数[h]"]
        return float(row["変更図形総数"]) / effort if effort != 0 else float("nan")

    # 図面変更作業効率[図番数/h] = 差分ペア総数 / 不具合対応工数[h]
    def _drawing_efficiency(row):
        effort = row["不具合対応工数[h]"]
        return float(row["差分ペア総数"]) / effort if effort != 0 else float("nan")

    merged["図形変更効率[図形変更数/h]"] = merged.apply(_shape_efficiency, axis=1)
    merged["図面変更作業効率[図番数/h]"] = merged.apply(_drawing_efficiency, axis=1)

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


# ---- Graphs ----

def make_reuse_rate_chart(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_bar(name="図形変更率[%]", x=df["指番"], y=df["図形変更率[%]"] * 100)
    fig.add_bar(name="流用率[%]", x=df["指番"], y=df["流用率[%]"] * 100)
    fig.add_bar(name="新規作成率[%]", x=df["指番"], y=df["新規作成率[%]"] * 100)
    fig.update_layout(barmode="group", yaxis_title="%", xaxis_title="指番")
    return fig


def make_recovery_effort_chart(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_bar(name="電気設計要因不具合対応工数[h]", x=df["指番"], y=df["電気設計要因不具合対応工数[h]"])
    fig.add_bar(name="サプライヤー要因不具合対応工数[h]", x=df["指番"], y=df["サプライヤー要因不具合対応工数[h]"])
    fig.update_layout(barmode="stack", yaxis_title="工数 [h]", xaxis_title="指番")
    return fig


def make_recovery_effectiveness_charts(df: pd.DataFrame) -> tuple[go.Figure, go.Figure]:
    fig_shape = go.Figure()
    fig_shape.add_bar(name="図形変更効率[図形変更数/h]", x=df["指番"], y=df["図形変更効率[図形変更数/h]"])
    fig_shape.update_layout(yaxis_title="図形変更数/h", xaxis_title="指番")

    fig_drawing = go.Figure()
    fig_drawing.add_bar(name="図面変更作業効率[図番数/h]", x=df["指番"], y=df["図面変更作業効率[図番数/h]"])
    fig_drawing.update_layout(yaxis_title="図番数/h", xaxis_title="指番")

    return fig_shape, fig_drawing


# ---- UI ----

st.set_page_config(page_title="設計生産性分析", page_icon="📊", layout="wide")
st.title("設計生産性分析")
st.caption("プロジェクトごとの流用率・不具合対応工数・不具合対応効率を集計します。")

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
    st.subheader("台帳データ（統合図面管理台帳.xlsx）")
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
            ledger_df = load_ledger(ledger_source)
        except Exception as e:
            st.error(f"台帳データの読み込みに失敗しました: {e}")
            st.stop()

        result = build_output(effort_df, ledger_df)

    if result.empty:
        st.warning("台帳データと工数データの双方に存在する指番がありませんでした。")
        st.stop()

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

    st.subheader("流用率")
    st.plotly_chart(make_reuse_rate_chart(result), width="stretch")

    st.subheader("不具合対応工数")
    st.plotly_chart(make_recovery_effort_chart(result), width="stretch")

    st.subheader("不具合対応効率")
    eff_col_l, eff_col_r = st.columns(2)
    fig_shape, fig_drawing = make_recovery_effectiveness_charts(result)
    with eff_col_l:
        st.plotly_chart(fig_shape, width="stretch")
    with eff_col_r:
        st.plotly_chart(fig_drawing, width="stretch")

    st.subheader("分析結果")
    st.dataframe(
        result.style.format(DISPLAY_FORMATS, na_rep=""),
        width="stretch",
        hide_index=True,
    )
