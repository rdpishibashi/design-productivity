# Design-productivity 技術文書

## 目的

プロジェクトごとの**差分要素数と不具合対応工数**を集計し、設計変更の生産性指標を算出する Streamlit アプリ。  
工数データ（merged_efforts.xlsx）と DXF 差分台帳（merged_ledger.xlsx）を統合して分析する。

---

## ファイル構成

```
Design-productivity/
├── app.py               # アプリ本体（全ロジック）
├── requirements.txt     # 依存パッケージ
├── TECHNICAL.md         # 本文書
├── merged_efforts.xlsx  # デフォルト工数データ（Effort-analyzer で生成）
├── merged_ledger.xlsx   # デフォルト台帳データ（DXF-diff-merger で生成）
└── .streamlit/
    └── config.toml      # Streamlit テーマ設定
```

---

## 入力ファイル仕様

### 工数データ（merged_efforts.xlsx）

Effort-analyzer の Job Organizer で生成した統合済み工数 Excel。  
全シートを結合して使用する。

| 使用カラム | 内容 |
|-----------|------|
| `WBS要素(代入)` | プロジェクト番号（結合キー） |
| `USER_FIELD_01` | 工数種別（`不具合対応` でフィルタリング） |
| `USER_FIELD_02` | 要因種別（`電気設計要因` / `サプライヤー要因` を集計） |
| `作業時間(h)` | 工数（時間単位） |

### 台帳データ（merged_ledger.xlsx）

DXF-diff-merger で生成した統合済み台帳 Excel。  
`Merged Data` シートを使用する（存在しない場合は先頭シート）。

| 使用カラム | 内容 |
|-----------|------|
| `Project Number` | プロジェクト番号（結合キー） |
| `Child` / `Parent` | 図番ペア（重複排除に使用） |
| `Recorded Date` | 記録日時（重複排除の優先順位に使用） |
| `Deleted / Added / Diff / Unchanged / Total Entities` | エンティティ数（集計対象） |

---

## 処理ロジック

### 処理フロー

```
[merged_efforts.xlsx]         [merged_ledger.xlsx]
        ↓                              ↓
  全シートを concat              Merged Data シート読込
        ↓                              ↓
  USER_FIELD_01=不具合対応 でフィルタ   Child/Parent が null の行を除外
        ↓                              ↓
  WBS要素(代入) × USER_FIELD_02 でピボット   Recorded Date 降順でソート
        ↓                              ↓
  電気設計要因・サプライヤー要因の時間を集計    Child+Parent 単位で drop_duplicates
        ↓                              ↓
  不具合対応工数[h] を算出            指番ごとに要素数を集計
              ↓               ↓
          外部結合（outer join on 指番）
                      ↓
          差分要素割合[%] = 差分要素数 / 合計要素数 × 100
          差分要素数／不具合対応工数[h] = 差分要素数 / 不具合対応工数[h]
                      ↓
              Excel 出力（生産性分析シート）
```

### 重複排除（台帳データ）

クロスプロジェクト重複（`Duplicate Projects` 列が空でない行）は、  
同一 Child+Parent ペアのうち **Recorded Date が最新のプロジェクトのエントリーを採用**する。

```python
dedup = (
    df.sort_values("Recorded Date", ascending=False, na_position="last")
    .drop_duplicates(subset=["Child", "Parent"], keep="first")
)
```

### 片側のみのプロジェクト

工数データまたは台帳データの一方にしか存在しないプロジェクトも出力に含める。  
欠損値は数値列に 0 を補完する。分母が 0 の場合、列10・列11は空白（NaN）となる。

---

## 出力 Excel フォーマット

シート名: **生産性分析**（1シート）

| # | 列名 | 型 | Excel 書式 | 備考 |
|---|------|----|-----------|------|
| 1 | 指番 | 文字列 | — | WBS要素(代入) / Project Number |
| 2 | 電気設計要因不具合対応工数[h] | 浮動小数点 | `#,##0.00` | |
| 3 | サプライヤー要因不具合対応工数[h] | 浮動小数点 | `#,##0.00` | |
| 4 | 不具合対応工数[h] | 浮動小数点 | `#,##0.00` | 2+3の合計 |
| 5 | 削除要素数 | 整数 | `#,##0` | |
| 6 | 追加要素数 | 整数 | `#,##0` | |
| 7 | 差分要素数 | 整数 | `#,##0` | |
| 8 | 不変要素数 | 整数 | `#,##0` | |
| 9 | 合計要素数 | 整数 | `#,##0` | |
| 10 | 差分要素割合[%] | 浮動小数点 | `#,##0.00"%"` | 7÷9×100、合計要素数=0の場合は空白 |
| 11 | 差分要素数／不具合対応工数[h] | 浮動小数点 | `#,##0.00` | 7÷4、不具合対応工数=0の場合は空白 |

書式:
- ヘッダー行: 青背景（#4472C4）・白文字・太字
- 先頭行固定（freeze_panes）
- 千の位区切り（`#,##0` 系書式）を全数値列に適用

### ブラウザ表示（st.dataframe）

`pandas.Styler` を使用し、Excel と同一の書式でブラウザ上に表示する。

| 列種別 | Python フォーマット | 表示例 |
|--------|-------------------|--------|
| 浮動小数点 | `"{:,.2f}"` | `1,234.56` |
| 整数 | `"{:,}"` | `106,201` |
| 差分要素割合[%] | `"{:,.2f}%"` | `12.91%` |
| NaN | `""` （空白） | |

> `st.column_config.NumberColumn` は千の位区切りをサポートしないため、`result.style.format(DISPLAY_FORMATS, na_rep="")` を使用する。

---

## 定数・設定値

`app.py` 冒頭の定数を変更することで主要な動作を調整できる。

```python
UF1_TARGET   = "不具合対応"      # USER_FIELD_01 のフィルタ条件
UF2_ELEC     = "電気設計要因"    # USER_FIELD_02 の集計対象1
UF2_SUPPLIER = "サプライヤー要因"  # USER_FIELD_02 の集計対象2

ENTITY_COL_MAP   # 台帳ソース列名 → 出力列名のマッピング（dict）
COL_FORMATS      # 列名 → Excel number format（dict）
DISPLAY_FORMATS  # 列名 → Python フォーマット文字列（dict、ブラウザ表示用）
```

---

## 依存パッケージ

| パッケージ | 役割 |
|-----------|------|
| `streamlit >= 1.30.0` | UI フレームワーク |
| `pandas >= 2.0.0` | データ処理・Styler によるブラウザ表示書式 |
| `openpyxl >= 3.0.0` | Excel 読み込み |
| `xlsxwriter >= 3.0.0` | Excel 書き込み・書式設定 |

---

## Streamlit Cloud へのデプロイ

1. このディレクトリを GitHub リポジトリにプッシュ
2. [Streamlit Cloud](https://share.streamlit.io/) でリポジトリを接続
3. Main file: `Design-productivity/app.py`（リポジトリルートからの相対パス）
4. `requirements.txt` が自動インストールされる

ローカル実行:

```bash
cd Design-productivity
pip install -r requirements.txt
streamlit run app.py
```

---

## 保守上の注意点

### USER_FIELD の条件を変更したい場合

`app.py` の定数 `UF1_TARGET` / `UF2_ELEC` / `UF2_SUPPLIER` を変更する。

### 台帳の列名が変わった場合

`ENTITY_COL_MAP` のキー（ソース列名）を更新する。値（出力列名）を変更した場合は、  
`OUT_COLS` / `COL_FORMATS` / `DISPLAY_FORMATS` / `COL_WIDTHS` も合わせて更新する。

### 出力列を追加したい場合

1. `OUT_COLS` に列名を追加
2. `build_output` 内で算出ロジックを追加
3. `COL_FORMATS` に Excel 書式を追加
4. `DISPLAY_FORMATS` にブラウザ用書式を追加
5. `COL_WIDTHS` に列幅を追加

### Streamlit の `width` パラメータ

`use_container_width=True/False` は 2025-12-31 に削除予定。  
`width='stretch'` / `width='content'` を使用すること。

---

## 改訂履歴

| 日付 | 変更内容 |
|------|----------|
| 2026-05-20 | 初版作成 |
| 2026-05-20 | 出力列を11列に変更。列10を差分要素割合[%]（差分/合計×100）、列11を差分要素数／不具合対応工数[h]（差分/工数）に変更。列名を全て日本語化。千の位区切りを Excel・ブラウザ表示の両方に適用。 |
