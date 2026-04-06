# Allシート列定義の可変化 — 設計書

**日付:** 2026-04-06  
**対象ファイル:** modConfig.bas, modDataProcess.bas, modAggregation.bas, modSetup.bas

---

## 概要

現在、Allシートの列構成（列数・列順・列名）はVBAコード内の定数（`ALL_COL_*`, `ALL_TOTAL_COLS`）として固定されている。  
これを、ConfigシートのヘッダーI列「Allシート列名」で制御できるよう変更する。

---

## 要件

- Configシートの名寄せテーブルにI列「Allシート列名」を追加する
- I列に値がある行だけ、その行の順番でAllシートの列として出力する
- I列が空白の行はAllシートに出力しない
- 計算列（製品名・口銭按分）とソースファイル名は常にAllシートの末尾に固定で追加する
- VBAコードを修正せずに、Configシートの変更だけで列の追加・削除・並べ替えが可能になること

---

## Configシートの変更

### 名寄せテーブル（G〜I列）

| 列 | 内容 | 変更 |
|----|------|------|
| G | 正規名 | 変更なし |
| H | 対応列名（カンマ区切りエイリアス） | 変更なし |
| I | **Allシート列名**（新規） | 追加 |

- I列に値があれば → Allシートにその列名で出力
- I列が空白 → Allシートに出力しない
- **行の並び順 = Allシートの列順**

### デフォルトサンプル（SetupConfigSheet初期データ）

| G（正規名） | H（エイリアス） | I（Allシート列名） |
|---|---|---|
| 客先名 | 得意先名,得意先コード,客先名 | 客先名 |
| 製品コード | 品番,ProductCode | 製品コード |
| 売上金額 | 金額,Amount,売上高 | 売上金額 |
| 単価単位 | 単価,価格 | 単価単位 |
| 売上数量 | 数量,Qty | 売上数量 |
| 売上発生日 | 日付,発注日,Date | 売上発生日 |
| 売上種別 | 売上区分,SaleType | 売上種別 |
| 部署 | 部門,Dept | 部署 |

### Allシート最終列構造

```
[I列で定義したN列（行順）] + 製品名 + 口銭按分 + ソースファイル名
```

---

## modConfigの変更

### 定数の追廃

| 変更 | 定数名 |
|------|--------|
| 追加 | `CFG_ALL_COL_NAME As Integer = 9`（I列） |
| 削除 | `ALL_COL_CLIENT` 〜 `ALL_COL_SOURCE`（11個） |
| 削除 | `ALL_TOTAL_COLS` |
| 残す | `HDR_*` 定数（ヘッダー文字列、GetAllColIndexの検索キーに使用） |

### 新関数: `LoadAllColDef() As Object`

- 名寄せテーブル（CFG_HEADER_HDR_ROW+1行目から）を読む
- I列（CFG_ALL_COL_NAME）が空白でない行だけを挿入順でDictionaryに格納
- 戻り値: `key = 正規名（G列）`, `value = Allシート列名（I列）`
- 挿入順がそのままAllシートの列順になる

### 新関数: `GetAllColIndex(wsAll As Worksheet, headerName As String) As Integer`

- Allシートの1行目を全列スキャン
- `headerName` に一致するセルの列番号（1-based）を返す
- 見つからない場合は `0` を返す
- modAggregation・CollectUniqueDeptsが `ALL_COL_*` 定数の代わりに使用する

---

## modDataProcessの変更

### `BuildAllSheet`

1. `LoadAllColDef()` を呼び出し `dictAllColDef` を取得（N件）
2. ヘッダー行を動的に書き込み：
   - 列1〜N: `dictAllColDef` の value を順番に書き込み
   - 列N+1: `HDR_PROD_NAME`
   - 列N+2: `HDR_MARGIN`
   - 列N+3: `HDR_SOURCE`
3. `dictAllColDef` を `ProcessSourceSheet` に引数として渡す

### `ProcessSourceSheet`

- `colMap` を `ReDim colMap(1 To dictAllColDef.Count)` で動的確保
  - `colMap(i)` = Allシートi列目に対応するTSV列番号（0=未マップ）
- TSVヘッダーのマッピング手順：
  1. TSV各列のヘッダーを `dictHeaderMap` で正規名に変換
  2. 正規名が `dictAllColDef` に存在すれば、`dictAllColDef` の挿入順でのインデックス（1-based）を `colMap` に記録
  - `dictAllColDef` のキー一覧を配列化し、線形検索でインデックスを求める
- `outArr` を `(1 To numRows, 1 To N+3)` で確保
- 出力ループ：
  - 列1〜N: `colMap` に従いTSVデータをコピー（未マップは空文字）
  - 列N+1: 製品名（製品マスタ参照）
  - 列N+2: 口銭按分（口銭マスタ参照）
  - 列N+3: ソースファイル名

### `CollectUniqueDepts`

- `ALL_COL_DEPT` 定数の代わりに `GetAllColIndex(wsAll, HDR_DEPT)` を使用
- 部署列が `0`（未設定）の場合は空のDictionaryを返す

---

## modAggregationの変更

### `Rebuild`

- `ALL_TOTAL_COLS` の代わりに `wsAll.Cells(1, wsAll.Columns.Count).End(xlToLeft).Column` で総列数を取得
- `ALL_COL_*` 定数をすべて冒頭の動的解決に置き換え：

```vba
Dim colDept     As Integer: colDept     = GetAllColIndex(wsAll, HDR_DEPT)
Dim colDate     As Integer: colDate     = GetAllColIndex(wsAll, HDR_DATE)
Dim colClient   As Integer: colClient   = GetAllColIndex(wsAll, HDR_CLIENT)
Dim colAmount   As Integer: colAmount   = GetAllColIndex(wsAll, HDR_AMOUNT)
Dim colQty      As Integer: colQty      = GetAllColIndex(wsAll, HDR_QTY)
Dim colProdName As Integer: colProdName = GetAllColIndex(wsAll, HDR_PROD_NAME)
Dim colMargin   As Integer: colMargin   = GetAllColIndex(wsAll, HDR_MARGIN)
```

- 集計に必須の列（部署・日付・製品名・客先名）が `0` の場合はログ出力して `Exit Sub`

---

## modSetupの変更

### `SetupConfigSheet`

- I2セルに列ヘッダー追加: `"Allシート列名"`
- I列の列幅を設定
- サンプルの名寄せ8行にI列の値を追加（正規名と同じ文字列）

### `SetupAllSheet`

- 固定ヘッダー書き込み（`HDR_CLIENT`〜`HDR_SOURCE` の11列）を削除
- `LoadAllColDef()` を呼び出して動的にヘッダーを書き込む
- `AutoFit` の対象を `"A:K"` 固定から総列数ベースに変更

---

## 変更ファイル一覧

| ファイル | 変更の概要 |
|----------|-----------|
| `src/modConfig.bas` | CFG_ALL_COL_NAME追加、ALL_COL_*/ALL_TOTAL_COLS削除、LoadAllColDef/GetAllColIndex追加 |
| `src/modDataProcess.bas` | BuildAllSheet/ProcessSourceSheet/CollectUniqueDeptsを動的列対応に変更 |
| `src/modAggregation.bas` | Rebuildの列参照をGetAllColIndex動的解決に変更 |
| `src/modSetup.bas` | ConfigシートI列追加、AllシートヘッダーをLoadAllColDef経由に変更 |

---

## 非変更事項

- `HDR_*` 定数（ヘッダー文字列）は残す
- `SH_*` / `CFG_*` 定数は残す
- 集計シート（modAggregation）の集計ロジック（製品名×客先名でキーイング）は変更しない
- SharePoint連携（modSharePoint）は変更しない
- グラフ生成（modChart）は変更しない
