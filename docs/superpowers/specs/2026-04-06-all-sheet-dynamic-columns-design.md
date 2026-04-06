# Allシート列定義の可変化 — 設計書

**日付:** 2026-04-06（設計改善レビュー反映: 2026-04-07）
**対象ファイル:** modConfig.bas, modDataProcess.bas, modAggregation.bas, modSetup.bas, modSharePoint.bas, modChart.bas

---

## 概要

Allシートの列構成（列数・列順・列名）を、ConfigシートのI列「Allシート列名」で制御できるようにした。
VBAコードを修正せずに、Configシートの変更だけで列の追加・削除・並べ替えが可能。

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

| 列 | 内容 |
|----|------|
| G | 正規名（変更なし） |
| H | 対応列名（カンマ区切りエイリアス）（変更なし） |
| I | **Allシート列名**（新規追加） |

- I列に値があれば → Allシートにその列名で出力
- I列が空白 → Allシートに出力しない
- **行の並び順 = Allシートの列順**

### Allシート最終列構造

```
[I列で定義したN列（行順）] + 製品名 + 口銭按分 + ソースファイル名
```

---

## modConfigの設計

### 定数

| 定数 | 値 | 役割 |
|------|----|------|
| `CFG_ALL_COL_NAME` | 9 | Config I列（Allシート列名） |
| `AGGR_INDENT` | `"　　"` | 集計表の子行インデント（全角スペース2文字）。**modAggregation と modChart の両方が参照するため定数化必須** |
| `AGGR_KEY_SEP` | `"||"` | dictSummaryキーのセパレータ（`製品名 & AGGR_KEY_SEP & 客先名`） |

### 関数: `LoadAllColDef() As Object`

- 名寄せテーブル（CFG_HEADER_HDR_ROW+1行目から）を1パスで読む
- I列（CFG_ALL_COL_NAME）が空白でない行のみをDictionaryに格納
- 戻り値: `key = 正規名（G列）`, `value = Allシート列名（I列）`、挿入順 = Allシートの列順
- **同じ名寄せテーブルを `LoadHeaderMap` も読んでいる**（2関数とも同じループ構造）

### 関数: `GetAllColIndex(wsAll As Worksheet, headerName As String) As Integer`

- Allシートの1行目を全列スキャン
- `headerName` に一致する列番号（1-based）を返す。見つからなければ `0`
- `Rebuild`・`UploadAllToSharePoint` が呼び出しの冒頭で使用する

---

## modDataProcessの設計

### `BuildAllSheet`

1. `LoadAllColDef()` で列定義を取得
2. ヘッダー行を動的に書き込み（dictAllColDefキーループ → HDR_PROD_NAME → HDR_MARGIN → HDR_SOURCE）
3. `ProcessSourceSheet` に `dictAllColDef` を渡す

### `ProcessSourceSheet`（設計改善後）

列マッピングの手順：

1. `canonicalKeys(1..N)` 配列を `dictAllColDef.Keys` 挿入順で構築
2. TSV各列のヘッダーを `dictHeaderMap` で正規名に変換
3. 正規名を `canonicalKeys` の**線形検索**でAllシート列インデックスに変換し `colMap(i)` に記録

> **設計ポイント:** 逆引きDictionary（dictCanonToIdx）は使わない。列数は高々20程度なので線形検索で十分シンプル。

計算列の検索（idxProdCode / idxSaleType / idxAmount）も `canonicalKeys` を1ループで走査。

### `CollectUniqueDepts`

`GetAllColIndex(wsAll, HDR_DEPT)` で部署列を動的解決。列が存在しない場合は空Dictionaryを返す。

---

## modAggregationの設計

### `Rebuild`

冒頭で `GetAllColIndex` を7回呼び出して列インデックスを動的解決：

```vba
colDept     = GetAllColIndex(wsAll, HDR_DEPT)
colDate     = GetAllColIndex(wsAll, HDR_DATE)
colClient   = GetAllColIndex(wsAll, HDR_CLIENT)
colAmount   = GetAllColIndex(wsAll, HDR_AMOUNT)
colQty      = GetAllColIndex(wsAll, HDR_QTY)
colProdName = GetAllColIndex(wsAll, HDR_PROD_NAME)
colMargin   = GetAllColIndex(wsAll, HDR_MARGIN)
```

`colProdName = 0 Or colClient = 0` の場合はログ出力して終了。

### `DrawAggrTable`

- 製品グループの親行: `wsAggr.Cells(currentRow, 1).Value = pName`（太字・グレー背景）
- 客先子行: `wsAggr.Cells(currentRow, 1).Value = AGGR_INDENT & cName`（`AGGR_INDENT` 定数使用）
- dictSummaryキー: `pName & AGGR_KEY_SEP & cName`（`AGGR_KEY_SEP` 定数使用）

> **設計ポイント:** `AGGR_INDENT` は `modChart` の子行判定（`Left(cellVal, Len(AGGR_INDENT)) <> AGGR_INDENT`）と対になっている。変更は `modConfig` の定数1か所で完結する。

### `ClearAggrTable` の呼び出し順序

```vba
If lastRow < 2 Then Exit Sub   ' allが空なら集計表を保持してそのまま終了
ClearAggrTable wsAggr          ' データがある場合のみクリアして再描画
```

---

## modSetupの設計

### `SetupConfigSheet`

- I2セル: `"Allシート列名"`
- I列の列幅: 16
- サンプル名寄せ8行のI列に正規名と同じ値を設定（デフォルトで全列表示）

### `SetupAllSheet`

- `LoadAllColDef()` を呼び出して動的にヘッダーを書き込む
- AutoFitの対象を `totalCols`（動的）で指定

### `InjectAggrEvent`

> **【注意】** 注入するVBAコード文字列の中で `AGGR_DEPT_CELL`, `AGGR_FROM_CELL`, `AGGR_TO_CELL` 定数を直接参照している。これらの定数名を変更した場合は文字列の中も合わせて変更すること（静的検索では発見できない）。

---

## modSharePointの設計

### 共通関数: `SendJson(paUrl, jsonBody, successMsg)`

HTTP POST送信・ステータス判定・エラーログを一か所に集約。

```vba
Private Sub SendJson(paUrl As String, jsonBody As String, successMsg As String)
```

`UploadToSharePoint` と `UploadAllToSharePoint` の両方がこれを呼び出す。
SharePointの仕様変更（認証追加・タイムアウト変更等）はこの1関数のみ修正すればよい。

### 列インデックスの解決

`UploadAllToSharePoint` は冒頭で `GetAllColIndex` を11回呼び出して全列インデックスを取得してから `allData` を読み込む。

---

## modChartの設計

### `DrawAggrChart`

集計表の親行/子行を `AGGR_INDENT` で判別：

```vba
If Left(cellVal, Len(AGGR_INDENT)) <> AGGR_INDENT And cellVal <> "合計" And Trim(cellVal) <> "" Then
    ' 親行（製品グループ）
```

> **設計ポイント:** `AGGR_INDENT` は `modConfig` の定数を直接参照。`DrawAggrTable` が書き込む字下げと同じ定数を使っているため、変更は1か所で完結する。

---

## 既知の設計課題（将来改善候補）

| 課題 | 内容 | 影響範囲 |
|------|------|---------|
| LoadHeaderMap と LoadAllColDef の重複ループ | 同じ名寄せテーブルを2関数が別々に読んでいる。名寄せテーブルに列を追加する場合は両方を修正する必要がある | modConfig |
| GetAllColIndex の多重呼び出し | RebuildとUploadAllToSharePointで合計18回の行スキャンが発生する。将来列が大幅に増えた場合は `BuildColIndexMap(wsAll) As Object` に集約することを検討 | modAggregation, modSharePoint |
