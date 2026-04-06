# Allシート列定義の可変化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** ConfigシートのヘッダーI列「Allシート列名」でAllシートの列構成を動的に制御できるようにする。

**Architecture:** modConfigに`LoadAllColDef()`と`GetAllColIndex()`を追加し、`ALL_COL_*`固定定数を廃止。modDataProcessはConfig定義に従って動的にAllシートを構築し、modAggregationは実行時にAllシートのヘッダー行を読んで列位置を解決する。

**Tech Stack:** Excel VBA (Scripting.Dictionary, Variant配列バルク書き込み), Shift-JIS (CP932) エンコード

---

## ファイル構成

| ファイル | 変更内容 |
|----------|---------|
| `src/modConfig.bas` | `CFG_ALL_COL_NAME`定数追加、`LoadAllColDef()`/`GetAllColIndex()`追加、後のタスクで`ALL_COL_*`/`ALL_TOTAL_COLS`削除 |
| `src/modDataProcess.bas` | `BuildAllSheet`/`ProcessSourceSheet`/`CollectUniqueDepts`を動的列対応に変更 |
| `src/modAggregation.bas` | `Rebuild`の`ALL_COL_*`参照を`GetAllColIndex`動的解決に変更 |
| `src/modSetup.bas` | `SetupConfigSheet`にI列追加、`SetupAllSheet`をLoadAllColDef経由に変更 |

---

## Task 1: modConfig — CFG_ALL_COL_NAME定数追加 + LoadAllColDef / GetAllColIndex 追加

**Files:**
- Modify: `src/modConfig.bas`

- [ ] **Step 1: CFG_ALL_COL_NAME定数をmodConfig.basに追加する**

`src/modConfig.bas` の `CFG_HEADER_COL` 定数の直後に追加：

```vba
Public Const CFG_ALL_COL_NAME   As Integer = 9   ' I: Allシート列名
```

- [ ] **Step 2: LoadAllColDef関数をmodConfig.basの末尾に追加する**

`LoadPowerAutomateUrl` 関数の直前（または末尾）に追加：

```vba
Public Function LoadAllColDef() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim r As Long
    Dim canonical As String
    Dim allColName As String

    Set dict = NewDict()
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    r = CFG_HEADER_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_HEADER_COL).Value)) <> ""
        canonical  = Trim(CStr(ws.Cells(r, CFG_HEADER_COL).Value))
        allColName = Trim(CStr(ws.Cells(r, CFG_ALL_COL_NAME).Value))
        If allColName <> "" And Not dict.Exists(canonical) Then
            dict(canonical) = allColName
        End If
        r = r + 1
    Loop

    Set LoadAllColDef = dict
End Function
```

- [ ] **Step 3: GetAllColIndex関数をmodConfig.basの末尾に追加する**

```vba
Public Function GetAllColIndex(wsAll As Worksheet, headerName As String) As Integer
    Dim lastCol As Integer
    Dim c As Integer

    lastCol = wsAll.Cells(1, wsAll.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(CStr(wsAll.Cells(1, c).Value)) = headerName Then
            GetAllColIndex = c
            Exit Function
        End If
    Next c
    GetAllColIndex = 0
End Function
```

- [ ] **Step 4: 動作確認 — この時点ではALL_COL_*定数は残したまま。コンパイルエラーがないことを確認する**

Excel VBAエディタで「デバッグ → コンパイル」を実行。エラーが出ないことを確認。

- [ ] **Step 5: コミット**

```bash
cd c:/work/autoSalesAggre
python3 -c "
import os
src_dir = 'src'
for fname in os.listdir(src_dir):
    if fname.endswith('.bas'):
        fpath = os.path.join(src_dir, fname)
        with open(fpath, 'r', encoding='utf-8') as f: text = f.read()
        with open(fpath, 'w', encoding='cp932') as f: f.write(text)
"
git add src/modConfig.bas
git commit -m "feat: modConfigにLoadAllColDef/GetAllColIndex追加"
```

---

## Task 2: modSetup — SetupConfigSheetにI列（Allシート列名）を追加

**Files:**
- Modify: `src/modSetup.bas`

- [ ] **Step 1: SetupConfigSheetのヘッダー設定部分を修正する**

`src/modSetup.bas` の `SetupConfigSheet` 内、ヘッダー名寄せ設定のブロックを以下に置き換える。

変更前：
```vba
    ' ヘッダー名寄せ (G1:H)
    ws.Cells(1, 7).Value = "ヘッダー名寄せ設定"
    ws.Cells(2, 7).Value = "正規名"
    ws.Cells(2, 8).Value = "対応列名（カンマ区切り）"
```

変更後：
```vba
    ' ヘッダー名寄せ (G1:I)
    ws.Cells(1, 7).Value = "ヘッダー名寄せ設定"
    ws.Cells(2, 7).Value = "正規名"
    ws.Cells(2, 8).Value = "対応列名（カンマ区切り）"
    ws.Cells(2, 9).Value = "Allシート列名"
```

- [ ] **Step 2: SetupConfigSheetの列幅・太字設定を修正する**

変更前：
```vba
    ws.Range("G2:H2").Font.Bold = True
    ...
    ws.Columns("G:H").ColumnWidth = 20
```

変更後：
```vba
    ws.Range("G2:I2").Font.Bold = True
    ...
    ws.Columns("G:H").ColumnWidth = 20
    ws.Columns("I").ColumnWidth = 16
```

- [ ] **Step 3: SetupConfigSheetのサンプル名寄せ8行にI列の値を追加する**

変更前（サンプルデータ部分）：
```vba
    ws.Cells(3, 7).Value = HDR_CLIENT:    ws.Cells(3, 8).Value = "得意先名,得意先コード,客先名"
    ws.Cells(4, 7).Value = HDR_PROD_CODE: ws.Cells(4, 8).Value = "品番,ProductCode"
    ws.Cells(5, 7).Value = HDR_AMOUNT:    ws.Cells(5, 8).Value = "金額,Amount,売上高"
    ws.Cells(6, 7).Value = HDR_UNIT_PRICE: ws.Cells(6, 8).Value = "単価,価格"
    ws.Cells(7, 7).Value = HDR_QTY:       ws.Cells(7, 8).Value = "数量,Qty"
    ws.Cells(8, 7).Value = HDR_DATE:      ws.Cells(8, 8).Value = "日付,発注日,Date"
    ws.Cells(9, 7).Value = HDR_SALE_TYPE: ws.Cells(9, 8).Value = "売上区分,SaleType"
    ws.Cells(10, 7).Value = HDR_DEPT:     ws.Cells(10, 8).Value = "部門,Dept"
```

変更後：
```vba
    ws.Cells(3, 7).Value = HDR_CLIENT:     ws.Cells(3, 8).Value = "得意先名,得意先コード,客先名": ws.Cells(3, 9).Value = HDR_CLIENT
    ws.Cells(4, 7).Value = HDR_PROD_CODE:  ws.Cells(4, 8).Value = "品番,ProductCode":             ws.Cells(4, 9).Value = HDR_PROD_CODE
    ws.Cells(5, 7).Value = HDR_AMOUNT:     ws.Cells(5, 8).Value = "金額,Amount,売上高":            ws.Cells(5, 9).Value = HDR_AMOUNT
    ws.Cells(6, 7).Value = HDR_UNIT_PRICE: ws.Cells(6, 8).Value = "単価,価格":                    ws.Cells(6, 9).Value = HDR_UNIT_PRICE
    ws.Cells(7, 7).Value = HDR_QTY:        ws.Cells(7, 8).Value = "数量,Qty":                     ws.Cells(7, 9).Value = HDR_QTY
    ws.Cells(8, 7).Value = HDR_DATE:       ws.Cells(8, 8).Value = "日付,発注日,Date":              ws.Cells(8, 9).Value = HDR_DATE
    ws.Cells(9, 7).Value = HDR_SALE_TYPE:  ws.Cells(9, 8).Value = "売上区分,SaleType":             ws.Cells(9, 9).Value = HDR_SALE_TYPE
    ws.Cells(10, 7).Value = HDR_DEPT:      ws.Cells(10, 8).Value = "部門,Dept":                   ws.Cells(10, 9).Value = HDR_DEPT
```

- [ ] **Step 4: コンパイル確認**

Excel VBAエディタで「デバッグ → コンパイル」を実行。エラーなしを確認。

- [ ] **Step 5: コミット**

```bash
cd c:/work/autoSalesAggre
python3 -c "
import os
src_dir = 'src'
for fname in os.listdir(src_dir):
    if fname.endswith('.bas'):
        fpath = os.path.join(src_dir, fname)
        with open(fpath, 'r', encoding='utf-8') as f: text = f.read()
        with open(fpath, 'w', encoding='cp932') as f: f.write(text)
"
git add src/modSetup.bas
git commit -m "feat: ConfigシートにAllシート列名列(I列)を追加"
```

---

## Task 3: modDataProcess — BuildAllSheet / ProcessSourceSheet を動的列対応に変更

**Files:**
- Modify: `src/modDataProcess.bas`

- [ ] **Step 1: BuildAllSheetを完全に書き換える**

`src/modDataProcess.bas` の `BuildAllSheet` 全体を以下に置き換える：

```vba
Public Sub BuildAllSheet(dictProduct As Object, dictCommission As Object, dictHeaderMap As Object)
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim allRowNum As Long
    Dim dictAllColDef As Object
    Dim i As Integer
    Dim k As Variant

    Set wsAll = ThisWorkbook.Sheets(SH_ALL)
    Set dictAllColDef = LoadAllColDef()

    ' Clear data rows, keep header
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then wsAll.Rows("2:" & lastRow).ClearContents

    ' Write header row dynamically
    i = 1
    For Each k In dictAllColDef.Keys
        wsAll.Cells(1, i).Value = dictAllColDef(k)
        i = i + 1
    Next k
    wsAll.Cells(1, i).Value = HDR_PROD_NAME
    wsAll.Cells(1, i + 1).Value = HDR_MARGIN
    wsAll.Cells(1, i + 2).Value = HDR_SOURCE

    allRowNum = 2

    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case SH_MAIN, SH_CONFIG, SH_ALL, SH_AGGR
                ' skip fixed sheets
            Case Else
                allRowNum = ProcessSourceSheet(ws, wsAll, allRowNum, dictProduct, dictCommission, dictHeaderMap, dictAllColDef)
        End Select
    Next ws
End Sub
```

- [ ] **Step 2: ProcessSourceSheetを完全に書き換える**

`src/modDataProcess.bas` の `ProcessSourceSheet` 全体を以下に置き換える：

```vba
Private Function ProcessSourceSheet(wsSrc As Worksheet, wsAll As Worksheet, _
    startRow As Long, dictProduct As Object, dictCommission As Object, _
    dictHeaderMap As Object, dictAllColDef As Object) As Long

    Dim lastSrcRow As Long
    Dim lastSrcCol As Integer
    Dim c As Integer
    Dim i As Integer
    Dim srcHeader As String
    Dim canonical As String
    Dim srcData As Variant
    Dim numRows As Long
    Dim outArr() As Variant
    Dim r As Long
    Dim prodCode As String
    Dim saleType As String
    Dim amount As Double
    Dim N As Integer
    Dim totalCols As Integer
    Dim colMap() As Integer
    Dim canonicalKeys() As String
    Dim dictCanonToIdx As Object
    Dim k As Variant
    Dim idxProdCode As Integer
    Dim idxSaleType As Integer
    Dim idxAmount As Integer

    N = dictAllColDef.Count
    totalCols = N + 3

    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    If lastSrcRow < 2 Then
        ProcessSourceSheet = startRow
        Exit Function
    End If

    lastSrcCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    ' canonicalKeys(1..N): dictAllColDef の挿入順キー配列
    ReDim canonicalKeys(1 To N)
    i = 1
    For Each k In dictAllColDef.Keys
        canonicalKeys(i) = CStr(k)
        i = i + 1
    Next k

    ' 逆引きDict: 正規名 → Allシート列インデックス(1-based)
    Set dictCanonToIdx = NewDict()
    For i = 1 To N
        dictCanonToIdx(canonicalKeys(i)) = i
    Next i

    ' colMap(i) = TSV列番号 (0=未マップ)
    ReDim colMap(1 To N)

    For c = 1 To lastSrcCol
        srcHeader = LCase(Trim(CStr(wsSrc.Cells(1, c).Value)))
        If dictHeaderMap.Exists(srcHeader) Then
            canonical = dictHeaderMap(srcHeader)
            If dictCanonToIdx.Exists(canonical) Then
                colMap(dictCanonToIdx(canonical)) = c
            End If
        End If
    Next c

    ' 計算に使う列インデックスを事前確定
    idxProdCode = 0: idxSaleType = 0: idxAmount = 0
    For i = 1 To N
        Select Case canonicalKeys(i)
            Case HDR_PROD_CODE:  idxProdCode = i
            Case HDR_SALE_TYPE:  idxSaleType = i
            Case HDR_AMOUNT:     idxAmount   = i
        End Select
    Next i

    ' Bulk read source data
    srcData = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastSrcRow, lastSrcCol)).Value
    numRows = lastSrcRow - 1
    ReDim outArr(1 To numRows, 1 To totalCols)

    For r = 1 To numRows
        ' 列1..N: TSVデータをコピー
        For i = 1 To N
            If colMap(i) > 0 Then
                outArr(r, i) = srcData(r, colMap(i))
            Else
                outArr(r, i) = ""
            End If
        Next i

        ' 列N+1: 製品名
        prodCode = ""
        If idxProdCode > 0 Then prodCode = Trim(CStr(outArr(r, idxProdCode)))
        If dictProduct.Exists(prodCode) Then
            outArr(r, N + 1) = dictProduct(prodCode)
        Else
            outArr(r, N + 1) = "[未登録]"
            If prodCode <> "" Then
                LogMessage "警告: 製品コード未登録 [" & prodCode & "] (" & wsSrc.Name & ")"
            End If
        End If

        ' 列N+2: 口銭按分
        saleType = ""
        If idxSaleType > 0 Then saleType = Trim(CStr(outArr(r, idxSaleType)))
        amount = 0
        If idxAmount > 0 Then
            If IsNumeric(outArr(r, idxAmount)) Then amount = CDbl(outArr(r, idxAmount))
        End If
        If dictCommission.Exists(saleType) Then
            outArr(r, N + 2) = amount * dictCommission(saleType) / 100
        Else
            outArr(r, N + 2) = 0
            If saleType <> "" Then
                LogMessage "警告: 売上種別未登録 [" & saleType & "] (" & wsSrc.Name & ")"
            End If
        End If

        ' 列N+3: ソースファイル名
        outArr(r, N + 3) = wsSrc.Name
    Next r

    ' Bulk write
    wsAll.Range(wsAll.Cells(startRow, 1), wsAll.Cells(startRow + numRows - 1, totalCols)).Value = outArr

    ProcessSourceSheet = startRow + numRows
End Function
```

- [ ] **Step 3: コンパイル確認**

Excel VBAエディタで「デバッグ → コンパイル」を実行。エラーなしを確認。

- [ ] **Step 4: コミット**

```bash
cd c:/work/autoSalesAggre
python3 -c "
import os
src_dir = 'src'
for fname in os.listdir(src_dir):
    if fname.endswith('.bas'):
        fpath = os.path.join(src_dir, fname)
        with open(fpath, 'r', encoding='utf-8') as f: text = f.read()
        with open(fpath, 'w', encoding='cp932') as f: f.write(text)
"
git add src/modDataProcess.bas
git commit -m "feat: BuildAllSheet/ProcessSourceSheetを動的列対応に変更"
```

---

## Task 4: modDataProcess — CollectUniqueDeptsを動的列対応に変更

**Files:**
- Modify: `src/modDataProcess.bas`

- [ ] **Step 1: CollectUniqueDeptsを書き換える**

`src/modDataProcess.bas` の `CollectUniqueDepts` 全体を以下に置き換える：

```vba
Public Function CollectUniqueDepts() As Object
    Dim dict As Object
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim deptArr As Variant
    Dim r As Long
    Dim dept As String
    Dim colDept As Integer

    Set dict = NewDict()
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)

    colDept = GetAllColIndex(wsAll, HDR_DEPT)
    If colDept = 0 Then
        Set CollectUniqueDepts = dict
        Exit Function
    End If

    lastRow = wsAll.Cells(wsAll.Rows.Count, colDept).End(xlUp).Row
    If lastRow < 2 Then
        Set CollectUniqueDepts = dict
        Exit Function
    End If

    deptArr = wsAll.Range(wsAll.Cells(2, colDept), wsAll.Cells(lastRow, colDept)).Value

    For r = 1 To UBound(deptArr, 1)
        dept = Trim(CStr(deptArr(r, 1)))
        If dept <> "" And Not dict.Exists(dept) Then dict(dept) = 1
    Next r

    Set CollectUniqueDepts = dict
End Function
```

- [ ] **Step 2: コンパイル確認**

Excel VBAエディタで「デバッグ → コンパイル」を実行。エラーなしを確認。

- [ ] **Step 3: コミット**

```bash
cd c:/work/autoSalesAggre
python3 -c "
import os
src_dir = 'src'
for fname in os.listdir(src_dir):
    if fname.endswith('.bas'):
        fpath = os.path.join(src_dir, fname)
        with open(fpath, 'r', encoding='utf-8') as f: text = f.read()
        with open(fpath, 'w', encoding='cp932') as f: f.write(text)
"
git add src/modDataProcess.bas
git commit -m "feat: CollectUniqueDeptsをGetAllColIndex動的解決に変更"
```

---

## Task 5: modAggregation — RebuildをALL_COL_*廃止・動的解決に変更

**Files:**
- Modify: `src/modAggregation.bas`

- [ ] **Step 1: Rebuildの先頭変数宣言と列解決処理を変更する**

`src/modAggregation.bas` の `Rebuild` 内、変数宣言ブロックと列定数参照部分を変更する。

変更前（変数宣言の一部）：
```vba
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim allData As Variant
```

変更後（変数宣言ブロック末尾に追加）：
```vba
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim allData As Variant
    Dim totalCols As Integer
    Dim colDept     As Integer
    Dim colDate     As Integer
    Dim colClient   As Integer
    Dim colAmount   As Integer
    Dim colQty      As Integer
    Dim colProdName As Integer
    Dim colMargin   As Integer
```

- [ ] **Step 2: allDataのロード前に列インデックスを動的解決するコードを追加する**

変更前：
```vba
    ' Load all sheet data
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then Exit Sub
    ClearAggrTable wsAggr

    allData = wsAll.Range(wsAll.Cells(2, 1), wsAll.Cells(lastRow, ALL_TOTAL_COLS)).Value
```

変更後：
```vba
    ' Load all sheet data
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then Exit Sub
    ClearAggrTable wsAggr

    ' 列インデックスを動的解決
    colDept     = GetAllColIndex(wsAll, HDR_DEPT)
    colDate     = GetAllColIndex(wsAll, HDR_DATE)
    colClient   = GetAllColIndex(wsAll, HDR_CLIENT)
    colAmount   = GetAllColIndex(wsAll, HDR_AMOUNT)
    colQty      = GetAllColIndex(wsAll, HDR_QTY)
    colProdName = GetAllColIndex(wsAll, HDR_PROD_NAME)
    colMargin   = GetAllColIndex(wsAll, HDR_MARGIN)

    If colProdName = 0 Or colClient = 0 Then
        LogMessage "[エラー] 集計に必要な列（製品名・客先名）がallシートに見つかりません"
        Exit Sub
    End If

    totalCols = wsAll.Cells(1, wsAll.Columns.Count).End(xlToLeft).Column
    allData = wsAll.Range(wsAll.Cells(2, 1), wsAll.Cells(lastRow, totalCols)).Value
```

- [ ] **Step 3: Rebuildのフィルタ・集計ループ内のALL_COL_*を動的変数に置き換える**

変更前：
```vba
        ' Dept filter
        If selectedDept <> "全部署" And selectedDept <> "" Then
            If Trim(CStr(allData(r, ALL_COL_DEPT))) <> selectedDept Then GoTo NextRow
        End If

        ' Date filter
        If useFrom Or useTo Then
            saleDateRaw = allData(r, ALL_COL_DATE)
            ...
        End If

        ' Accumulate totals
        pName = Trim(CStr(allData(r, ALL_COL_PROD_NAME)))
        cName = Trim(CStr(allData(r, ALL_COL_CLIENT)))
        ...
        If IsNumeric(allData(r, ALL_COL_AMOUNT)) Then amt = CDbl(allData(r, ALL_COL_AMOUNT))
        If IsNumeric(allData(r, ALL_COL_QTY)) Then qty = CDbl(allData(r, ALL_COL_QTY))
        If IsNumeric(allData(r, ALL_COL_MARGIN)) Then margin = CDbl(allData(r, ALL_COL_MARGIN))
```

変更後：
```vba
        ' Dept filter
        If selectedDept <> "全部署" And selectedDept <> "" Then
            If colDept > 0 Then
                If Trim(CStr(allData(r, colDept))) <> selectedDept Then GoTo NextRow
            End If
        End If

        ' Date filter
        If useFrom Or useTo Then
            If colDate = 0 Then GoTo NextRow
            saleDateRaw = allData(r, colDate)
            If Not IsDate(saleDateRaw) Then GoTo NextRow
            saleDate = CDate(saleDateRaw)
            If useFrom And saleDate < fromDate Then GoTo NextRow
            If useTo And saleDate > toDate Then GoTo NextRow
        End If

        ' Accumulate totals
        pName = Trim(CStr(allData(r, colProdName)))
        cName = ""
        If colClient > 0 Then cName = Trim(CStr(allData(r, colClient)))
        ...
        If colAmount > 0 Then
            If IsNumeric(allData(r, colAmount)) Then amt = CDbl(allData(r, colAmount))
        End If
        If colQty > 0 Then
            If IsNumeric(allData(r, colQty)) Then qty = CDbl(allData(r, colQty))
        End If
        If colMargin > 0 Then
            If IsNumeric(allData(r, colMargin)) Then margin = CDbl(allData(r, colMargin))
        End If
```

- [ ] **Step 4: コンパイル確認**

Excel VBAエディタで「デバッグ → コンパイル」を実行。エラーなしを確認。

- [ ] **Step 5: コミット**

```bash
cd c:/work/autoSalesAggre
python3 -c "
import os
src_dir = 'src'
for fname in os.listdir(src_dir):
    if fname.endswith('.bas'):
        fpath = os.path.join(src_dir, fname)
        with open(fpath, 'r', encoding='utf-8') as f: text = f.read()
        with open(fpath, 'w', encoding='cp932') as f: f.write(text)
"
git add src/modAggregation.bas
git commit -m "feat: RebuildをGetAllColIndex動的解決に変更"
```

---

## Task 6: modSetup — SetupAllSheetを動的ヘッダーに変更

**Files:**
- Modify: `src/modSetup.bas`

- [ ] **Step 1: SetupAllSheetを書き換える**

`src/modSetup.bas` の `SetupAllSheet` 全体を以下に置き換える：

```vba
Private Sub SetupAllSheet()
    Dim ws As Worksheet
    Dim dictAllColDef As Object
    Dim i As Integer
    Dim k As Variant
    Dim totalCols As Integer

    Set ws = ThisWorkbook.Sheets(SH_ALL)
    Set dictAllColDef = LoadAllColDef()

    ' 動的ヘッダー書き込み
    i = 1
    For Each k In dictAllColDef.Keys
        ws.Cells(1, i).Value = dictAllColDef(k)
        i = i + 1
    Next k
    ws.Cells(1, i).Value = HDR_PROD_NAME
    ws.Cells(1, i + 1).Value = HDR_MARGIN
    ws.Cells(1, i + 2).Value = HDR_SOURCE

    totalCols = i + 2

    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 240)
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(1, totalCols)).EntireColumn.AutoFit

    ' SharePointアップロードボタン
    Dim uploadBtn As Object
    Set uploadBtn = ws.Buttons.Add(700, 5, 180, 28)
    uploadBtn.Caption = "SharePointへアップロード"
    uploadBtn.OnAction = "modSharePoint.UploadAllToSharePoint"
End Sub
```

- [ ] **Step 2: コンパイル確認**

Excel VBAエディタで「デバッグ → コンパイル」を実行。エラーなしを確認。

- [ ] **Step 3: コミット**

```bash
cd c:/work/autoSalesAggre
python3 -c "
import os
src_dir = 'src'
for fname in os.listdir(src_dir):
    if fname.endswith('.bas'):
        fpath = os.path.join(src_dir, fname)
        with open(fpath, 'r', encoding='utf-8') as f: text = f.read()
        with open(fpath, 'w', encoding='cp932') as f: f.write(text)
"
git add src/modSetup.bas
git commit -m "feat: SetupAllSheetをLoadAllColDef動的ヘッダーに変更"
```

---

## Task 7: modConfig — ALL_COL_* / ALL_TOTAL_COLS 定数を削除

**Files:**
- Modify: `src/modConfig.bas`

> このタスクはTask 3〜6が完了してから実施すること。全ての参照が削除済みであることを確認してから定数を削除する。

- [ ] **Step 1: modConfig.basから削除する定数ブロックを確認する**

削除対象（`src/modConfig.bas` の `' ===== all sheet column indices =====` セクション全体）：

```vba
' ===== all sheet column indices (1-based) =====
Public Const ALL_COL_CLIENT     As Integer = 1
Public Const ALL_COL_PROD_CODE  As Integer = 2
Public Const ALL_COL_AMOUNT     As Integer = 3
Public Const ALL_COL_UNIT_PRICE As Integer = 4
Public Const ALL_COL_QTY        As Integer = 5
Public Const ALL_COL_DATE       As Integer = 6
Public Const ALL_COL_SALE_TYPE  As Integer = 7
Public Const ALL_COL_DEPT       As Integer = 8
Public Const ALL_COL_PROD_NAME  As Integer = 9
Public Const ALL_COL_MARGIN     As Integer = 10
Public Const ALL_COL_SOURCE     As Integer = 11
Public Const ALL_TOTAL_COLS     As Integer = 11
```

- [ ] **Step 2: 上記ブロックをmodConfig.basから削除する**

該当セクション（セクションコメント行を含む12行）を削除する。

- [ ] **Step 3: 残存参照がないことをGrepで確認する**

```bash
cd c:/work/autoSalesAggre
grep -r "ALL_COL_\|ALL_TOTAL_COLS" src/
```

期待結果: 出力なし（0件）

- [ ] **Step 4: コンパイル確認**

Excel VBAエディタで「デバッグ → コンパイル」を実行。エラーなしを確認。

- [ ] **Step 5: コミット**

```bash
cd c:/work/autoSalesAggre
python3 -c "
import os
src_dir = 'src'
for fname in os.listdir(src_dir):
    if fname.endswith('.bas'):
        fpath = os.path.join(src_dir, fname)
        with open(fpath, 'r', encoding='utf-8') as f: text = f.read()
        with open(fpath, 'w', encoding='cp932') as f: f.write(text)
"
git add src/modConfig.bas
git commit -m "feat: ALL_COL_*/ALL_TOTAL_COLS定数を廃止"
```

---

## Task 8: ワークブック再構築と動作確認

- [ ] **Step 1: 既存のワークブックを削除してセットアップを再実行する**

```bash
cd c:/work/autoSalesAggre
rm -f autoSalesAggre.xlsm
cscript setup/create_workbook.vbs
```

- [ ] **Step 2: ConfigシートのI列を確認する**

`autoSalesAggre.xlsm` を開いてConfigシートを確認：
- I2セルに「Allシート列名」と表示されている
- G3:I10の8行にサンプルデータが入っている（I列は正規名と同じ値）

- [ ] **Step 3: Allシートのヘッダーを確認する**

Allシートを確認：
- A1〜H1: 客先名〜部署（名寄せ順）
- I1: 製品名
- J1: 口銭按分
- K1: ソースファイル名

- [ ] **Step 4: TSVを読み込んでAllシートへのデータ出力を確認する**

mainシートの「ファイルを読み込む」ボタンでサンプルTSVを読み込み：
- Allシートにデータが11列で出力されること
- 製品名・口銭按分が末尾2列に入ること

- [ ] **Step 5: 集計シートで絞り込みが動作することを確認する**

集計シートで部署・期間フィルタを変更し、集計テーブルが正しく更新されること。

- [ ] **Step 6: Config I列を変更して列の可変動作を確認する**

ConfigシートのI5セル（売上金額行のAllシート列名）を空白にしてからRunAllを再実行：
- Allシートの列から「売上金額」が除外されること
- 製品名・口銭按分・ソースファイル名は末尾に表示され続けること

空白にしたI5を元の値「売上金額」に戻してRunAllを再実行：
- 売上金額列が復活すること

- [ ] **Step 7: 最終コミット**

```bash
cd c:/work/autoSalesAggre
git add -A
git commit -m "test: ワークブック再構築・動作確認完了"
```
