# 売上集計自動化ツール (autoSalesAggre) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an Excel VBA macro workbook that aggregates TSV sales data files with header normalization, master lookups, and dynamic department/date-filtered hierarchical reporting.

**Architecture:** Five standard VBA modules handle configuration constants + master loading (`modConfig`), TSV file I/O (`modFileIO`), all-sheet data processing (`modDataProcess`), aggregation and rendering (`modAggregation`), and UI orchestration + logging (`modUIControl`). A one-time initialization module (`modSetup`) sets up sheet layouts and injects the `Worksheet_Change` event into the 集計 sheet. Source code lives in `.bas` text files; a VBScript setup script generates `autoSalesAggre.xlsm`.

**Tech Stack:** Excel VBA 7.x, Microsoft Scripting Runtime (`Scripting.Dictionary`), VBScript (workbook generation)

---

## File Map

| File | Purpose |
|------|---------|
| `src/modConfig.bas` | All constants (column indices, sheet names, cell addresses) + master-loading functions + `RefreshDeptList` |
| `src/modFileIO.bas` | File selection dialog + TSV-to-sheet loading |
| `src/modDataProcess.bas` | `BuildAllSheet` (header map → all sheet) + `CollectUniqueDepts` |
| `src/modAggregation.bas` | `Rebuild` (filter + aggregate) + hierarchical table rendering |
| `src/modUIControl.bas` | `RunAll` orchestrator + `LogMessage` |
| `src/modSetup.bas` | One-time `InitWorkbook`: sheet naming/layout + event injection |
| `setup/create_workbook.vbs` | Creates `.xlsm`, imports modules, calls `InitWorkbook`, saves |
| `CLAUDE.md` | Project setup and architecture notes |

---

### Task 1: Project structure and .gitignore

**Files:**
- Create: `.gitignore`

- [ ] **Step 1: Create .gitignore**

```
# Generated workbook (rebuild from src/ with setup/create_workbook.vbs)
*.xlsm
*.xlsb
*.xls

# Excel temp files
~$*

# OS
.DS_Store
Thumbs.db
```

- [ ] **Step 2: Commit**

```bash
git add .gitignore
git commit -m "chore: add .gitignore for VBA project"
```

---

### Task 2: modConfig.bas — Constants and master loading

**Files:**
- Create: `src/modConfig.bas`

- [ ] **Step 1: Create src/modConfig.bas**

```vba
Attribute VB_Name = "modConfig"
Option Explicit

' ===== Config sheet table positions =====
Public Const CFG_PRODUCT_HDR_ROW    As Integer = 2   ' 製品マスタ header row (A2)
Public Const CFG_PRODUCT_COL        As Integer = 1   ' A: 製品コード
Public Const CFG_COMMISSION_HDR_ROW As Integer = 2   ' 口銭マスタ header row (D2)
Public Const CFG_COMMISSION_COL     As Integer = 4   ' D: 売上種別
Public Const CFG_HEADER_HDR_ROW     As Integer = 2   ' 名寄せ header row (G2)
Public Const CFG_HEADER_COL         As Integer = 7   ' G: 正規名
Public Const CFG_DEPT_HDR_ROW       As Integer = 2   ' 部署リスト header row (J2)
Public Const CFG_DEPT_COL           As Integer = 10  ' J: 部署リスト

' ===== all sheet column indices (1-based) =====
Public Const ALL_COL_CLIENT     As Integer = 1   ' 客先名
Public Const ALL_COL_PROD_CODE  As Integer = 2   ' 製品コード
Public Const ALL_COL_AMOUNT     As Integer = 3   ' 売上金額
Public Const ALL_COL_UNIT_PRICE As Integer = 4   ' 製品単価
Public Const ALL_COL_QTY        As Integer = 5   ' 売上数量
Public Const ALL_COL_DATE       As Integer = 6   ' 売上発生日
Public Const ALL_COL_SALE_TYPE  As Integer = 7   ' 売上種別
Public Const ALL_COL_DEPT       As Integer = 8   ' 部署
Public Const ALL_COL_PROD_NAME  As Integer = 9   ' 製品名 (calculated)
Public Const ALL_COL_MARGIN     As Integer = 10  ' 部署取り分 (calculated)
Public Const ALL_COL_SOURCE     As Integer = 11  ' ソースファイル
Public Const ALL_TOTAL_COLS     As Integer = 11

' ===== Sheet names =====
Public Const SH_MAIN   As String = "main"
Public Const SH_CONFIG As String = "Config"
Public Const SH_ALL    As String = "all"
Public Const SH_AGGR   As String = "集計"

' ===== 集計 sheet cell addresses =====
Public Const AGGR_DEPT_CELL As String = "B1"
Public Const AGGR_FROM_CELL As String = "B2"
Public Const AGGR_TO_CELL   As String = "B3"
Public Const AGGR_HDR_ROW   As Integer = 5
Public Const AGGR_DATA_ROW  As Integer = 6

' ===== main sheet =====
Public Const MAIN_LOG_START_ROW As Integer = 3

' ---------- Master loading ----------

Public Function LoadProductDict() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    Dim r As Long
    r = CFG_PRODUCT_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_PRODUCT_COL).Value)) <> ""
        Dim code As String
        code = Trim(CStr(ws.Cells(r, CFG_PRODUCT_COL).Value))
        If Not dict.Exists(code) Then
            dict(code) = Trim(CStr(ws.Cells(r, CFG_PRODUCT_COL + 1).Value))
        End If
        r = r + 1
    Loop

    Set LoadProductDict = dict
End Function

Public Function LoadCommissionDict() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    Dim r As Long
    r = CFG_COMMISSION_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_COMMISSION_COL).Value)) <> ""
        Dim stype As String
        stype = Trim(CStr(ws.Cells(r, CFG_COMMISSION_COL).Value))
        If Not dict.Exists(stype) Then
            dict(stype) = CDbl(ws.Cells(r, CFG_COMMISSION_COL + 1).Value)
        End If
        r = r + 1
    Loop

    Set LoadCommissionDict = dict
End Function

Public Function LoadHeaderMap() As Object
    ' Returns dict: LCase(trimmed_alias) -> canonical_column_name
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    Dim r As Long
    r = CFG_HEADER_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_HEADER_COL).Value)) <> ""
        Dim canonical As String
        canonical = Trim(CStr(ws.Cells(r, CFG_HEADER_COL).Value))
        Dim aliases As String
        aliases = Trim(CStr(ws.Cells(r, CFG_HEADER_COL + 1).Value))

        ' Register canonical name itself
        If Not dict.Exists(LCase(canonical)) Then dict(LCase(canonical)) = canonical

        ' Register each alias
        Dim parts() As String
        parts = Split(aliases, ",")
        Dim i As Integer
        For i = 0 To UBound(parts)
            Dim a As String
            a = LCase(Trim(parts(i)))
            If a <> "" And Not dict.Exists(a) Then dict(a) = canonical
        Next i
        r = r + 1
    Loop

    Set LoadHeaderMap = dict
End Function

Public Sub RefreshDeptList(dictDept As Object)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    ' Clear J3 downward
    Dim clearRow As Long
    clearRow = CFG_DEPT_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(clearRow, CFG_DEPT_COL).Value)) <> ""
        ws.Cells(clearRow, CFG_DEPT_COL).ClearContents
        clearRow = clearRow + 1
    Loop

    ' J2 = "全部署" (fixed)
    ws.Cells(CFG_DEPT_HDR_ROW, CFG_DEPT_COL).Value = "全部署"

    ' Write unique depts from J3
    Dim r As Long
    r = CFG_DEPT_HDR_ROW + 1
    Dim key As Variant
    For Each key In dictDept.Keys
        ws.Cells(r, CFG_DEPT_COL).Value = key
        r = r + 1
    Next key

    Dim lastDeptRow As Long
    lastDeptRow = r - 1

    ' Update 集計!B1 dropdown
    Dim wsAggr As Worksheet
    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)
    With wsAggr.Range(AGGR_DEPT_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & SH_CONFIG & "!$J$" & CFG_DEPT_HDR_ROW & ":$J$" & lastDeptRow
    End With

    If Trim(CStr(wsAggr.Range(AGGR_DEPT_CELL).Value)) = "" Then
        wsAggr.Range(AGGR_DEPT_CELL).Value = "全部署"
    End If
End Sub
```

- [ ] **Step 2: Commit**

```bash
git add src/modConfig.bas
git commit -m "feat: add modConfig with constants and master-loading functions"
```

---

### Task 3: modFileIO.bas — TSV file selection and loading

**Files:**
- Create: `src/modFileIO.bas`

- [ ] **Step 1: Create src/modFileIO.bas**

```vba
Attribute VB_Name = "modFileIO"
Option Explicit

Public Function SelectFiles() As Variant
    ' Returns Variant array of paths, or Boolean False if cancelled
    SelectFiles = Application.GetOpenFilename( _
        FileFilter:="テキストファイル (*.txt;*.tsv),*.txt;*.tsv", _
        Title:="読み込むTSVファイルを選択してください", _
        MultiSelect:=True)
End Function

Public Function LoadTsvToSheet(filePath As String) As Boolean
    ' Creates or replaces a sheet named after the file. Returns True on success.
    On Error GoTo ErrHandler

    Dim sheetName As String
    sheetName = FilePathToSheetName(filePath)

    ' Delete existing sheet with same name
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws

    ' Insert before 集計 sheet to keep sheet order consistent
    Dim newSheet As Worksheet
    Set newSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(SH_AGGR))
    newSheet.Name = sheetName

    ' Read TSV line by line, store all values as text to preserve leading zeros
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    Dim rowNum As Long
    rowNum = 1
    Do While Not EOF(fileNum)
        Dim lineText As String
        Line Input #fileNum, lineText

        Dim cols() As String
        cols = Split(lineText, vbTab)

        Dim c As Integer
        For c = 0 To UBound(cols)
            With newSheet.Cells(rowNum, c + 1)
                .NumberFormat = "@"
                .Value = cols(c)
            End With
        Next c
        rowNum = rowNum + 1
    Loop

    Close #fileNum
    LoadTsvToSheet = True
    Exit Function

ErrHandler:
    If fileNum > 0 Then Close #fileNum
    LoadTsvToSheet = False
End Function

Private Function FilePathToSheetName(filePath As String) As String
    ' Extract filename without extension; strip invalid sheet name chars; truncate to 31 chars
    Dim fileName As String
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    Dim dotPos As Integer
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then fileName = Left(fileName, dotPos - 1)

    ' Remove characters invalid for sheet names: \ / ? * [ ] :
    Dim invalids As String
    invalids = "\/?*[]:"
    Dim i As Integer
    For i = 1 To Len(invalids)
        fileName = Join(Split(fileName, Mid(invalids, i, 1)), "_")
    Next i

    If Len(fileName) > 31 Then fileName = Left(fileName, 31)
    FilePathToSheetName = fileName
End Function
```

- [ ] **Step 2: Commit**

```bash
git add src/modFileIO.bas
git commit -m "feat: add modFileIO for TSV file selection and sheet loading"
```

---

### Task 4: modDataProcess.bas — all sheet construction

**Files:**
- Create: `src/modDataProcess.bas`

- [ ] **Step 1: Create src/modDataProcess.bas**

```vba
Attribute VB_Name = "modDataProcess"
Option Explicit

Public Sub BuildAllSheet(dictProduct As Object, dictCommission As Object, dictHeaderMap As Object)
    Dim wsAll As Worksheet
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)

    ' Clear data rows, keep header
    Dim lastRow As Long
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then wsAll.Rows("2:" & lastRow).ClearContents

    ' Write header row
    wsAll.Cells(1, ALL_COL_CLIENT).Value = "客先名"
    wsAll.Cells(1, ALL_COL_PROD_CODE).Value = "製品コード"
    wsAll.Cells(1, ALL_COL_AMOUNT).Value = "売上金額"
    wsAll.Cells(1, ALL_COL_UNIT_PRICE).Value = "製品単価"
    wsAll.Cells(1, ALL_COL_QTY).Value = "売上数量"
    wsAll.Cells(1, ALL_COL_DATE).Value = "売上発生日"
    wsAll.Cells(1, ALL_COL_SALE_TYPE).Value = "売上種別"
    wsAll.Cells(1, ALL_COL_DEPT).Value = "部署"
    wsAll.Cells(1, ALL_COL_PROD_NAME).Value = "製品名"
    wsAll.Cells(1, ALL_COL_MARGIN).Value = "部署取り分"
    wsAll.Cells(1, ALL_COL_SOURCE).Value = "ソースファイル"

    Dim allRowNum As Long
    allRowNum = 2

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case SH_MAIN, SH_CONFIG, SH_ALL, SH_AGGR
                ' skip fixed sheets
            Case Else
                allRowNum = ProcessSourceSheet(ws, wsAll, allRowNum, dictProduct, dictCommission, dictHeaderMap)
        End Select
    Next ws
End Sub

Private Function ProcessSourceSheet(wsSrc As Worksheet, wsAll As Worksheet, _
    startRow As Long, dictProduct As Object, dictCommission As Object, _
    dictHeaderMap As Object) As Long

    Dim lastSrcRow As Long
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    If lastSrcRow < 2 Then
        ProcessSourceSheet = startRow
        Exit Function
    End If

    Dim lastSrcCol As Integer
    lastSrcCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    ' colMap(allColIndex - 1) = source column number (0 = unmapped)
    Dim colMap(10) As Integer
    Dim c As Integer
    For c = 1 To lastSrcCol
        Dim srcHeader As String
        srcHeader = LCase(Trim(CStr(wsSrc.Cells(1, c).Value)))
        If dictHeaderMap.Exists(srcHeader) Then
            Select Case dictHeaderMap(srcHeader)
                Case "客先名":      colMap(ALL_COL_CLIENT - 1) = c
                Case "製品コード":  colMap(ALL_COL_PROD_CODE - 1) = c
                Case "売上金額":    colMap(ALL_COL_AMOUNT - 1) = c
                Case "製品単価":    colMap(ALL_COL_UNIT_PRICE - 1) = c
                Case "売上数量":    colMap(ALL_COL_QTY - 1) = c
                Case "売上発生日":  colMap(ALL_COL_DATE - 1) = c
                Case "売上種別":    colMap(ALL_COL_SALE_TYPE - 1) = c
                Case "部署":        colMap(ALL_COL_DEPT - 1) = c
            End Select
        End If
    Next c

    ' Bulk read source data into Variant array
    Dim srcData As Variant
    srcData = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastSrcRow, lastSrcCol)).Value

    Dim numRows As Long
    numRows = lastSrcRow - 1
    Dim outArr() As Variant
    ReDim outArr(1 To numRows, 1 To ALL_TOTAL_COLS)

    Dim r As Long
    For r = 1 To numRows
        ' Copy source columns ALL_COL_CLIENT to ALL_COL_DEPT (cols 1-8)
        Dim allCol As Integer
        For allCol = ALL_COL_CLIENT To ALL_COL_DEPT
            If colMap(allCol - 1) > 0 Then
                outArr(r, allCol) = srcData(r, colMap(allCol - 1))
            Else
                outArr(r, allCol) = ""
            End If
        Next allCol

        ' Calculate 製品名 (col 9)
        Dim prodCode As String
        prodCode = Trim(CStr(outArr(r, ALL_COL_PROD_CODE)))
        If dictProduct.Exists(prodCode) Then
            outArr(r, ALL_COL_PROD_NAME) = dictProduct(prodCode)
        Else
            outArr(r, ALL_COL_PROD_NAME) = "[未登録]"
            If prodCode <> "" Then
                LogMessage "警告: 製品コード未登録 [" & prodCode & "] (" & wsSrc.Name & ")"
            End If
        End If

        ' Calculate 部署取り分 (col 10)
        Dim saleType As String
        saleType = Trim(CStr(outArr(r, ALL_COL_SALE_TYPE)))
        Dim amount As Double
        If IsNumeric(outArr(r, ALL_COL_AMOUNT)) Then amount = CDbl(outArr(r, ALL_COL_AMOUNT))
        If dictCommission.Exists(saleType) Then
            outArr(r, ALL_COL_MARGIN) = amount * dictCommission(saleType) / 100
        Else
            outArr(r, ALL_COL_MARGIN) = 0
            If saleType <> "" Then
                LogMessage "警告: 売上種別未登録 [" & saleType & "] (" & wsSrc.Name & ")"
            End If
        End If

        ' Source file name (col 11)
        outArr(r, ALL_COL_SOURCE) = wsSrc.Name
    Next r

    ' Bulk write to all sheet
    wsAll.Range(wsAll.Cells(startRow, 1), wsAll.Cells(startRow + numRows - 1, ALL_TOTAL_COLS)).Value = outArr

    ProcessSourceSheet = startRow + numRows
End Function

Public Function CollectUniqueDepts() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim wsAll As Worksheet
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)

    Dim lastRow As Long
    lastRow = wsAll.Cells(wsAll.Rows.Count, ALL_COL_DEPT).End(xlUp).Row
    If lastRow < 2 Then
        Set CollectUniqueDepts = dict
        Exit Function
    End If

    Dim deptArr As Variant
    deptArr = wsAll.Range(wsAll.Cells(2, ALL_COL_DEPT), wsAll.Cells(lastRow, ALL_COL_DEPT)).Value

    Dim r As Long
    For r = 1 To UBound(deptArr, 1)
        Dim dept As String
        dept = Trim(CStr(deptArr(r, 1)))
        If dept <> "" And Not dict.Exists(dept) Then dict(dept) = 1
    Next r

    Set CollectUniqueDepts = dict
End Function
```

- [ ] **Step 2: Commit**

```bash
git add src/modDataProcess.bas
git commit -m "feat: add modDataProcess for all sheet construction and dept collection"
```

---

### Task 5: modAggregation.bas — Aggregation logic and table rendering

**Files:**
- Create: `src/modAggregation.bas`

- [ ] **Step 1: Create src/modAggregation.bas**

```vba
Attribute VB_Name = "modAggregation"
Option Explicit

Public Sub Rebuild()
    Dim wsAggr As Worksheet
    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)

    ' Read filter conditions
    Dim selectedDept As String
    selectedDept = Trim(CStr(wsAggr.Range(AGGR_DEPT_CELL).Value))

    Dim fromDateRaw As Variant
    Dim toDateRaw As Variant
    fromDateRaw = wsAggr.Range(AGGR_FROM_CELL).Value
    toDateRaw = wsAggr.Range(AGGR_TO_CELL).Value

    ' Validate date inputs
    If fromDateRaw <> "" And Not IsDate(fromDateRaw) Then
        MsgBox "開始日の形式が正しくありません。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    If toDateRaw <> "" And Not IsDate(toDateRaw) Then
        MsgBox "終了日の形式が正しくありません。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    Dim useFrom As Boolean: useFrom = (fromDateRaw <> "")
    Dim useTo As Boolean:   useTo = (toDateRaw <> "")
    Dim fromDate As Date
    Dim toDate As Date
    If useFrom Then fromDate = CDate(fromDateRaw)
    If useTo Then toDate = CDate(toDateRaw)

    ' Load all sheet data
    Dim wsAll As Worksheet
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)
    Dim lastRow As Long
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

    ClearAggrTable wsAggr
    If lastRow < 2 Then Exit Sub

    Dim allData As Variant
    allData = wsAll.Range(wsAll.Cells(2, 1), wsAll.Cells(lastRow, ALL_TOTAL_COLS)).Value

    ' Aggregate
    ' Key:   製品名 & "||" & 客先名
    ' Value: Array(売上金額合計, 数量合計, 口銭合計)
    Dim dictSummary As Object
    Set dictSummary = CreateObject("Scripting.Dictionary")
    dictSummary.CompareMode = vbTextCompare

    Dim r As Long
    For r = 1 To UBound(allData, 1)
        ' Dept filter
        If selectedDept <> "全部署" And selectedDept <> "" Then
            If Trim(CStr(allData(r, ALL_COL_DEPT))) <> selectedDept Then GoTo NextRow
        End If

        ' Date filter
        If useFrom Or useTo Then
            Dim saleDateRaw As Variant
            saleDateRaw = allData(r, ALL_COL_DATE)
            If Not IsDate(saleDateRaw) Then GoTo NextRow
            Dim saleDate As Date
            saleDate = CDate(saleDateRaw)
            If useFrom And saleDate < fromDate Then GoTo NextRow
            If useTo And saleDate > toDate Then GoTo NextRow
        End If

        ' Accumulate totals
        Dim pName As String: pName = Trim(CStr(allData(r, ALL_COL_PROD_NAME)))
        Dim cName As String: cName = Trim(CStr(allData(r, ALL_COL_CLIENT)))
        Dim key As String: key = pName & "||" & cName

        Dim amt As Double
        Dim qty As Double
        Dim margin As Double
        If IsNumeric(allData(r, ALL_COL_AMOUNT)) Then amt = CDbl(allData(r, ALL_COL_AMOUNT))
        If IsNumeric(allData(r, ALL_COL_QTY)) Then qty = CDbl(allData(r, ALL_COL_QTY))
        If IsNumeric(allData(r, ALL_COL_MARGIN)) Then margin = CDbl(allData(r, ALL_COL_MARGIN))

        If dictSummary.Exists(key) Then
            Dim existing As Variant
            existing = dictSummary(key)
            existing(0) = existing(0) + amt
            existing(1) = existing(1) + qty
            existing(2) = existing(2) + margin
            dictSummary(key) = existing
        Else
            dictSummary(key) = Array(amt, qty, margin)
        End If

NextRow:
    Next r

    DrawAggrTable wsAggr, dictSummary
End Sub

Private Sub ClearAggrTable(wsAggr As Worksheet)
    Dim lastRow As Long
    lastRow = wsAggr.Cells(wsAggr.Rows.Count, 1).End(xlUp).Row
    If lastRow >= AGGR_DATA_ROW Then
        wsAggr.Rows(AGGR_DATA_ROW & ":" & lastRow).Delete
    End If
End Sub

Private Sub DrawAggrTable(wsAggr As Worksheet, dictSummary As Object)
    If dictSummary.Count = 0 Then Exit Sub

    ' Collect keys and sort ascending by product name (left of "||")
    Dim keys() As String
    ReDim keys(0 To dictSummary.Count - 1)
    Dim i As Integer
    i = 0
    Dim k As Variant
    For Each k In dictSummary.Keys
        keys(i) = CStr(k)
        i = i + 1
    Next k

    ' Bubble sort
    Dim j As Integer
    Dim tmp As String
    For i = 0 To UBound(keys) - 1
        For j = 0 To UBound(keys) - i - 1
            If Split(keys(j), "||")(0) > Split(keys(j + 1), "||")(0) Then
                tmp = keys(j): keys(j) = keys(j + 1): keys(j + 1) = tmp
            End If
        Next j
    Next i

    ' Render rows
    Dim currentRow As Long: currentRow = AGGR_DATA_ROW
    Dim currentProd As String: currentProd = ""
    Dim prodStartRow As Long
    Dim prodSubAmt As Double, prodSubQty As Double, prodSubMargin As Double
    Dim totalAmt As Double, totalQty As Double, totalMargin As Double

    For i = 0 To UBound(keys)
        Dim parts() As String: parts = Split(keys(i), "||")
        Dim pName As String: pName = parts(0)
        Dim cName As String: cName = parts(1)
        Dim vals As Variant: vals = dictSummary(keys(i))

        ' New product group: insert parent row
        If pName <> currentProd Then
            prodStartRow = currentRow
            wsAggr.Cells(currentRow, 1).Value = pName
            With wsAggr.Rows(currentRow)
                .Font.Bold = True
                .Interior.Color = RGB(220, 220, 220)
            End With
            currentProd = pName
            prodSubAmt = 0: prodSubQty = 0: prodSubMargin = 0
            currentRow = currentRow + 1
        End If

        ' Client row
        wsAggr.Cells(currentRow, 1).Value = "　　" & cName  ' 2 full-width spaces
        wsAggr.Cells(currentRow, 2).Value = vals(0)
        wsAggr.Cells(currentRow, 3).Value = vals(1)
        wsAggr.Cells(currentRow, 4).Value = vals(2)
        ApplyNumFormat wsAggr.Cells(currentRow, 2)
        ApplyNumFormat wsAggr.Cells(currentRow, 3)
        ApplyNumFormat wsAggr.Cells(currentRow, 4)

        ' Update product parent row with running subtotals
        prodSubAmt = prodSubAmt + vals(0)
        prodSubQty = prodSubQty + vals(1)
        prodSubMargin = prodSubMargin + vals(2)
        wsAggr.Cells(prodStartRow, 2).Value = prodSubAmt
        wsAggr.Cells(prodStartRow, 3).Value = prodSubQty
        wsAggr.Cells(prodStartRow, 4).Value = prodSubMargin
        ApplyNumFormat wsAggr.Cells(prodStartRow, 2)
        ApplyNumFormat wsAggr.Cells(prodStartRow, 3)
        ApplyNumFormat wsAggr.Cells(prodStartRow, 4)

        totalAmt = totalAmt + vals(0)
        totalQty = totalQty + vals(1)
        totalMargin = totalMargin + vals(2)
        currentRow = currentRow + 1
    Next i

    ' Total row
    With wsAggr.Rows(currentRow)
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    End With
    wsAggr.Cells(currentRow, 1).Value = "総合計"
    wsAggr.Cells(currentRow, 2).Value = totalAmt
    wsAggr.Cells(currentRow, 3).Value = totalQty
    wsAggr.Cells(currentRow, 4).Value = totalMargin
    ApplyNumFormat wsAggr.Cells(currentRow, 2)
    ApplyNumFormat wsAggr.Cells(currentRow, 3)
    ApplyNumFormat wsAggr.Cells(currentRow, 4)
End Sub

Private Sub ApplyNumFormat(cell As Range)
    cell.NumberFormat = "#,##0"
End Sub
```

- [ ] **Step 2: Commit**

```bash
git add src/modAggregation.bas
git commit -m "feat: add modAggregation with dept/date filtering and hierarchical rendering"
```

---

### Task 6: modUIControl.bas — Orchestration and logging

**Files:**
- Create: `src/modUIControl.bas`

- [ ] **Step 1: Create src/modUIControl.bas**

```vba
Attribute VB_Name = "modUIControl"
Option Explicit

Public Sub RunAll()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    LogMessage "===== 処理開始 ====="

    ' Load masters
    LogMessage "マスタ読み込み中..."
    Dim dictProduct As Object
    Dim dictCommission As Object
    Dim dictHeaderMap As Object
    Set dictProduct = LoadProductDict()
    Set dictCommission = LoadCommissionDict()
    Set dictHeaderMap = LoadHeaderMap()
    LogMessage "  製品マスタ: " & dictProduct.Count & "件 / 口銭マスタ: " & dictCommission.Count & "件 / 名寄せ: " & dictHeaderMap.Count & "エントリ"

    ' Select files
    Dim files As Variant
    files = SelectFiles()
    If VarType(files) = vbBoolean Then
        LogMessage "ファイル選択がキャンセルされました"
        GoTo Cleanup
    End If

    LogMessage CStr(UBound(files)) & "件のファイルを読み込みます"

    ' Load each file
    Dim i As Integer
    Dim successCount As Integer
    For i = 1 To UBound(files)
        LogMessage "  読込: " & files(i)
        If LoadTsvToSheet(CStr(files(i))) Then
            successCount = successCount + 1
        Else
            LogMessage "  [エラー] 読み込み失敗: " & files(i)
        End If
    Next i
    LogMessage successCount & "件のファイルを読み込みました"

    ' Build all sheet
    LogMessage "allシート構築中..."
    BuildAllSheet dictProduct, dictCommission, dictHeaderMap
    LogMessage "allシート構築完了"

    ' Refresh dept list
    Dim dictDept As Object
    Set dictDept = CollectUniqueDepts()
    RefreshDeptList dictDept
    LogMessage "部署リスト更新完了 (" & dictDept.Count & "部署)"

    ' Rebuild aggregation (re-enable events first so the 集計 sheet renders correctly)
    Application.EnableEvents = True
    Rebuild
    LogMessage "集計完了"

    LogMessage "===== 処理完了 ====="

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    LogMessage "[エラー] " & Err.Description
    MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "エラー"
    Resume Cleanup
End Sub

Public Sub LogMessage(msg As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MAIN)

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < MAIN_LOG_START_ROW Then nextRow = MAIN_LOG_START_ROW

    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value = msg
End Sub
```

- [ ] **Step 2: Commit**

```bash
git add src/modUIControl.bas
git commit -m "feat: add modUIControl for RunAll orchestration and logging"
```

---

### Task 7: modSetup.bas — One-time workbook initialization

This module is called once by the VBScript setup script to configure sheet layouts (with Japanese text) and inject the `Worksheet_Change` event into 集計. It remains in the workbook but is never called again after setup.

**Files:**
- Create: `src/modSetup.bas`

- [ ] **Step 1: Create src/modSetup.bas**

```vba
Attribute VB_Name = "modSetup"
Option Explicit

Public Sub InitWorkbook()
    ' Step 1: Rename placeholder sheet (Sheet4 or Shuukei) to 集計
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Shuukei" Or ws.Name = "Sheet4" Or ws.Name = "Sheet3" Then
            ws.Name = SH_AGGR
            Exit For
        End If
    Next ws

    SetupMainSheet
    SetupConfigSheet
    SetupAllSheet
    SetupAggrSheet
    InjectAggrEvent
End Sub

Private Sub SetupMainSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MAIN)
    ws.Cells(1, 1).Value = "実行ログ"
    ws.Cells(2, 1).Value = "日時"
    ws.Cells(2, 2).Value = "メッセージ"
    ws.Cells(1, 1).Font.Bold = True
    With ws.Rows(2)
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 240)
    End With
    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 80

    ' Add command button
    Dim btn As Object
    Set btn = ws.Buttons.Add(10, 10, 160, 30)
    btn.Caption = "ファイルを読み込む"
    btn.OnAction = "modUIControl.RunAll"
End Sub

Private Sub SetupConfigSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    ' 製品マスタ (A1:B)
    ws.Cells(1, 1).Value = "製品マスタ"
    ws.Cells(2, 1).Value = "製品コード"
    ws.Cells(2, 2).Value = "製品名"

    ' 口銭マスタ (D1:E)
    ws.Cells(1, 4).Value = "口銭マスタ"
    ws.Cells(2, 4).Value = "売上種別"
    ws.Cells(2, 5).Value = "口銭比率%"

    ' ヘッダー名寄せ (G1:H)
    ws.Cells(1, 7).Value = "ヘッダー名寄せ設定"
    ws.Cells(2, 7).Value = "正規名"
    ws.Cells(2, 8).Value = "対応列名（カンマ区切り）"

    ' 部署リスト (J1:J)
    ws.Cells(1, 10).Value = "集計用部署リスト"
    ws.Cells(2, 10).Value = "全部署"

    ' Bold section headers
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 4).Font.Bold = True
    ws.Cells(1, 7).Font.Bold = True
    ws.Cells(1, 10).Font.Bold = True

    ' Bold column headers
    ws.Range("A2:B2").Font.Bold = True
    ws.Range("D2:E2").Font.Bold = True
    ws.Range("G2:H2").Font.Bold = True
    ws.Range("J2").Font.Bold = True

    ws.Columns("A:B").ColumnWidth = 16
    ws.Columns("D:E").ColumnWidth = 14
    ws.Columns("G:H").ColumnWidth = 20
    ws.Columns("J").ColumnWidth = 16

    ' Sample 製品マスタ data
    ws.Cells(3, 1).Value = "P001": ws.Cells(3, 2).Value = "製品A"
    ws.Cells(4, 1).Value = "P002": ws.Cells(4, 2).Value = "製品B"

    ' Sample 口銭マスタ data
    ws.Cells(3, 4).Value = "直販":  ws.Cells(3, 5).Value = 10
    ws.Cells(4, 4).Value = "代理店": ws.Cells(4, 5).Value = 5

    ' Sample 名寄せ data
    ws.Cells(3, 7).Value = "客先名":   ws.Cells(3, 8).Value = "得意先名,得意先コード,顧客名"
    ws.Cells(4, 7).Value = "製品コード": ws.Cells(4, 8).Value = "品番,ProductCode"
    ws.Cells(5, 7).Value = "売上金額":  ws.Cells(5, 8).Value = "金額,Amount,売上高"
    ws.Cells(6, 7).Value = "製品単価":  ws.Cells(6, 8).Value = "単価,定価"
    ws.Cells(7, 7).Value = "売上数量":  ws.Cells(7, 8).Value = "数量,Qty"
    ws.Cells(8, 7).Value = "売上発生日": ws.Cells(8, 8).Value = "日付,売上日,Date"
    ws.Cells(9, 7).Value = "売上種別":  ws.Cells(9, 8).Value = "取引区分,SaleType"
    ws.Cells(10, 7).Value = "部署":    ws.Cells(10, 8).Value = "部門,Dept"
End Sub

Private Sub SetupAllSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ALL)
    ws.Cells(1, ALL_COL_CLIENT).Value = "客先名"
    ws.Cells(1, ALL_COL_PROD_CODE).Value = "製品コード"
    ws.Cells(1, ALL_COL_AMOUNT).Value = "売上金額"
    ws.Cells(1, ALL_COL_UNIT_PRICE).Value = "製品単価"
    ws.Cells(1, ALL_COL_QTY).Value = "売上数量"
    ws.Cells(1, ALL_COL_DATE).Value = "売上発生日"
    ws.Cells(1, ALL_COL_SALE_TYPE).Value = "売上種別"
    ws.Cells(1, ALL_COL_DEPT).Value = "部署"
    ws.Cells(1, ALL_COL_PROD_NAME).Value = "製品名"
    ws.Cells(1, ALL_COL_MARGIN).Value = "部署取り分"
    ws.Cells(1, ALL_COL_SOURCE).Value = "ソースファイル"
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 240)
    End With
    ws.Columns("A:K").AutoFit
End Sub

Private Sub SetupAggrSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_AGGR)

    ' Filter labels (A1:B3)
    ws.Cells(1, 1).Value = "部署選択"
    ws.Cells(2, 1).Value = "開始日"
    ws.Cells(3, 1).Value = "終了日"
    ws.Range("A1:A3").Font.Bold = True
    ws.Range("B1").Value = "全部署"

    ' Aggregate header row (row 5)
    ws.Cells(AGGR_HDR_ROW, 2).Value = "売上金額合計"
    ws.Cells(AGGR_HDR_ROW, 3).Value = "売上数量合計"
    ws.Cells(AGGR_HDR_ROW, 4).Value = "口銭総額"
    With ws.Rows(AGGR_HDR_ROW)
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 240)
    End With

    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B:D").ColumnWidth = 15
End Sub

Private Sub InjectAggrEvent()
    ' Requires "Trust access to the VBA project object model" to be enabled
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_AGGR)

    Dim codeModule As Object
    Set codeModule = ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule

    Dim code As String
    code = "Option Explicit" & vbNewLine & vbNewLine & _
           "Private Sub Worksheet_Change(ByVal Target As Range)" & vbNewLine & _
           "    Dim triggerRange As Range" & vbNewLine & _
           "    Set triggerRange = Me.Range(AGGR_DEPT_CELL & "","" & AGGR_FROM_CELL & "","" & AGGR_TO_CELL)" & vbNewLine & _
           "    If Intersect(Target, triggerRange) Is Nothing Then Exit Sub" & vbNewLine & _
           "    Application.ScreenUpdating = False" & vbNewLine & _
           "    Application.Calculation = xlCalculationManual" & vbNewLine & _
           "    Application.EnableEvents = False" & vbNewLine & _
           "    On Error GoTo ErrHandler" & vbNewLine & _
           "    modAggregation.Rebuild" & vbNewLine & _
           "ErrHandler:" & vbNewLine & _
           "    Application.ScreenUpdating = True" & vbNewLine & _
           "    Application.Calculation = xlCalculationAutomatic" & vbNewLine & _
           "    Application.EnableEvents = True" & vbNewLine & _
           "End Sub"

    codeModule.AddFromString code
End Sub
```

- [ ] **Step 2: Commit**

```bash
git add src/modSetup.bas
git commit -m "feat: add modSetup for one-time workbook initialization and event injection"
```

---

### Task 8: create_workbook.vbs — Workbook generation script

**Files:**
- Create: `setup/create_workbook.vbs`

**Prerequisite (one-time manual step):**
Excel → ファイル → オプション → トラストセンター → トラストセンターの設定 → マクロの設定 → 「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」をチェック

- [ ] **Step 1: Create setup/create_workbook.vbs**

```vbscript
Option Explicit

' ============================================================
' autoSalesAggre workbook setup script
' Usage: cscript setup\create_workbook.vbs
' Prereq: Enable "Trust access to the VBA project object model"
'         in Excel Trust Center settings before running.
' ============================================================

Dim fso, scriptDir, srcPath, outputFile
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
srcPath   = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\src")) & "\"
outputFile = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\autoSalesAggre.xlsm"))

WScript.Echo "Setup started"
WScript.Echo "Source : " & srcPath
WScript.Echo "Output : " & outputFile

Dim xlApp, wb
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set wb = xlApp.Workbooks.Add

' ---- Create 4 sheets: main, Config, all, Shuukei (placeholder for 集計) ----
' Remove extra default sheets first
Do While wb.Sheets.Count > 1
    wb.Sheets(wb.Sheets.Count).Delete
Loop
wb.Sheets(1).Name = "main"
wb.Sheets.Add(After:=wb.Sheets("main")).Name = "Config"
wb.Sheets.Add(After:=wb.Sheets("Config")).Name = "all"
wb.Sheets.Add(After:=wb.Sheets("all")).Name = "Shuukei"  ' renamed to 集計 by modSetup.InitWorkbook

' ---- Import VBA modules ----
WScript.Echo "Importing VBA modules..."
Dim vbp
Set vbp = wb.VBProject

vbp.VBComponents.Import srcPath & "modConfig.bas"
vbp.VBComponents.Import srcPath & "modFileIO.bas"
vbp.VBComponents.Import srcPath & "modDataProcess.bas"
vbp.VBComponents.Import srcPath & "modAggregation.bas"
vbp.VBComponents.Import srcPath & "modUIControl.bas"
vbp.VBComponents.Import srcPath & "modSetup.bas"
WScript.Echo "Modules imported."

' ---- Run one-time setup (renames sheets, sets layouts, injects event) ----
WScript.Echo "Running InitWorkbook..."
xlApp.Run "modSetup.InitWorkbook"
WScript.Echo "InitWorkbook complete."

' ---- Save as .xlsm (52 = xlOpenXMLMacroEnabled) ----
wb.SaveAs outputFile, 52
WScript.Echo "Saved: " & outputFile

xlApp.Quit
Set xlApp = Nothing

WScript.Echo ""
WScript.Echo "Done. Open autoSalesAggre.xlsm in Excel and enable macros."
```

- [ ] **Step 2: Commit**

```bash
git add setup/create_workbook.vbs
git commit -m "feat: add VBScript to generate autoSalesAggre.xlsm from src modules"
```

---

### Task 9: CLAUDE.md — Project documentation

**Files:**
- Create: `CLAUDE.md` (in `c:\work\autoSalesAggre\`)

- [ ] **Step 1: Create CLAUDE.md**

```markdown
# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

`autoSalesAggre.xlsm` — an Excel VBA macro workbook that:
- Reads multiple TSV files with mismatched headers via a Config name-normalization table
- Consolidates rows into an "all" sheet with master lookups (product name, commission)
- Provides department- and date-filtered hierarchical aggregation in a 集計 sheet

## Setup (first time)

1. In Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings** → check "Trust access to the VBA project object model"
2. Run: `cscript setup\create_workbook.vbs`
3. Open `autoSalesAggre.xlsm` and enable macros

## Rebuild after src/ changes

Delete `autoSalesAggre.xlsm`, then re-run `cscript setup\create_workbook.vbs`.

## Module responsibilities

| Module | Responsibility |
|--------|---------------|
| `modConfig` | All constants (column indices, sheet/cell addresses) + `LoadProductDict`, `LoadCommissionDict`, `LoadHeaderMap`, `RefreshDeptList` |
| `modFileIO` | `SelectFiles` (GetOpenFilename) + `LoadTsvToSheet` (TSV → sheet, all-text format) |
| `modDataProcess` | `BuildAllSheet` (header map + master lookup + bulk write) + `CollectUniqueDepts` |
| `modAggregation` | `Rebuild` (filter allData → dictSummary → `DrawAggrTable`) |
| `modUIControl` | `RunAll` (orchestrates all modules) + `LogMessage` (callable from any module) |
| `modSetup` | One-time init: sheet naming/layout + `InjectAggrEvent` (adds Worksheet_Change to 集計 sheet module) |

## Config sheet layout

| Column | Content |
|--------|---------|
| A–B | 製品マスタ: 製品コード → 製品名 (from row 3) |
| D–E | 口銭マスタ: 売上種別 → 口銭比率% (from row 3) |
| G–H | ヘッダー名寄せ: 正規名 → カンマ区切りエイリアス (from row 3) |
| J | 集計用部署リスト: J2="全部署" (fixed), J3+ auto-updated after each RunAll |

## Key design decisions

- All cross-module calls work without qualifiers in VBA (e.g. `LogMessage` called from `modDataProcess`)
- `Application.EnableEvents = False` is set during `RunAll` to prevent Worksheet_Change firing mid-process; re-enabled before calling `Rebuild` at the end
- `dictSummary` key format: `製品名 & "||" & 客先名` — the `||` separator avoids collisions with normal text
- Source TSV data is loaded with `NumberFormat = "@"` (text) to preserve leading zeros in codes
```

- [ ] **Step 2: Commit**

```bash
git add CLAUDE.md
git commit -m "docs: add CLAUDE.md with setup instructions and architecture notes"
```

---

### Task 10: Generate and verify workbook

- [ ] **Step 1: Run setup script**

```bash
cscript setup/create_workbook.vbs
```

Expected output:
```
Setup started
Source : C:\work\autoSalesAggre\src\
Output : C:\work\autoSalesAggre\autoSalesAggre.xlsm
Importing VBA modules...
Modules imported.
Running InitWorkbook...
InitWorkbook complete.
Saved: C:\work\autoSalesAggre\autoSalesAggre.xlsm

Done. Open autoSalesAggre.xlsm in Excel and enable macros.
```

- [ ] **Step 2: Verify sheet structure in Excel**

Open `autoSalesAggre.xlsm`, enable macros. Confirm:
1. 4 tabs exist: `main`, `Config`, `all`, `集計`
2. `main` has a button "ファイルを読み込む" and a bold "実行ログ" header
3. `Config` has section headers at A1, D1, G1, J1 with sample data in rows 3–10
4. `all` has a bold header row with all 11 column names
5. `集計` has filter labels in A1:A3, a header row at row 5, "全部署" in B1

- [ ] **Step 3: Verify VBA modules**

Press Alt+F11. In Project Explorer confirm:
- `Modules` folder: `modAggregation`, `modConfig`, `modDataProcess`, `modFileIO`, `modSetup`, `modUIControl`
- `Microsoft Excel Objects` → `集計` sheet: contains `Worksheet_Change` procedure

- [ ] **Step 4: End-to-end test with sample data**

Create a file `test_sales.txt` with the content below (tab-separated):

```
客先名	製品コード	売上金額	製品単価	売上数量	売上発生日	売上種別	部署
得意先X社	P001	1000000	5000	200	2024/01/15	直販	営業部
得意先Y社	P001	500000	5000	100	2024/02/20	代理店	東京支店
得意先Z社	P002	300000	3000	100	2024/03/10	直販	営業部
```

Click "ファイルを読み込む" and select `test_sales.txt`.

Expected results:
- Sheet `test_sales` created with 3 data rows
- `all` sheet rows 2–4 populated; column I (製品名) shows "製品A" / "製品B"; column J (部署取り分) = 100000, 25000, 30000
- `集計` shows "製品A" parent row with subtotals, "製品B" parent row, 総合計 row
- `main` log shows timestamped completion entries

- [ ] **Step 5: Test dept filter**

In 集計 B1, select "営業部" from dropdown.
Expected: table updates to show 得意先X社 (P001, 1,000,000) and 得意先Z社 (P002, 300,000) only.

- [ ] **Step 6: Test date filter**

Set B1 back to "全部署". Set B2 = `2024/01/01`, B3 = `2024/01/31`.
Expected: table shows only 得意先X社 row (January 15).

- [ ] **Step 7: Commit final state**

```bash
git add CLAUDE.md src/ setup/
git status
git commit -m "chore: verify all tasks complete — autoSalesAggre implementation done"
```
