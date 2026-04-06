Attribute VB_Name = "modAggregation"
Option Explicit

Public Sub Rebuild()
    Dim wsAggr As Worksheet
    Dim selectedDept As String
    Dim fromDateRaw As Variant
    Dim toDateRaw As Variant
    Dim useFrom As Boolean
    Dim useTo As Boolean
    Dim fromDate As Date
    Dim toDate As Date
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim allData As Variant
    Dim dictSummary As Object
    Dim r As Long
    Dim saleDateRaw As Variant
    Dim saleDate As Date
    Dim pName As String
    Dim cName As String
    Dim key As String
    Dim amt As Double
    Dim qty As Double
    Dim margin As Double
    Dim existing As Variant
    Dim totalCols As Integer
    Dim colDept     As Integer
    Dim colDate     As Integer
    Dim colClient   As Integer
    Dim colAmount   As Integer
    Dim colQty      As Integer
    Dim colProdName As Integer
    Dim colMargin   As Integer

    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)

    ' Read filter conditions
    selectedDept = Trim(CStr(wsAggr.Range(AGGR_DEPT_CELL).Value))
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

    useFrom = (fromDateRaw <> "")
    useTo = (toDateRaw <> "")
    If useFrom Then fromDate = CDate(fromDateRaw)
    If useTo Then toDate = CDate(toDateRaw)

    ' Load all sheet data
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then Exit Sub
    ' allが空のときは集計表を保持するため、データ件数ガードの後でクリアする
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

    ' Aggregate into dictSummary
    ' Key:   製品名 & "||" & 客先名
    ' Value: Array(売上金額合計, 数量合計, 口銭合計)
    Set dictSummary = NewDict()

    For r = 1 To UBound(allData, 1)
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
        key = pName & AGGR_KEY_SEP & cName

        amt = 0: qty = 0: margin = 0
        If colAmount > 0 Then
            If IsNumeric(allData(r, colAmount)) Then amt = CDbl(allData(r, colAmount))
        End If
        If colQty > 0 Then
            If IsNumeric(allData(r, colQty)) Then qty = CDbl(allData(r, colQty))
        End If
        If colMargin > 0 Then
            If IsNumeric(allData(r, colMargin)) Then margin = CDbl(allData(r, colMargin))
        End If

        If dictSummary.Exists(key) Then
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
    Dim keys() As String
    Dim i As Integer
    Dim j As Integer
    Dim tmp As String
    Dim currentRow As Long
    Dim currentProd As String
    Dim prodStartRow As Long
    Dim prodSubAmt As Double
    Dim prodSubQty As Double
    Dim prodSubMargin As Double
    Dim totalAmt As Double
    Dim totalQty As Double
    Dim totalMargin As Double
    Dim parts() As String
    Dim pName As String
    Dim cName As String
    Dim vals As Variant
    Dim k As Variant

    If dictSummary.Count = 0 Then Exit Sub

    ' Collect keys and sort ascending by product name (left of "||")
    ReDim keys(0 To dictSummary.Count - 1)
    i = 0
    For Each k In dictSummary.Keys
        keys(i) = CStr(k)
        i = i + 1
    Next k

    ' Bubble sort by product name portion
    For i = 0 To UBound(keys) - 1
        For j = 0 To UBound(keys) - i - 1
            If Split(keys(j), AGGR_KEY_SEP)(0) > Split(keys(j + 1), AGGR_KEY_SEP)(0) Then
                tmp = keys(j): keys(j) = keys(j + 1): keys(j + 1) = tmp
            End If
        Next j
    Next i

    ' Render rows
    currentRow = AGGR_DATA_ROW
    currentProd = ""
    prodStartRow = 0
    prodSubAmt = 0: prodSubQty = 0: prodSubMargin = 0
    totalAmt = 0: totalQty = 0: totalMargin = 0

    For i = 0 To UBound(keys)
        parts = Split(keys(i), AGGR_KEY_SEP)
        pName = parts(0)
        cName = parts(1)
        vals = dictSummary(keys(i))

        ' New product group: insert parent row
        If pName <> currentProd Then
            prodStartRow = currentRow
            wsAggr.Cells(currentRow, 1).Value = pName
            With wsAggr.Rows(currentRow)
                .Font.Bold = True
                .Interior.Color = RGB(220, 220, 220)
            End With
            ' Apply number format once when parent row is created
            ApplyNumFormat wsAggr.Cells(currentRow, 2)
            ApplyNumFormat wsAggr.Cells(currentRow, 3)
            ApplyNumFormat wsAggr.Cells(currentRow, 4)
            currentProd = pName
            prodSubAmt = 0: prodSubQty = 0: prodSubMargin = 0
            currentRow = currentRow + 1
        End If

        ' Client row
        wsAggr.Cells(currentRow, 1).Value = AGGR_INDENT & cName
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
