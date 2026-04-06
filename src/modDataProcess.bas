Attribute VB_Name = "modDataProcess"
Option Explicit

Public Sub BuildAllSheet(dictProduct As Object, dictCommission As Object, dictHeaderMap As Object)
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim allRowNum As Long

    Set wsAll = ThisWorkbook.Sheets(SH_ALL)

    ' Clear data rows, keep header
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then wsAll.Rows("2:" & lastRow).ClearContents

    ' Write header row
    wsAll.Cells(1, ALL_COL_CLIENT).Value = HDR_CLIENT
    wsAll.Cells(1, ALL_COL_PROD_CODE).Value = HDR_PROD_CODE
    wsAll.Cells(1, ALL_COL_AMOUNT).Value = HDR_AMOUNT
    wsAll.Cells(1, ALL_COL_UNIT_PRICE).Value = HDR_UNIT_PRICE
    wsAll.Cells(1, ALL_COL_QTY).Value = HDR_QTY
    wsAll.Cells(1, ALL_COL_DATE).Value = HDR_DATE
    wsAll.Cells(1, ALL_COL_SALE_TYPE).Value = HDR_SALE_TYPE
    wsAll.Cells(1, ALL_COL_DEPT).Value = HDR_DEPT
    wsAll.Cells(1, ALL_COL_PROD_NAME).Value = HDR_PROD_NAME
    wsAll.Cells(1, ALL_COL_MARGIN).Value = HDR_MARGIN
    wsAll.Cells(1, ALL_COL_SOURCE).Value = HDR_SOURCE

    allRowNum = 2

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
    Dim lastSrcCol As Integer
    Dim colMap(7) As Integer
    Dim c As Integer
    Dim srcHeader As String
    Dim srcData As Variant
    Dim numRows As Long
    Dim outArr() As Variant
    Dim r As Long
    Dim allCol As Integer
    Dim prodCode As String
    Dim saleType As String
    Dim amount As Double

    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    If lastSrcRow < 2 Then
        ProcessSourceSheet = startRow
        Exit Function
    End If

    lastSrcCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    ' colMap(allColIndex - 1) = source column number (0 = unmapped)
    For c = 1 To lastSrcCol
        srcHeader = LCase(Trim(CStr(wsSrc.Cells(1, c).Value)))
        If dictHeaderMap.Exists(srcHeader) Then
            Select Case dictHeaderMap(srcHeader)
                Case HDR_CLIENT:      colMap(ALL_COL_CLIENT - 1) = c
                Case HDR_PROD_CODE:   colMap(ALL_COL_PROD_CODE - 1) = c
                Case HDR_AMOUNT:      colMap(ALL_COL_AMOUNT - 1) = c
                Case HDR_UNIT_PRICE:  colMap(ALL_COL_UNIT_PRICE - 1) = c
                Case HDR_QTY:         colMap(ALL_COL_QTY - 1) = c
                Case HDR_DATE:        colMap(ALL_COL_DATE - 1) = c
                Case HDR_SALE_TYPE:   colMap(ALL_COL_SALE_TYPE - 1) = c
                Case HDR_DEPT:        colMap(ALL_COL_DEPT - 1) = c
            End Select
        End If
    Next c

    ' Bulk read source data into Variant array
    srcData = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastSrcRow, lastSrcCol)).Value

    numRows = lastSrcRow - 1
    ReDim outArr(1 To numRows, 1 To ALL_TOTAL_COLS)

    For r = 1 To numRows
        ' Copy source columns ALL_COL_CLIENT to ALL_COL_DEPT (cols 1-8)
        For allCol = ALL_COL_CLIENT To ALL_COL_DEPT
            If colMap(allCol - 1) > 0 Then
                outArr(r, allCol) = srcData(r, colMap(allCol - 1))
            Else
                outArr(r, allCol) = ""
            End If
        Next allCol

        ' Calculate êªïiñº (col 9)
        prodCode = Trim(CStr(outArr(r, ALL_COL_PROD_CODE)))
        If dictProduct.Exists(prodCode) Then
            outArr(r, ALL_COL_PROD_NAME) = dictProduct(prodCode)
        Else
            outArr(r, ALL_COL_PROD_NAME) = "[ñ¢ìoò^]"
            If prodCode <> "" Then
                LogMessage "åxçê: êªïiÉRÅ[Éhñ¢ìoò^ [" & prodCode & "] (" & wsSrc.Name & ")"
            End If
        End If

        ' Calculate ïîèêéÊÇËï™ (col 10)
        saleType = Trim(CStr(outArr(r, ALL_COL_SALE_TYPE)))
        amount = 0
        If IsNumeric(outArr(r, ALL_COL_AMOUNT)) Then amount = CDbl(outArr(r, ALL_COL_AMOUNT))
        If dictCommission.Exists(saleType) Then
            outArr(r, ALL_COL_MARGIN) = amount * dictCommission(saleType) / 100
        Else
            outArr(r, ALL_COL_MARGIN) = 0
            If saleType <> "" Then
                LogMessage "åxçê: îÑè„éÌï ñ¢ìoò^ [" & saleType & "] (" & wsSrc.Name & ")"
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
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim deptArr As Variant
    Dim r As Long
    Dim dept As String

    Set dict = NewDict()

    Set wsAll = ThisWorkbook.Sheets(SH_ALL)

    lastRow = wsAll.Cells(wsAll.Rows.Count, ALL_COL_DEPT).End(xlUp).Row
    If lastRow < 2 Then
        Set CollectUniqueDepts = dict
        Exit Function
    End If

    deptArr = wsAll.Range(wsAll.Cells(2, ALL_COL_DEPT), wsAll.Cells(lastRow, ALL_COL_DEPT)).Value

    For r = 1 To UBound(deptArr, 1)
        dept = Trim(CStr(deptArr(r, 1)))
        If dept <> "" And Not dict.Exists(dept) Then dict(dept) = 1
    Next r

    Set CollectUniqueDepts = dict
End Function
