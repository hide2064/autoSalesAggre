Attribute VB_Name = "modDataProcess"
Option Explicit

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

    ' colMap(i) = TSV列番号 (0=未マップ)
    ReDim colMap(1 To N)

    For c = 1 To lastSrcCol
        srcHeader = LCase(Trim(CStr(wsSrc.Cells(1, c).Value)))
        If dictHeaderMap.Exists(srcHeader) Then
            canonical = dictHeaderMap(srcHeader)
            For i = 1 To N
                If canonicalKeys(i) = canonical Then
                    colMap(i) = c
                    Exit For
                End If
            Next i
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
