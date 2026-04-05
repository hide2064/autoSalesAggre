Attribute VB_Name = "modAggregation"
Option Explicit

' ============================================================
' modAggregation — 集計・描画・フィルター管理モジュール
'
' 役割:
'   ・集計シートのフィルタ条件（部署・期間）を読み取り、
'     all シートのデータを絞り込んで製品×客先ごとに集計する Rebuild を提供する。
'   ・集計結果を階層形式（製品グループ行 + 客先明細行 + 総合計行）で
'     集計シートに描画する DrawAggrTable を提供する。
'   ・フィルター条件の保存・復元機能 (SaveFilter / RestoreFilter) を提供する。
'
' フィルター条件の永続化:
'   SaveFilter  — 集計シートの B1/B2/B3 の値を Config シート O2/O3/O4 に保存する
'   RestoreFilter — Config シート O2/O3/O4 の値を B1/B2/B3 に復元して Rebuild する
'   Rebuild 完了後に SaveFilter が自動呼び出しされるため、最後の使用条件が
'   常に保存される。次回起動時も B1/B2/B3 は Excel のセル値として保持されており、
'   手動で「条件を復元」ボタンを押すと前回の保存条件に戻せる。
' ============================================================

' ============================================================
' Rebuild — 集計シートのフィルタを読み取り集計を再描画する
'
' 呼び出されるタイミング:
'   ・modUIControl.RunAll の最後（全データ読み込み後）
'   ・集計シートの Worksheet_Change イベント（B1/B2/B3 変更時）
'   ・RestoreFilter から（フィルター復元後）
' ============================================================
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

    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)

    ' --- フィルタ条件の読み取り ---
    selectedDept = Trim(CStr(wsAggr.Range(AGGR_DEPT_CELL).Value))
    fromDateRaw  = wsAggr.Range(AGGR_FROM_CELL).Value
    toDateRaw    = wsAggr.Range(AGGR_TO_CELL).Value

    ' --- 日付の形式チェック（空欄は許可、無効な文字列は拒否）---
    If fromDateRaw <> "" And Not IsDate(fromDateRaw) Then
        MsgBox "開始日の形式が正しくありません。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    If toDateRaw <> "" And Not IsDate(toDateRaw) Then
        MsgBox "終了日の形式が正しくありません。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    useFrom = (fromDateRaw <> "")
    useTo   = (toDateRaw   <> "")
    If useFrom Then fromDate = CDate(fromDateRaw)
    If useTo   Then toDate   = CDate(toDateRaw)

    ' --- all シートのデータを一括読み込み ---
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

    ' データが1行もない場合は集計テーブルをクリアせず終了
    If lastRow < 2 Then Exit Sub
    ClearAggrTable wsAggr

    allData = wsAll.Range(wsAll.Cells(2, 1), wsAll.Cells(lastRow, ALL_TOTAL_COLS)).Value

    ' ============================================================
    ' 集計ループ: 各行にフィルタを適用し dictSummary に累積する
    '
    ' dictSummary のキー: 「製品名 & DICT_KEY_SEP & 客先名」
    ' dictSummary の値 : Array(売上金額合計, 数量合計, 口銭合計)
    ' ============================================================
    Set dictSummary = NewDict()

    For r = 1 To UBound(allData, 1)
        ' --- 部署フィルタ ---
        If selectedDept <> "全部署" And selectedDept <> "" Then
            If Trim(CStr(allData(r, ALL_COL_DEPT))) <> selectedDept Then GoTo NextRow
        End If

        ' --- 日付フィルタ ---
        If useFrom Or useTo Then
            saleDateRaw = allData(r, ALL_COL_DATE)
            If Not IsDate(saleDateRaw) Then GoTo NextRow
            saleDate = CDate(saleDateRaw)
            If useFrom And saleDate < fromDate Then GoTo NextRow
            If useTo   And saleDate > toDate   Then GoTo NextRow
        End If

        ' --- フィルタ通過: dictSummary に累積 ---
        pName = Trim(CStr(allData(r, ALL_COL_PROD_NAME)))
        cName = Trim(CStr(allData(r, ALL_COL_CLIENT)))
        key   = pName & DICT_KEY_SEP & cName

        amt    = 0 : qty = 0 : margin = 0
        If IsNumeric(allData(r, ALL_COL_AMOUNT)) Then amt    = CDbl(allData(r, ALL_COL_AMOUNT))
        If IsNumeric(allData(r, ALL_COL_QTY))    Then qty    = CDbl(allData(r, ALL_COL_QTY))
        If IsNumeric(allData(r, ALL_COL_MARGIN))  Then margin = CDbl(allData(r, ALL_COL_MARGIN))

        If dictSummary.Exists(key) Then
            existing    = dictSummary(key)
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

    ' --- フィルター条件を自動保存（次回の「条件を復元」で使えるようにする）---
    SaveFilter
End Sub

' ============================================================
' SaveFilter — 集計シートの現在のフィルター条件を Config シートに保存する（Public）
'
' 保存先: Config シートの O列 (CFG_SAVED_FILTER_COL)
'   O2 = 部署, O3 = 開始日, O4 = 終了日
'
' 集計シートの「条件を保存」ボタンおよび Rebuild 完了後に自動呼び出しされる。
' ============================================================
Public Sub SaveFilter()
    Dim wsAggr As Worksheet
    Dim wsCfg As Worksheet

    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)
    Set wsCfg  = ThisWorkbook.Sheets(SH_CONFIG)

    wsCfg.Cells(CFG_SAVED_DEPT_ROW, CFG_SAVED_FILTER_COL).Value = _
        wsAggr.Range(AGGR_DEPT_CELL).Value
    wsCfg.Cells(CFG_SAVED_FROM_ROW, CFG_SAVED_FILTER_COL).Value = _
        wsAggr.Range(AGGR_FROM_CELL).Value
    wsCfg.Cells(CFG_SAVED_TO_ROW, CFG_SAVED_FILTER_COL).Value = _
        wsAggr.Range(AGGR_TO_CELL).Value
End Sub

' ============================================================
' RestoreFilter — Config シートに保存されたフィルター条件を集計シートに復元する（Public）
'
' 集計シートの「条件を復元」ボタンに接続される。
' 保存値が空の場合はその旨を通知して終了する。
' 復元後は Rebuild を呼び出して集計を更新する。
' ============================================================
Public Sub RestoreFilter()
    Dim wsAggr As Worksheet
    Dim wsCfg As Worksheet
    Dim savedDept As String

    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)
    Set wsCfg  = ThisWorkbook.Sheets(SH_CONFIG)

    savedDept = Trim(CStr(wsCfg.Cells(CFG_SAVED_DEPT_ROW, CFG_SAVED_FILTER_COL).Value))

    If savedDept = "" Then
        MsgBox "保存されたフィルター条件がありません。" & vbCrLf & _
               "まず「条件を保存」ボタンで保存してください。", _
               vbInformation, "保存なし"
        Exit Sub
    End If

    ' B1/B2/B3 に保存値を書き込む
    ' EnableEvents を一時停止して Worksheet_Change の多重発火を防ぐ
    Application.EnableEvents = False
    wsAggr.Range(AGGR_DEPT_CELL).Value = savedDept
    wsAggr.Range(AGGR_FROM_CELL).Value = _
        wsCfg.Cells(CFG_SAVED_FROM_ROW, CFG_SAVED_FILTER_COL).Value
    wsAggr.Range(AGGR_TO_CELL).Value = _
        wsCfg.Cells(CFG_SAVED_TO_ROW, CFG_SAVED_FILTER_COL).Value
    Application.EnableEvents = True

    Rebuild
    LogMessage "フィルター条件を復元しました [" & savedDept & "]"
End Sub

' ============================================================
' ClearAggrTable — 集計テーブルのデータ行を削除する（プライベート）
' ============================================================
Private Sub ClearAggrTable(wsAggr As Worksheet)
    Dim lastRow As Long
    lastRow = wsAggr.Cells(wsAggr.Rows.Count, 1).End(xlUp).Row
    If lastRow >= AGGR_DATA_ROW Then
        wsAggr.Rows(AGGR_DATA_ROW & ":" & lastRow).Delete
    End If
End Sub

' ============================================================
' DrawAggrTable — dictSummary の内容を集計シートに描画する（プライベート）
'
' 描画レイアウト:
'   製品グループ行: 製品名（太字・CLR_GROUP_ROW 背景）、B〜D列に小計
'   客先行        : 「　　」+ 客先名（字下げ）、B〜D列に個別値
'   総合計行      : 「総合計」（太字・上罫線）、B〜D列に全体合計
' ============================================================
Private Sub DrawAggrTable(wsAggr As Worksheet, dictSummary As Object)
    Dim keys()     As String
    Dim prodKeys() As String
    Dim i As Integer
    Dim j As Integer
    Dim tmp As String
    Dim currentRow As Long
    Dim currentProd As String
    Dim prodStartRow As Long
    Dim prodSubAmt    As Double
    Dim prodSubQty    As Double
    Dim prodSubMargin As Double
    Dim totalAmt    As Double
    Dim totalQty    As Double
    Dim totalMargin As Double
    Dim parts() As String
    Dim pName As String
    Dim cName As String
    Dim vals As Variant
    Dim k As Variant

    If dictSummary.Count = 0 Then Exit Sub

    ' --- dictSummary のキーを配列に収集 ---
    ReDim keys(0 To dictSummary.Count - 1)
    ReDim prodKeys(0 To dictSummary.Count - 1)
    i = 0
    For Each k In dictSummary.Keys
        keys(i)     = CStr(k)
        prodKeys(i) = Split(CStr(k), DICT_KEY_SEP)(0)  ' 製品名部分を事前抽出
        i = i + 1
    Next k

    ' --- 製品名部分でバブルソート（昇順）---
    For i = 0 To UBound(keys) - 1
        For j = 0 To UBound(keys) - i - 1
            If prodKeys(j) > prodKeys(j + 1) Then
                tmp = keys(j)     : keys(j)     = keys(j + 1)     : keys(j + 1)     = tmp
                tmp = prodKeys(j) : prodKeys(j) = prodKeys(j + 1) : prodKeys(j + 1) = tmp
            End If
        Next j
    Next i

    ' ============================================================
    ' 描画ループ
    ' ============================================================
    currentRow    = AGGR_DATA_ROW
    currentProd   = ""
    prodStartRow  = 0
    prodSubAmt    = 0 : prodSubQty    = 0 : prodSubMargin    = 0
    totalAmt      = 0 : totalQty      = 0 : totalMargin      = 0

    For i = 0 To UBound(keys)
        parts = Split(keys(i), DICT_KEY_SEP)
        pName = parts(0)
        cName = parts(1)
        vals  = dictSummary(keys(i))

        ' --- 製品グループが切り替わった場合: 親行を挿入 ---
        If pName <> currentProd Then
            prodStartRow = currentRow
            wsAggr.Cells(currentRow, 1).Value = pName
            With wsAggr.Rows(currentRow)
                .Font.Bold      = True
                .Interior.Color = CLR_GROUP_ROW  ' modConfig 定数: RGB(220,220,220) グレー
            End With
            ApplyNumFormat wsAggr.Cells(currentRow, 2)
            ApplyNumFormat wsAggr.Cells(currentRow, 3)
            ApplyNumFormat wsAggr.Cells(currentRow, 4)
            currentProd   = pName
            prodSubAmt    = 0 : prodSubQty    = 0 : prodSubMargin    = 0
            currentRow    = currentRow + 1
        End If

        ' --- 客先行を書き込む ---
        wsAggr.Cells(currentRow, 1).Value = "　　" & cName  ' 全角スペース2文字で字下げ
        wsAggr.Cells(currentRow, 2).Value = vals(0)
        wsAggr.Cells(currentRow, 3).Value = vals(1)
        wsAggr.Cells(currentRow, 4).Value = vals(2)
        ApplyNumFormat wsAggr.Cells(currentRow, 2)
        ApplyNumFormat wsAggr.Cells(currentRow, 3)
        ApplyNumFormat wsAggr.Cells(currentRow, 4)

        ' --- 製品グループ親行の小計を更新 ---
        prodSubAmt    = prodSubAmt    + vals(0)
        prodSubQty    = prodSubQty    + vals(1)
        prodSubMargin = prodSubMargin + vals(2)
        wsAggr.Cells(prodStartRow, 2).Value = prodSubAmt
        wsAggr.Cells(prodStartRow, 3).Value = prodSubQty
        wsAggr.Cells(prodStartRow, 4).Value = prodSubMargin

        ' --- 全体合計に累積 ---
        totalAmt    = totalAmt    + vals(0)
        totalQty    = totalQty    + vals(1)
        totalMargin = totalMargin + vals(2)
        currentRow  = currentRow + 1
    Next i

    ' --- 総合計行 ---
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

' ============================================================
' ApplyNumFormat — 数値セルに千区切り書式を適用する（プライベート）
' ============================================================
Private Sub ApplyNumFormat(cell As Range)
    cell.NumberFormat = "#,##0"
End Sub
