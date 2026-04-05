Attribute VB_Name = "modMonthly"
Option Explicit

' ============================================================
' modMonthly — 月次サマリーモジュール
'
' 役割:
'   ・all シートのデータを年月ごとに集計し、「月次サマリー」シートに
'     書き出す BuildMonthly を提供する。
'   ・RunAll 完了後に自動呼び出しされるほか、「月次サマリー更新」
'     ボタンからも手動実行できる。
'
' 月次サマリーシートのレイアウト:
'   1行目: タイトル
'   2行目: ヘッダー（年月 / 売上金額合計 / 数量合計 / 取り分合計 / 件数）
'   3行目〜: 月別データ行（年月の昇順）
'   最終行: 合計行（太字・上罫線）
'
' 年月のキー形式:
'   ソート用: "YYYYMM" (文字列比較で昇順ソート可能)
'   表示用  : "YYYY年MM月"
' ============================================================

' --- 月次サマリーシートの行番号定数 ---
Private Const MO_HDR_ROW  As Integer = 2  ' ヘッダー行番号
Private Const MO_DATA_ROW As Integer = 3  ' データ開始行番号

' --- 月次サマリーシートの列インデックス ---
Private Const MO_COL_MONTH   As Integer = 1  ' A: 年月
Private Const MO_COL_AMOUNT  As Integer = 2  ' B: 売上金額合計
Private Const MO_COL_QTY     As Integer = 3  ' C: 売上数量合計
Private Const MO_COL_MARGIN  As Integer = 4  ' D: 部署取り分合計
Private Const MO_COL_COUNT   As Integer = 5  ' E: レコード数

' ============================================================
' BuildMonthly — 月次サマリーを作成・更新する（Public）
'
' 処理概要:
'   1. all シートのデータを Variant 配列に一括読み込み
'   2. 各行の売上発生日から "YYYYMM" キーを生成して dictMo に累積
'      dictMo の値: Array(売上金額合計, 数量合計, 取り分合計, 件数, 表示文字列)
'   3. YYYYMM キーをバブルソート（昇順）
'   4. 月次サマリーシートをクリアして書き出し
' ============================================================
Public Sub BuildMonthly()
    Dim wsAll  As Worksheet
    Dim wsMo   As Worksheet
    Dim lastRow As Long
    Dim allData As Variant
    Dim r As Long
    Dim dateRaw As Variant
    Dim sortKey As String   ' ソート用: "YYYYMM"
    Dim dispKey As String   ' 表示用: "YYYY年MM月"
    Dim dictMo As Object    ' キー: sortKey、値: Array(金額, 数量, 取り分, 件数, dispKey)
    Dim existing As Variant
    Dim amt As Double, qty As Double, margin As Double
    Dim keys() As String
    Dim i As Integer, j As Integer
    Dim tmp As String
    Dim k As Variant
    Dim wRow As Long
    Dim vals As Variant
    Dim totalAmt As Double, totalQty As Double
    Dim totalMargin As Double, totalCount As Long

    Set wsAll = ThisWorkbook.Sheets(SH_ALL)
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        LogMessage "月次サマリー: allシートにデータがありません（スキップ）"
        Exit Sub
    End If

    ' --- all シートのデータを一括読み込み ---
    allData = wsAll.Range(wsAll.Cells(2, 1), wsAll.Cells(lastRow, ALL_TOTAL_COLS)).Value

    ' ============================================================
    ' 年月ごとの集計
    '
    ' 日付に変換できない行はスキップし、有効な行だけを集計する。
    ' dictMo の値は 5要素の Array: (金額合計, 数量合計, 取り分合計, 件数, 表示文字列)
    ' ============================================================
    Set dictMo = NewDict()

    For r = 1 To UBound(allData, 1)
        dateRaw = allData(r, ALL_COL_DATE)
        If Not IsDate(dateRaw) Then GoTo NextRow

        sortKey = Format(CDate(dateRaw), "yyyymm")     ' "202601" 形式でソート可
        dispKey = Format(CDate(dateRaw), "yyyy年mm月") ' "2026年01月" 形式で表示

        amt    = 0 : qty = 0 : margin = 0
        If IsNumeric(allData(r, ALL_COL_AMOUNT)) Then amt    = CDbl(allData(r, ALL_COL_AMOUNT))
        If IsNumeric(allData(r, ALL_COL_QTY))    Then qty    = CDbl(allData(r, ALL_COL_QTY))
        If IsNumeric(allData(r, ALL_COL_MARGIN))  Then margin = CDbl(allData(r, ALL_COL_MARGIN))

        If dictMo.Exists(sortKey) Then
            existing    = dictMo(sortKey)
            existing(0) = existing(0) + amt
            existing(1) = existing(1) + qty
            existing(2) = existing(2) + margin
            existing(3) = existing(3) + 1
            dictMo(sortKey) = existing
        Else
            dictMo(sortKey) = Array(amt, qty, margin, 1, dispKey)
        End If

NextRow:
    Next r

    If dictMo.Count = 0 Then
        LogMessage "月次サマリー: 有効な日付データがありません"
        Exit Sub
    End If

    ' --- "YYYYMM" キーをバブルソート（昇順）---
    ' "YYYYMM" は文字列の辞書順 = 時系列順のため、文字列比較でソートできる
    ReDim keys(0 To dictMo.Count - 1)
    i = 0
    For Each k In dictMo.Keys
        keys(i) = CStr(k)
        i = i + 1
    Next k

    For i = 0 To UBound(keys) - 1
        For j = 0 To UBound(keys) - i - 1
            If keys(j) > keys(j + 1) Then
                tmp = keys(j) : keys(j) = keys(j + 1) : keys(j + 1) = tmp
            End If
        Next j
    Next i

    ' --- 月次サマリーシートの既存データをクリア ---
    Set wsMo = ThisWorkbook.Sheets(SH_MONTHLY)
    Dim clearLast As Long
    clearLast = wsMo.Cells(wsMo.Rows.Count, MO_COL_MONTH).End(xlUp).Row
    If clearLast >= MO_DATA_ROW Then
        wsMo.Rows(MO_DATA_ROW & ":" & clearLast).Delete
    End If

    ' --- ソート済み月順に書き出し ---
    wRow = MO_DATA_ROW
    totalAmt = 0 : totalQty = 0 : totalMargin = 0 : totalCount = 0

    For i = 0 To UBound(keys)
        vals = dictMo(keys(i))
        wsMo.Cells(wRow, MO_COL_MONTH).Value  = vals(4)  ' "YYYY年MM月"
        wsMo.Cells(wRow, MO_COL_AMOUNT).Value = vals(0)
        wsMo.Cells(wRow, MO_COL_QTY).Value    = vals(1)
        wsMo.Cells(wRow, MO_COL_MARGIN).Value = vals(2)
        wsMo.Cells(wRow, MO_COL_COUNT).Value  = vals(3)
        ' 数値列に千区切り書式を適用
        wsMo.Range(wsMo.Cells(wRow, MO_COL_AMOUNT), _
                   wsMo.Cells(wRow, MO_COL_MARGIN)).NumberFormat = "#,##0"

        totalAmt    = totalAmt    + vals(0)
        totalQty    = totalQty    + vals(1)
        totalMargin = totalMargin + vals(2)
        totalCount  = totalCount  + CLng(vals(3))
        wRow = wRow + 1
    Next i

    ' --- 合計行（太字・上罫線）---
    With wsMo.Rows(wRow)
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    End With
    wsMo.Cells(wRow, MO_COL_MONTH).Value  = "合計"
    wsMo.Cells(wRow, MO_COL_AMOUNT).Value = totalAmt
    wsMo.Cells(wRow, MO_COL_QTY).Value    = totalQty
    wsMo.Cells(wRow, MO_COL_MARGIN).Value = totalMargin
    wsMo.Cells(wRow, MO_COL_COUNT).Value  = totalCount
    wsMo.Range(wsMo.Cells(wRow, MO_COL_AMOUNT), _
               wsMo.Cells(wRow, MO_COL_MARGIN)).NumberFormat = "#,##0"

    LogMessage "月次サマリーを更新しました (" & dictMo.Count & "ヶ月分、" & totalCount & "件)"
End Sub
