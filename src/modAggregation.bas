Attribute VB_Name = "modAggregation"
Option Explicit

' ============================================================
' modAggregation — 集計・描画モジュール
'
' 役割:
'   ・集計シートのフィルタ条件（部署・期間）を読み取り、
'     all シートのデータを絞り込んで製品×客先ごとに集計する Rebuild を提供する。
'   ・集計結果を階層形式（製品グループ行 + 客先明細行 + 総合計行）で
'     集計シートに描画する DrawAggrTable を提供する。
'
' 集計テーブルの構造（集計シート）:
'   親行  : 製品名（太字・グレー背景）— 配下の客先行の小計
'   子行  : 「　　」+ 客先名（字下げ）— 個別金額
'   最終行: 「総合計」（太字・上罫線）— 全行の合計
'
' dictSummary のキー形式:
'   「製品名 & "||" & 客先名」
'   "||" はセパレータとして通常テキストに含まれにくい文字列。
'   同一製品名・客先名の組み合わせで金額・数量・口銭を累積する。
' ============================================================

' ============================================================
' Rebuild — 集計シートのフィルタを読み取り集計を再描画する
'
' 呼び出されるタイミング:
'   ・modUIControl.RunAll の最後（全データ読み込み後）
'   ・集計シートの Worksheet_Change イベント（B1/B2/B3 変更時）
'
' 処理概要:
'   1. フィルタ条件（部署・開始日・終了日）を集計シートから読み取る
'   2. 日付が入力されている場合は形式チェックを行う
'   3. all シートのデータを Variant 配列に一括読み込み
'   4. 各行についてフィルタ判定を行い、通過した行を dictSummary に累積
'   5. DrawAggrTable に dictSummary を渡して描画する
' ============================================================
Public Sub Rebuild()
    Dim wsAggr As Worksheet
    Dim selectedDept As String  ' B1: 部署選択（"全部署" または特定部署名）
    Dim fromDateRaw As Variant  ' B2: 開始日入力値（空欄時は ""）
    Dim toDateRaw As Variant    ' B3: 終了日入力値（空欄時は ""）
    Dim useFrom As Boolean      ' 開始日フィルタを適用するか
    Dim useTo As Boolean        ' 終了日フィルタを適用するか
    Dim fromDate As Date
    Dim toDate As Date
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim allData As Variant      ' all シートデータの一括読み込み用
    Dim dictSummary As Object   ' キー: 製品名 & DICT_KEY_SEP & 客先名、値: Array(金額合計, 数量合計, 口銭合計)
    Dim r As Long
    Dim saleDateRaw As Variant
    Dim saleDate As Date
    Dim pName As String         ' 製品名
    Dim cName As String         ' 客先名
    Dim key As String           ' dictSummary のキー文字列
    Dim amt As Double
    Dim qty As Double
    Dim margin As Double
    Dim existing As Variant     ' 既存の集計値配列（累積用）

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
    ' （ClearAggrTable を呼ばないことで既存の表示を保持する）
    If lastRow < 2 Then Exit Sub
    ClearAggrTable wsAggr

    allData = wsAll.Range(wsAll.Cells(2, 1), wsAll.Cells(lastRow, ALL_TOTAL_COLS)).Value

    ' ============================================================
    ' 集計ループ: 各行にフィルタを適用し dictSummary に累積する
    '
    ' dictSummary のキー: 「製品名 & DICT_KEY_SEP & 客先名」(DICT_KEY_SEP は modConfig 定数)
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
            If Not IsDate(saleDateRaw) Then GoTo NextRow  ' 日付に変換できない行はスキップ
            saleDate = CDate(saleDateRaw)
            If useFrom And saleDate < fromDate Then GoTo NextRow
            If useTo   And saleDate > toDate   Then GoTo NextRow
        End If

        ' --- フィルタ通過: 金額・数量・口銭を dictSummary に累積 ---
        pName = Trim(CStr(allData(r, ALL_COL_PROD_NAME)))
        cName = Trim(CStr(allData(r, ALL_COL_CLIENT)))
        key   = pName & DICT_KEY_SEP & cName  ' DICT_KEY_SEP(modConfig) で衝突を防ぐ

        amt    = 0 : qty = 0 : margin = 0
        If IsNumeric(allData(r, ALL_COL_AMOUNT)) Then amt    = CDbl(allData(r, ALL_COL_AMOUNT))
        If IsNumeric(allData(r, ALL_COL_QTY))    Then qty    = CDbl(allData(r, ALL_COL_QTY))
        If IsNumeric(allData(r, ALL_COL_MARGIN))  Then margin = CDbl(allData(r, ALL_COL_MARGIN))

        If dictSummary.Exists(key) Then
            ' 既存エントリに累積（配列を取り出して更新後に再格納）
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
End Sub

' ============================================================
' ClearAggrTable — 集計テーブルのデータ行を削除する（プライベート）
'
' AGGR_DATA_ROW(6行目) 以降の行を削除する。
' ヘッダー行(5行目)・フィルタ欄(1〜3行目)は保持する。
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
' 引数:
'   wsAggr      — 描画先の集計シート
'   dictSummary — Rebuild が構築した「製品名||客先名 → Array(金額,数量,口銭)」辞書
'
' 描画レイアウト:
'   製品グループ行: 製品名（太字・グレー背景）、B〜D列に小計
'   客先行        : 「　　」+ 客先名（字下げ）、B〜D列に個別値
'   総合計行      : 「総合計」（太字・上罫線）、B〜D列に全体合計
'
' 製品グループ行の小計は客先行の追加ごとにリアルタイムで更新する。
' 製品名でバブルソートしてから描画するため、表示順は常に製品名昇順。
' ============================================================
Private Sub DrawAggrTable(wsAggr As Worksheet, dictSummary As Object)
    Dim keys()     As String  ' dictSummary のキーを格納する配列（ソート用）
    Dim prodKeys() As String  ' ソート比較用に Split 済みの製品名部分
    Dim i As Integer
    Dim j As Integer
    Dim tmp As String
    Dim currentRow As Long    ' 現在の書き込み行
    Dim currentProd As String ' 処理中の製品名（グループ切り替え検知用）
    Dim prodStartRow As Long  ' 現在の製品グループの親行の行番号
    Dim prodSubAmt    As Double  ' 製品グループの小計（売上金額）
    Dim prodSubQty    As Double  ' 製品グループの小計（数量）
    Dim prodSubMargin As Double  ' 製品グループの小計（口銭）
    Dim totalAmt    As Double    ' 全体合計（売上金額）
    Dim totalQty    As Double    ' 全体合計（数量）
    Dim totalMargin As Double    ' 全体合計（口銭）
    Dim parts() As String  ' key を "||" で分割した結果
    Dim pName As String
    Dim cName As String
    Dim vals As Variant    ' dictSummary の値 Array(金額,数量,口銭)
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
    ' prodKeys に事前抽出済みのため、ループ内で Split を繰り返さない
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
    ' keys を順に処理し、製品名が切り替わるタイミングで
    ' 新しいグループ親行を挿入する。
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
                .Interior.Color = RGB(220, 220, 220)  ' グレー背景
            End With
            ' 数値書式を設定（小計はこの後に累積更新されるが書式は1回だけ適用）
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

        ' --- 製品グループ親行の小計をリアルタイムで更新 ---
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

    ' --- 総合計行を書き込む ---
    With wsAggr.Rows(currentRow)
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous  ' 上罫線で視覚的に区切る
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
'
' "#,##0" で統一することで 1234567 → 1,234,567 と表示する。
' 複数箇所で同じ書式を使うため1関数に集約した。
' ============================================================
Private Sub ApplyNumFormat(cell As Range)
    cell.NumberFormat = "#,##0"
End Sub
