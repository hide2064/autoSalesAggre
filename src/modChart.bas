Attribute VB_Name = "modChart"
Option Explicit

' ============================================================
' modChart — グラフ作成モジュール
'
' 役割:
'   ・集計シートの集計テーブルを元に製品別売上グラフを作成・更新する
'     DrawAggrChart を提供する。
'   ・集計シートの「グラフ作成」ボタンの OnAction に設定されている。
'
' グラフ仕様:
'   種類    : 集合縦棒 (xlColumnClustered)
'   X 軸    : 製品名（製品グループ親行のみ・客先明細行と総合計行は除外）
'   系列1   : 売上金額合計（B列、青色）
'   系列2   : 口銭総額（D列、オレンジ色）
'   タイトル: 「製品別売上集計」+ 現在のフィルタ条件（部署・期間）
'   位置    : 集計テーブル直下（lastRow + 2 行目のTop位置から）
'   サイズ  : 幅 500pt × 高さ 300pt
'
' 親行の識別方法:
'   DrawAggrTable が客先行を "　　"(全角スペース2文字)で始まる文字列として
'   書き込む仕様を利用し、Left(cellVal, 2) <> "　　" で親行を判定する。
'   また "総合計" 行も除外する。
'
' 更新の仕組み:
'   CHART_NAME ("AggrChart") と一致する ChartObject が既に存在する場合は
'   削除してから新規作成することで、ボタンを何度押しても重複しない。
' ============================================================

' グラフオブジェクトの識別に使う名前（重複防止・削除時の検索に使用）
Private Const CHART_NAME As String = "AggrChart"

' ============================================================
' DrawAggrChart — 集計テーブルから製品別グラフを作成・更新する
'
' 処理概要:
'   1. 集計データの有無チェック
'   2. 第1パス: 製品グループ親行の件数をカウントして配列サイズを確定
'   3. 第2パス: 親行の製品名・金額・口銭値を配列に収集
'   4. 既存の AggrChart を削除（更新時の重複防止）
'   5. グラフタイトル文字列をフィルタ条件から組み立て
'   6. ChartObject を新規作成してグラフを設定
' ============================================================
Public Sub DrawAggrChart()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim cellVal As String
    Dim count As Integer      ' 製品グループ親行の件数（配列サイズ）
    Dim idx As Integer        ' 配列書き込みインデックス
    Dim labels()     As String  ' X 軸ラベル配列（製品名）
    Dim amtVals()    As Double  ' 系列1 データ配列（売上金額合計）
    Dim marginVals() As Double  ' 系列2 データ配列（口銭総額）
    Dim chartTop As Double      ' グラフの Top 位置（ポイント）
    Dim co As ChartObject       ' 既存グラフ検索用
    Dim newChart As ChartObject ' 新規作成したグラフオブジェクト
    Dim titleText As String     ' グラフタイトル文字列
    Dim dept As String
    Dim fromDate As String
    Dim toDate As String

    Set ws = ThisWorkbook.Sheets(SH_AGGR)

    ' --- データ存在チェック ---
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < AGGR_DATA_ROW Then
        MsgBox "集計データがありません。先にデータを集計してください。", _
               vbExclamation, "データなし"
        Exit Sub
    End If

    ' ============================================================
    ' 第1パス: 製品グループ親行の件数をカウント
    '
    ' 親行の判定条件:
    '   (1) 先頭が "　　"（全角スペース2文字）でない → 客先行でない
    '   (2) "総合計" でない                          → 合計行でない
    '   (3) 空欄でない
    ' ============================================================
    count = 0
    For r = AGGR_DATA_ROW To lastRow
        cellVal = CStr(ws.Cells(r, 1).Value)
        If Left(cellVal, 2) <> "　　" And cellVal <> "総合計" And Trim(cellVal) <> "" Then
            count = count + 1
        End If
    Next r

    If count = 0 Then
        MsgBox "グラフ化できる製品データがありません。", vbExclamation, "データなし"
        Exit Sub
    End If

    ' ============================================================
    ' 第2パス: 配列サイズを確定してから値を収集
    ' 第1パスで件数を把握することで ReDim Preserve の使用を避け、
    ' 配列の再アロケーションによるパフォーマンス劣化を防ぐ。
    ' ============================================================
    ReDim labels(1 To count)
    ReDim amtVals(1 To count)
    ReDim marginVals(1 To count)

    idx = 1
    For r = AGGR_DATA_ROW To lastRow
        cellVal = CStr(ws.Cells(r, 1).Value)
        If Left(cellVal, 2) <> "　　" And cellVal <> "総合計" And Trim(cellVal) <> "" Then
            labels(idx) = cellVal
            ' IsNumeric チェック: 数値でない場合はデフォルト値 0.0 を維持
            If IsNumeric(ws.Cells(r, 2).Value) Then amtVals(idx)    = CDbl(ws.Cells(r, 2).Value)
            If IsNumeric(ws.Cells(r, 4).Value) Then marginVals(idx) = CDbl(ws.Cells(r, 4).Value)
            idx = idx + 1
        End If
    Next r

    ' --- 既存の AggrChart を削除（ボタン再押し時の重複防止）---
    For Each co In ws.ChartObjects
        If co.Name = CHART_NAME Then
            co.Delete
            Exit For
        End If
    Next co

    ' --- グラフタイトルをフィルタ条件から組み立て ---
    ' 部署が "全部署" 以外の場合は "[部署名]" を付加
    ' 日付フィルタが設定されている場合は "(開始日 ～ 終了日)" を付加
    dept     = Trim(CStr(ws.Range(AGGR_DEPT_CELL).Value))
    fromDate = Trim(CStr(ws.Range(AGGR_FROM_CELL).Value))
    toDate   = Trim(CStr(ws.Range(AGGR_TO_CELL).Value))

    titleText = "製品別売上集計"
    If dept <> "" And dept <> "全部署" Then titleText = titleText & "　[" & dept & "]"
    If fromDate <> "" Or toDate <> "" Then
        titleText = titleText & "　(" & fromDate & " ～ " & toDate & ")"
    End If

    ' --- ChartObject を集計テーブル直下に作成 ---
    ' lastRow + 2 行目の Top 座標を取得して1行分の余白を確保
    chartTop = ws.Cells(lastRow + 2, 1).Top
    Set newChart = ws.ChartObjects.Add( _
        Left:=ws.Cells(1, 1).Left, _  ' A列の左端に揃える
        Top:=chartTop, _
        Width:=500, _
        Height:=300)
    newChart.Name = CHART_NAME

    ' ============================================================
    ' グラフの詳細設定
    ' ============================================================
    With newChart.Chart
        .ChartType = xlColumnClustered  ' 集合縦棒グラフ

        ' --- 系列1: 売上金額合計（青色）---
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name    = "売上金額合計"
            .Values  = amtVals       ' Y 軸データ
            .XValues = labels        ' X 軸ラベル（系列1 に設定することで全系列に適用）
            .Interior.Color = RGB(70, 130, 180)  ' スチールブルー
        End With

        ' --- 系列2: 口銭総額（オレンジ色）---
        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .Name   = "口銭総額"
            .Values = marginVals
            .Interior.Color = RGB(255, 165, 0)  ' オレンジ
        End With

        ' --- タイトル設定 ---
        .HasTitle          = True
        .ChartTitle.Text   = titleText
        .ChartTitle.Font.Size = 12

        ' --- 軸設定 ---
        .Axes(xlCategory).HasTitle = False             ' X 軸タイトルは不要（製品名が明確なため）
        .Axes(xlValue).HasTitle    = True
        .Axes(xlValue).AxisTitle.Text     = "金額（円）"
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"  ' 縦軸の目盛を千区切り表示

        ' --- 凡例設定 ---
        .HasLegend        = True
        .Legend.Position  = xlLegendPositionBottom  ' 凡例をグラフ下部に表示

        ' --- データラベル（系列1のみ: 棒の上部に金額を表示）---
        .SeriesCollection(1).HasDataLabels = True
        With .SeriesCollection(1).DataLabels
            .NumberFormat = "#,##0"               ' 千区切り表示
            .Position     = xlLabelPositionOutsideEnd  ' 棒の上端の外側
            .Font.Size    = 8
        End With

        ' --- プロットエリアの背景色をわずかに調整（視認性向上）---
        .PlotArea.Interior.Color = RGB(248, 248, 248)  ' 薄いグレー背景
    End With

    LogMessage "グラフを作成しました（" & count & "製品）"
    MsgBox "グラフを作成しました。(" & count & "製品)", vbInformation, "完了"
End Sub
