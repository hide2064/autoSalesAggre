Attribute VB_Name = "modChart"
Option Explicit

Private Const CHART_NAME As String = "AggrChart"

' 集計シートの製品グループ行を元に縦棒グラフを作成（再実行で更新）
'
' グラフ構成:
'   種類   : 集合縦棒 (xlColumnClustered)
'   X軸   : 製品名（親行のみ・客先行/総合計は除外）
'   系列1  : 売上金額合計 (B列)
'   系列2  : 口銭総額     (D列)
'   位置   : 集計テーブル直下、A列左端から幅500px
Public Sub DrawAggrChart()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim cellVal As String
    Dim count As Integer
    Dim idx As Integer
    Dim labels() As String
    Dim amtVals() As Double
    Dim marginVals() As Double
    Dim chartTop As Double
    Dim co As ChartObject
    Dim newChart As ChartObject
    Dim titleText As String
    Dim dept As String
    Dim fromDate As String
    Dim toDate As String

    Set ws = ThisWorkbook.Sheets(SH_AGGR)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < AGGR_DATA_ROW Then
        MsgBox "集計データがありません。先にデータを集計してください。", _
               vbExclamation, "データなし"
        Exit Sub
    End If

    ' --- 第1パス: 製品グループ行（親行）の件数カウント ---
    ' 親行の識別: 客先行は "　　"(全角スペース2文字)で始まる、総合計行は除外
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

    ' --- 第2パス: 配列に値を収集 ---
    ReDim labels(1 To count)
    ReDim amtVals(1 To count)
    ReDim marginVals(1 To count)

    idx = 1
    For r = AGGR_DATA_ROW To lastRow
        cellVal = CStr(ws.Cells(r, 1).Value)
        If Left(cellVal, 2) <> "　　" And cellVal <> "総合計" And Trim(cellVal) <> "" Then
            labels(idx) = cellVal
            If IsNumeric(ws.Cells(r, 2).Value) Then
                amtVals(idx) = CDbl(ws.Cells(r, 2).Value)
            End If
            If IsNumeric(ws.Cells(r, 4).Value) Then
                marginVals(idx) = CDbl(ws.Cells(r, 4).Value)
            End If
            idx = idx + 1
        End If
    Next r

    ' --- 既存グラフを削除（更新用）---
    For Each co In ws.ChartObjects
        If co.Name = CHART_NAME Then
            co.Delete
            Exit For
        End If
    Next co

    ' --- グラフタイトル文字列を組み立て ---
    dept     = Trim(CStr(ws.Range(AGGR_DEPT_CELL).Value))
    fromDate = Trim(CStr(ws.Range(AGGR_FROM_CELL).Value))
    toDate   = Trim(CStr(ws.Range(AGGR_TO_CELL).Value))

    titleText = "製品別売上集計"
    If dept <> "" And dept <> "全部署" Then titleText = titleText & "　[" & dept & "]"
    If fromDate <> "" Or toDate <> "" Then
        titleText = titleText & "　(" & fromDate & " ～ " & toDate & ")"
    End If

    ' --- チャートオブジェクトを作成（テーブル直下に配置）---
    chartTop = ws.Cells(lastRow + 2, 1).Top
    Set newChart = ws.ChartObjects.Add( _
        Left:=ws.Cells(1, 1).Left, _
        Top:=chartTop, _
        Width:=500, _
        Height:=300)
    newChart.Name = CHART_NAME

    ' --- グラフ設定 ---
    With newChart.Chart
        .ChartType = xlColumnClustered

        ' 系列1: 売上金額合計
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "売上金額合計"
            .Values = amtVals
            .XValues = labels
            .Interior.Color = RGB(70, 130, 180)
        End With

        ' 系列2: 口銭総額
        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .Name = "口銭総額"
            .Values = marginVals
            .Interior.Color = RGB(255, 165, 0)
        End With

        ' タイトル
        .HasTitle = True
        .ChartTitle.Text = titleText
        .ChartTitle.Font.Size = 12

        ' 軸
        .Axes(xlCategory).HasTitle = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "金額（円）"
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"

        ' 凡例
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        ' データラベル（金額のみ表示）
        With .SeriesCollection(1).DataLabels
        End With
        .SeriesCollection(1).HasDataLabels = True
        With .SeriesCollection(1).DataLabels
            .NumberFormat = "#,##0"
            .Position = xlLabelPositionOutsideEnd
            .Font.Size = 8
        End With

        ' プロットエリア余白を調整
        .PlotArea.Interior.Color = RGB(248, 248, 248)
    End With

    LogMessage "グラフを作成しました（" & count & "製品）"
    MsgBox "グラフを作成しました。(" & count & "製品)", vbInformation, "完了"
End Sub
