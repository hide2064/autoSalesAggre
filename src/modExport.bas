Attribute VB_Name = "modExport"
Option Explicit

' ============================================================
' modExport — エクスポートモジュール
'
' 役割:
'   ・集計シートの内容を独立した .xlsx ファイルとして書き出す
'     ExportAggrToFile を提供する。
'   ・集計シートの「エクスポート」ボタンに接続される。
'
' エクスポートの仕組み:
'   wsAggr.Copy を呼ぶと Excel が新規ワークブックにシートを複製する。
'   その新規ワークブックを .xlsx 形式で保存して閉じる。
'   データとセル書式（色・太字・罫線・数値書式）がそのまま複製される。
' ============================================================

' ============================================================
' ExportAggrToFile — 集計シートを新規 Excel ファイルとして保存する
'
' 処理概要:
'   1. 集計データの有無を確認
'   2. 名前を付けて保存ダイアログを表示
'   3. wsAggr.Copy で新規ワークブックに集計シートを複製
'   4. .xlsx 形式（FileFormat=51）で保存して閉じる
' ============================================================
Public Sub ExportAggrToFile()
    Dim wsAggr As Worksheet
    Dim newWb As Workbook
    Dim savePath As Variant
    Dim defaultName As String
    Dim lastRow As Long

    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)
    lastRow = wsAggr.Cells(wsAggr.Rows.Count, 1).End(xlUp).Row

    If lastRow < AGGR_DATA_ROW Then
        MsgBox "集計データがありません。先にデータを集計してください。", _
               vbExclamation, "データなし"
        Exit Sub
    End If

    ' デフォルトのファイル名: 集計_YYYYMMDD.xlsx
    defaultName = "集計_" & Format(Now(), "yyyymmdd") & ".xlsx"

    ' --- 保存先をダイアログで取得 ---
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:="Excelファイル (*.xlsx), *.xlsx", _
        Title:="集計結果の保存先を指定してください")

    ' キャンセルされた場合は False が返る
    If VarType(savePath) = vbBoolean Then Exit Sub

    On Error GoTo ErrHandler
    Set newWb = Nothing

    ' --- 集計シートを新規ワークブックに複製 ---
    ' Copy はシートをコピーして新規ワークブックを ActiveWorkbook として返す
    wsAggr.Copy
    Set newWb = ActiveWorkbook

    ' コピー先シートを "集計" に名前変更（Excelが自動付与する名前を上書き）
    On Error Resume Next
    newWb.Sheets(1).Name = SH_AGGR
    On Error GoTo ErrHandler

    ' --- .xlsx 形式(51 = xlOpenXMLWorkbook)で保存 ---
    Application.DisplayAlerts = False
    newWb.SaveAs Filename:=CStr(savePath), FileFormat:=51
    Application.DisplayAlerts = True

    newWb.Close SaveChanges:=False
    Set newWb = Nothing

    LogMessage "集計シートをエクスポートしました: " & CStr(savePath)
    MsgBox "エクスポートが完了しました。" & vbCrLf & CStr(savePath), _
           vbInformation, "エクスポート完了"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    If Not newWb Is Nothing Then
        On Error Resume Next
        newWb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "エクスポート中にエラーが発生しました:" & vbCrLf & Err.Description, _
           vbCritical, "エラー"
End Sub
