Attribute VB_Name = "modError"
Option Explicit

' ============================================================
' modError — エラーレポートシート管理モジュール
'
' 役割:
'   ・データ処理中に発生した警告・エラーを "エラー" シートに追記する
'     LogError を提供する。
'   ・RunAll 開始時に呼ぶ ClearErrorSheet でシートをリセットする。
'   ・処理後に GetErrorCount でエラー件数を確認できる。
'
' エラーシートのレイアウト:
'   A列: タイムスタンプ
'   B列: ソースファイル名
'   C列: 行番号（ソースシート内の行番号）
'   D列: エラー種別（例: "製品コード未登録", "重複行"）
'   E列: 詳細メッセージ
'   F列: 問題の値
'
' エラー行はすべて CLR_ERROR_ROW（薄赤）で色付けされる。
' ============================================================

' --- エラーシートの列インデックス ---
Private Const ERR_COL_TIMESTAMP As Integer = 1  ' A: タイムスタンプ
Private Const ERR_COL_SOURCE    As Integer = 2  ' B: ソースファイル名
Private Const ERR_COL_ROW       As Integer = 3  ' C: 行番号
Private Const ERR_COL_TYPE      As Integer = 4  ' D: エラー種別
Private Const ERR_COL_DETAIL    As Integer = 5  ' E: 詳細メッセージ
Private Const ERR_COL_VALUE     As Integer = 6  ' F: 問題の値

' エラーシートのデータ開始行（1行目=タイトル, 2行目=ヘッダー）
Public Const ERR_DATA_ROW As Integer = 3

' ============================================================
' ClearErrorSheet — エラーシートのデータ行をクリアする
'
' RunAll 開始時に呼び出し、前回実行時のエラー記録を消去する。
' タイトル行(1行目)とヘッダー行(2行目)は保持する。
' ============================================================
Public Sub ClearErrorSheet()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets(SH_ERROR)
    lastRow = ws.Cells(ws.Rows.Count, ERR_COL_TIMESTAMP).End(xlUp).Row

    If lastRow >= ERR_DATA_ROW Then
        ws.Rows(ERR_DATA_ROW & ":" & lastRow).Delete
    End If
End Sub

' ============================================================
' LogError — エラー1件をエラーシートの末尾に追記する
'
' 引数:
'   sourceFile — エラーが発生したソースファイル名（シート名）
'   rowNum     — ソースシート内の行番号
'   errType    — エラー種別（例: "製品コード未登録", "重複行"）
'   detail     — 詳細メッセージ
'   value      — 問題となった値
'
' 追記した行は CLR_ERROR_ROW（薄赤背景）で色付けする。
' ============================================================
Public Sub LogError(sourceFile As String, rowNum As Long, _
                    errType As String, detail As String, value As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Sheets(SH_ERROR)
    nextRow = ws.Cells(ws.Rows.Count, ERR_COL_TIMESTAMP).End(xlUp).Row + 1
    If nextRow < ERR_DATA_ROW Then nextRow = ERR_DATA_ROW

    ws.Cells(nextRow, ERR_COL_TIMESTAMP).Value        = Now()
    ws.Cells(nextRow, ERR_COL_TIMESTAMP).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(nextRow, ERR_COL_SOURCE).Value            = sourceFile
    ws.Cells(nextRow, ERR_COL_ROW).Value               = rowNum
    ws.Cells(nextRow, ERR_COL_TYPE).Value              = errType
    ws.Cells(nextRow, ERR_COL_DETAIL).Value            = detail
    ws.Cells(nextRow, ERR_COL_VALUE).Value             = value

    ' エラー行を薄赤で色付け（視認性向上）
    ws.Rows(nextRow).Interior.Color = CLR_ERROR_ROW
End Sub

' ============================================================
' GetErrorCount — 現在のエラー件数を返す
'
' 戻り値: エラーシートに記録されているエラーの件数（0 = エラーなし）
' ============================================================
Public Function GetErrorCount() As Long
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets(SH_ERROR)
    lastRow = ws.Cells(ws.Rows.Count, ERR_COL_TIMESTAMP).End(xlUp).Row
    GetErrorCount = IIf(lastRow < ERR_DATA_ROW, 0, lastRow - ERR_DATA_ROW + 1)
End Function

' ============================================================
' ActivateErrorSheet — エラーシートをアクティブにする（ボタン用）
'
' main シートの「エラー確認」ボタンに接続する想定。
' ============================================================
Public Sub ActivateErrorSheet()
    ThisWorkbook.Sheets(SH_ERROR).Activate
End Sub
