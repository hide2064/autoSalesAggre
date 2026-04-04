Attribute VB_Name = "modUIControl"
Option Explicit

Public Sub RunAll()
    Dim dictProduct As Object
    Dim dictCommission As Object
    Dim dictHeaderMap As Object
    Dim files As Variant
    Dim i As Integer
    Dim successCount As Integer
    Dim dictDept As Object

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    LogMessage "===== 処理開始 ====="

    ' Load masters
    LogMessage "マスタ読み込み中..."
    Set dictProduct = LoadProductDict()
    Set dictCommission = LoadCommissionDict()
    Set dictHeaderMap = LoadHeaderMap()
    LogMessage "  製品マスタ: " & dictProduct.Count & "件 / 口銭マスタ: " & dictCommission.Count & "件 / 名寄せ: " & dictHeaderMap.Count & "エントリ"

    ' Select files
    files = SelectFiles()
    If VarType(files) = vbBoolean Then
        LogMessage "ファイル選択がキャンセルされました"
        GoTo Cleanup
    End If

    LogMessage CStr(UBound(files)) & "件のファイルを読み込みます"

    ' Load each file
    successCount = 0
    For i = 1 To UBound(files)
        LogMessage "  読込: " & files(i)
        If LoadTsvToSheet(CStr(files(i))) Then
            successCount = successCount + 1
        Else
            LogMessage "  [エラー] 読み込み失敗: " & files(i)
        End If
    Next i
    LogMessage successCount & "件のファイルを読み込みました"

    ' Build all sheet
    LogMessage "allシート構築中..."
    BuildAllSheet dictProduct, dictCommission, dictHeaderMap
    LogMessage "allシート構築完了"

    ' Refresh dept list
    Set dictDept = CollectUniqueDepts()
    RefreshDeptList dictDept
    LogMessage "部署リスト更新完了 (" & dictDept.Count & "部署)"

    ' Rebuild aggregation (re-enable events first so Worksheet_Change fires correctly)
    Application.EnableEvents = True
    Rebuild
    LogMessage "集計完了"

    LogMessage "===== 処理完了 ====="

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    LogMessage "[エラー] " & Err.Description
    MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "エラー"
    Resume Cleanup
End Sub

Public Sub LogMessage(msg As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Sheets(SH_MAIN)
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < MAIN_LOG_START_ROW Then nextRow = MAIN_LOG_START_ROW

    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value = msg
End Sub
