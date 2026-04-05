Attribute VB_Name = "modUIControl"
Option Explicit

' ============================================================
' modUIControl — UI制御・処理オーケストレーションモジュール
'
' 役割:
'   ・main シートの「ファイルを読み込む」ボタンから呼ばれる RunAll を提供する。
'   ・指定フォルダのファイルを自動処理する RunAllHeadless を提供する。
'   ・全モジュールから呼び出せるログ書き込み関数 LogMessage を提供する。
'
' RunAll の処理順序:
'   1. エラーシートのクリア (modError)
'   2. Config バリデーション (modConfig)
'   3. マスタ読み込み (modConfig)
'   4. ファイル選択 (modFileIO)
'   5. ファイル → シート読み込み (modFileIO)
'   6. all シート構築 (modDataProcess)
'   7. 部署リスト更新 (modConfig)
'   8. 集計再描画 (modAggregation)
'   9. ピボットテーブル更新 (modPivot)
'  10. 月次サマリー更新 (modMonthly)
'
' RunAllHeadless の処理順序:
'   引数のフォルダパスにある対応ファイルを列挙して RunAll 相当の処理を実行する。
'   ダイアログを一切表示しないためスクリプト (run_automate.vbs) から呼び出せる。
' ============================================================

' ============================================================
' RunAll — ファイル読み込みから集計完了までの全処理を実行する
' ============================================================
Public Sub RunAll()
    Dim dictProduct    As Object
    Dim dictCommission As Object
    Dim dictHeaderMap  As Object
    Dim files As Variant
    Dim i As Integer
    Dim successCount As Integer
    Dim dictDept As Object
    Dim issueCount As Integer

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False

    LogMessage "===== 処理開始 ====="

    ' --- エラーシートをクリア（前回実行分のエラーを消去）---
    ClearErrorSheet
    LogMessage "エラーシートをクリアしました"

    ' --- Config バリデーション ---
    LogMessage "Config シートを検証中..."
    Set dictProduct    = LoadProductDict()
    Set dictCommission = LoadCommissionDict()
    Set dictHeaderMap  = LoadHeaderMap()
    issueCount = ValidateConfig()
    If issueCount > 0 Then
        LogMessage "Config検証: " & issueCount & "件の警告があります（処理は継続します）"
    Else
        LogMessage "Config検証: 問題なし"
    End If
    LogMessage "  製品マスタ: " & dictProduct.Count & "件 / 口銭マスタ: " & _
               dictCommission.Count & "件 / 名寄せ: " & dictHeaderMap.Count & "エントリ"

    ' --- ファイル選択ダイアログ ---
    files = SelectFiles()
    If VarType(files) = vbBoolean Then
        LogMessage "ファイル選択がキャンセルされました"
        GoTo Cleanup
    End If

    ' --- 前回読み込んだファイルシートをすべて削除 ---
    ' 新しいファイル群だけがワークブックに残るようにするため、
    ' ファイル読み込みを開始する前に旧ソースシートを一括消去する。
    ClearSourceSheets
    LogMessage "旧ファイルシートをクリアしました"

    LogMessage CStr(UBound(files)) & "件のファイルを読み込みます"

    ' --- 選択されたファイルを順次読み込む（拡張子に応じてローダーを自動選択）---
    successCount = 0
    For i = 1 To UBound(files)
        LogMessage "  読込: " & files(i)
        If LoadFileToSheet(CStr(files(i))) Then
            successCount = successCount + 1
        Else
            LogMessage "  [エラー] 読み込み失敗: " & files(i)
        End If
    Next i
    LogMessage successCount & "件のファイルを読み込みました"

    ' --- all シートの構築 ---
    LogMessage "allシート構築中..."
    BuildAllSheet dictProduct, dictCommission, dictHeaderMap
    LogMessage "allシート構築完了"

    ' エラー件数のサマリーをログに出力
    Dim errCnt As Long
    errCnt = GetErrorCount()
    If errCnt > 0 Then
        LogMessage "  ※ " & errCnt & "件の警告/エラーが「エラー」シートに記録されました"
    End If

    ' --- 部署リストの更新 ---
    Set dictDept = CollectUniqueDepts()
    RefreshDeptList dictDept
    LogMessage "部署リスト更新完了 (" & dictDept.Count & "部署)"

    ' --- 集計再描画 ---
    Application.EnableEvents = True
    Rebuild
    LogMessage "集計完了"

    ' --- ピボットテーブル更新 ---
    BuildPivot
    LogMessage "ピボットテーブル更新完了"

    ' --- 月次サマリー更新 ---
    BuildMonthly
    LogMessage "月次サマリー更新完了"

    LogMessage "===== 処理完了 ====="

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Exit Sub

ErrHandler:
    LogMessage "[エラー] " & Err.Description
    MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "エラー"
    Resume Cleanup
End Sub

' ============================================================
' RunAllHeadless — 指定フォルダのファイルを自動処理する（ダイアログなし）
'
' 引数:
'   folderPath — 処理対象ファイルが置かれているフォルダのパス
'               （末尾の "\" は有無どちらでも可）
'
' 対応拡張子: .tsv .txt .csv .xlsx .xls .xlsm
'
' 呼び出し元: setup\run_automate.vbs から xlApp.Run で呼ばれる。
' ダイアログ・MsgBox を一切表示しないため自動実行に適する。
' 処理完了後にワークブックを上書き保存する。
' ============================================================
Public Sub RunAllHeadless(folderPath As String)
    Dim dictProduct    As Object
    Dim dictCommission As Object
    Dim dictHeaderMap  As Object
    Dim dictDept As Object
    Dim normalizedPath As String
    Dim exts(5) As String
    Dim ei As Integer
    Dim fname As String
    Dim fileList() As String
    Dim fileCount As Integer
    Dim i As Integer

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False
    Application.DisplayAlerts  = False

    LogMessage "===== 自動処理開始 [" & folderPath & "] ====="

    ClearErrorSheet

    ' --- 末尾の "\" を保証 ---
    normalizedPath = folderPath
    If Right(normalizedPath, 1) <> "\" Then normalizedPath = normalizedPath & "\"

    ' --- 対応拡張子のファイルを列挙 ---
    ' Dir() を拡張子ごとに順次呼び出してファイルリストを構築する
    exts(0) = "*.tsv" : exts(1) = "*.txt" : exts(2) = "*.csv"
    exts(3) = "*.xlsx" : exts(4) = "*.xls" : exts(5) = "*.xlsm"
    fileCount = 0
    ReDim fileList(0)

    For ei = 0 To 5
        fname = Dir(normalizedPath & exts(ei))
        Do While fname <> ""
            ReDim Preserve fileList(fileCount)
            fileList(fileCount) = normalizedPath & fname
            fileCount = fileCount + 1
            fname = Dir()
        Loop
    Next ei

    If fileCount = 0 Then
        LogMessage "対象ファイルが見つかりませんでした: " & normalizedPath
        GoTo Cleanup
    End If

    LogMessage fileCount & "件のファイルを処理します"

    ' --- 前回読み込んだファイルシートをすべて削除 ---
    ClearSourceSheets

    ' --- マスタ読み込み ---
    Set dictProduct    = LoadProductDict()
    Set dictCommission = LoadCommissionDict()
    Set dictHeaderMap  = LoadHeaderMap()
    ValidateConfig  ' 警告は LogMessage に出力（戻り値は無視）

    ' --- ファイルを読み込む ---
    For i = 0 To fileCount - 1
        LogMessage "  読込: " & fileList(i)
        If Not LoadFileToSheet(fileList(i)) Then
            LogMessage "  [エラー] 読み込み失敗: " & fileList(i)
        End If
    Next i

    ' --- all シート構築 ---
    BuildAllSheet dictProduct, dictCommission, dictHeaderMap

    ' --- 部署リスト・集計・ピボット・月次サマリー更新 ---
    Set dictDept = CollectUniqueDepts()
    RefreshDeptList dictDept
    Application.EnableEvents = True
    Rebuild
    BuildPivot
    BuildMonthly

    ' --- ワークブックを保存 ---
    ThisWorkbook.Save
    LogMessage "ワークブックを保存しました: " & ThisWorkbook.FullName

    LogMessage "===== 自動処理完了 ====="

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.DisplayAlerts  = True
    Exit Sub

ErrHandler:
    LogMessage "[エラー] RunAllHeadless: " & Err.Description
    Resume Cleanup
End Sub

' ============================================================
' LogMessage — main シートにタイムスタンプ付きのログを書き込む
'
' Public で宣言することで modDataProcess など他のモジュールから
' モジュール名を指定せずに直接呼び出せる。
' ============================================================
Public Sub LogMessage(msg As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Sheets(SH_MAIN)
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < MAIN_LOG_START_ROW Then nextRow = MAIN_LOG_START_ROW

    ws.Cells(nextRow, 1).Value        = Now()
    ws.Cells(nextRow, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value        = msg
End Sub
