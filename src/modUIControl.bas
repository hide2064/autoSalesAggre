Attribute VB_Name = "modUIControl"
Option Explicit

' ============================================================
' modUIControl — UI制御・処理オーケストレーションモジュール
'
' 役割:
'   ・main シートの「ファイルを読み込む」ボタンから呼ばれる RunAll を提供する。
'     RunAll は各モジュールの処理を正しい順序で呼び出すオーケストレーターとして機能する。
'   ・全モジュールから呼び出せるログ書き込み関数 LogMessage を提供する。
'
' RunAll の処理順序:
'   1. マスタ読み込み (modConfig)
'   2. ファイル選択 (modFileIO)
'   3. TSV → シート読み込み (modFileIO)
'   4. all シート構築 (modDataProcess)
'   5. 部署リスト更新 (modConfig)
'   6. 集計再描画 (modAggregation)
'
' Application.EnableEvents の制御:
'   ・RunAll 開始時に False にして集計シートの Worksheet_Change が
'     途中で発火しないようにする。
'   ・Rebuild 呼び出し前に True に戻す。これにより Rebuild 内での
'     セル書き込みに対して Worksheet_Change が発火しても問題ない。
'   ・エラー発生時も Cleanup ラベルで必ず True に戻す。
' ============================================================

' ============================================================
' RunAll — ファイル読み込みから集計完了までの全処理を実行する
'
' main シートの「ファイルを読み込む」ボタンの OnAction に設定されている。
' エラーが発生した場合は ErrHandler でメッセージを表示し、
' Cleanup で Application の状態を必ず復元してから終了する。
' ============================================================
Public Sub RunAll()
    Dim dictProduct    As Object  ' 製品マスタ辞書
    Dim dictCommission As Object  ' 口銭マスタ辞書
    Dim dictHeaderMap  As Object  ' ヘッダー名寄せ辞書
    Dim files As Variant          ' 選択ファイルパス配列（キャンセル時は False）
    Dim i As Integer
    Dim successCount As Integer   ' 正常読み込みファイル数
    Dim dictDept As Object        ' 部署名一覧辞書

    On Error GoTo ErrHandler

    ' --- Application 状態の設定（パフォーマンス向上・誤動作防止）---
    Application.ScreenUpdating = False          ' 画面更新を停止（高速化）
    Application.Calculation   = xlCalculationManual  ' 自動計算を停止（高速化）
    Application.EnableEvents  = False           ' イベント発火を抑制（Worksheet_Change の誤動作防止）

    LogMessage "===== 処理開始 ====="

    ' --- マスタ読み込み ---
    LogMessage "マスタ読み込み中..."
    Set dictProduct    = LoadProductDict()
    Set dictCommission = LoadCommissionDict()
    Set dictHeaderMap  = LoadHeaderMap()
    LogMessage "  製品マスタ: " & dictProduct.Count & "件 / 口銭マスタ: " & dictCommission.Count & "件 / 名寄せ: " & dictHeaderMap.Count & "エントリ"

    ' --- ファイル選択ダイアログ ---
    files = SelectFiles()
    If VarType(files) = vbBoolean Then
        ' ユーザーがキャンセルした場合は Cleanup へ（エラーではない）
        LogMessage "ファイル選択がキャンセルされました"
        GoTo Cleanup
    End If

    LogMessage CStr(UBound(files)) & "件のファイルを読み込みます"

    ' --- 選択されたファイルを順次読み込む ---
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

    ' --- all シートの構築 ---
    LogMessage "allシート構築中..."
    BuildAllSheet dictProduct, dictCommission, dictHeaderMap
    LogMessage "allシート構築完了"

    ' --- 部署リストの更新（Config!J列 と 集計!B1 ドロップダウン）---
    Set dictDept = CollectUniqueDepts()
    RefreshDeptList dictDept
    LogMessage "部署リスト更新完了 (" & dictDept.Count & "部署)"

    ' --- 集計再描画 ---
    ' EnableEvents を True に戻してから Rebuild を呼ぶ。
    ' これにより Rebuild 内のセル書き込みで Worksheet_Change が発火しても
    ' 正常動作する（ただし RunAll 中は既に all シートが構築済みなので問題なし）。
    Application.EnableEvents = True
    Rebuild
    LogMessage "集計完了"

    LogMessage "===== 処理完了 ====="

Cleanup:
    ' Application の状態を必ず元に戻す（エラー・キャンセル時も通過する）
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
' LogMessage — main シートにタイムスタンプ付きのログを書き込む
'
' 引数:
'   msg — 記録するメッセージ文字列
'
' Public で宣言することで modDataProcess など他のモジュールから
' モジュール名を指定せずに直接呼び出せる。
'
' 書き込み先: main シートの A列(日時) / B列(メッセージ)
'             MAIN_LOG_START_ROW(3行目) 以降に追記する。
' ============================================================
Public Sub LogMessage(msg As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Sheets(SH_MAIN)

    ' A列の最終入力行の次の行に書き込む
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < MAIN_LOG_START_ROW Then nextRow = MAIN_LOG_START_ROW  ' 最小行を保証

    ws.Cells(nextRow, 1).Value         = Now()
    ws.Cells(nextRow, 1).NumberFormat  = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value         = msg
End Sub
