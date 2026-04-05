Option Explicit

' ============================================================
' run_automate.vbs — autoSalesAggre の自動実行スクリプト
'
' 使い方:
'   cscript setup\run_automate.vbs <フォルダパス>
'
' 引数:
'   <フォルダパス> — 処理対象の TSV/CSV/Excel ファイルが置かれたフォルダのパス
'
' 動作概要:
'   1. autoSalesAggre.xlsm をバックグラウンドで開く
'   2. modUIControl.RunAllHeadless を呼び出して全ファイルを自動処理する
'      （ファイル選択ダイアログを表示しない、MsgBox を表示しない）
'   3. ワークブックを保存して Excel を終了する
'
' 使用例 (バッチファイルやタスクスケジューラから呼び出す):
'   cscript "C:\Projects\autoSalesAggre\setup\run_automate.vbs" "C:\Data\Sales\"
'
' 注意:
'   ・autoSalesAggre.xlsm は既に存在している必要があります
'     （存在しない場合は先に create_workbook.vbs を実行してください）
'   ・Excel の「マクロを有効にする」設定が必要です
' ============================================================

' --- 引数チェック ---
If WScript.Arguments.Count < 1 Then
    WScript.Echo "使い方: cscript run_automate.vbs <フォルダパス>"
    WScript.Echo "例   : cscript run_automate.vbs ""C:\Data\Sales\"""
    WScript.Quit 1
End If

Dim folderPath
folderPath = WScript.Arguments(0)

' --- パス解決 ---
Dim fso, scriptDir, xlsmPath
Set fso       = CreateObject("Scripting.FileSystemObject")
scriptDir     = fso.GetParentFolderName(WScript.ScriptFullName)
xlsmPath      = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\autoSalesAggre.xlsm"))

' --- 事前チェック ---
If Not fso.FolderExists(folderPath) Then
    WScript.Echo "[ERROR] フォルダが見つかりません: " & folderPath
    WScript.Quit 1
End If

If Not fso.FileExists(xlsmPath) Then
    WScript.Echo "[ERROR] autoSalesAggre.xlsm が見つかりません: " & xlsmPath
    WScript.Echo "先に cscript create_workbook.vbs を実行してワークブックを生成してください。"
    WScript.Quit 1
End If

WScript.Echo "autoSalesAggre 自動実行"
WScript.Echo "対象フォルダ: " & folderPath
WScript.Echo "ワークブック: " & xlsmPath

' ============================================================
' Excel 操作ブロック
' ============================================================
Dim xlApp, wb
Set xlApp = Nothing

On Error Resume Next

' --- Excel をバックグラウンドで起動 ---
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "[ERROR] Excel の起動に失敗しました: " & Err.Description
    WScript.Quit 1
End If

xlApp.Visible       = False
xlApp.DisplayAlerts = False

' --- autoSalesAggre.xlsm を開く ---
Set wb = xlApp.Workbooks.Open(xlsmPath, False, False)
If Err.Number <> 0 Then
    WScript.Echo "[ERROR] ワークブックのオープンに失敗しました: " & Err.Description
    xlApp.Quit
    WScript.Quit 1
End If
WScript.Echo "ワークブックを開きました"

' --- RunAllHeadless を実行（フォルダパスを渡す）---
WScript.Echo "RunAllHeadless を実行中..."
xlApp.Run "modUIControl.RunAllHeadless", folderPath
If Err.Number <> 0 Then
    WScript.Echo "[ERROR] RunAllHeadless の実行に失敗しました: " & Err.Description
    wb.Close False
    xlApp.Quit
    WScript.Quit 1
End If
WScript.Echo "RunAllHeadless 完了"

' --- 保存・終了 ---
' RunAllHeadless 内でも ThisWorkbook.Save を呼んでいるが念のため再保存
wb.Save
If Err.Number <> 0 Then
    WScript.Echo "[警告] 保存に失敗しました: " & Err.Description
End If

On Error GoTo 0
wb.Close False
xlApp.Quit
Set xlApp = Nothing

WScript.Echo ""
WScript.Echo "処理が完了しました。ログは autoSalesAggre.xlsm の main シートを確認してください。"
