Option Explicit

' ============================================================
' create_workbook.vbs — autoSalesAggre.xlsm 生成スクリプト
'
' 使い方:
'   cscript setup\create_workbook.vbs
'
' 前提条件:
'   Excel の「ファイル → オプション → トラストセンター → トラストセンターの設定
'   → マクロの設定 → VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」
'   を有効にしておくこと。
'
' 処理概要:
'   1. src\ フォルダとモジュールファイルの存在確認
'   2. Excel をバックグラウンドで起動する
'   3. 新規ワークブックに 7 シートを作成する:
'      main / Config / all / Shuukei(→集計) / Pivot(→ピボット)
'      Error(→エラー) / Monthly(→月次サマリー)
'   4. src\ 配下の .bas モジュールをインポートする
'   5. modSetup.InitWorkbook を実行してシートレイアウトを設定する
'   6. .xlsm 形式で保存する
'
' 再ビルド手順:
'   autoSalesAggre.xlsm を削除してからこのスクリプトを再実行する。
' ============================================================

' --- パス解決 ---
Dim fso, scriptDir, srcPath, outputFile
Set fso       = CreateObject("Scripting.FileSystemObject")
scriptDir     = fso.GetParentFolderName(WScript.ScriptFullName)
srcPath       = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\src")) & "\"
outputFile    = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\autoSalesAggre.xlsm"))

WScript.Echo "Setup started"
WScript.Echo "Source : " & srcPath
WScript.Echo "Output : " & outputFile

' --- src\ フォルダの存在確認 ---
If Not fso.FolderExists(srcPath) Then
    WScript.Echo "[ERROR] src フォルダが見つかりません: " & srcPath
    WScript.Quit 1
End If

' --- インポートするモジュールファイルの一覧（インポート順序=依存順）---
Dim requiredModules(11)
requiredModules(0)  = "modConfig.bas"
requiredModules(1)  = "modFileIO.bas"
requiredModules(2)  = "modDataProcess.bas"
requiredModules(3)  = "modAggregation.bas"
requiredModules(4)  = "modUIControl.bas"
requiredModules(5)  = "modSharePoint.bas"
requiredModules(6)  = "modChart.bas"
requiredModules(7)  = "modPivot.bas"
requiredModules(8)  = "modError.bas"
requiredModules(9)  = "modExport.bas"
requiredModules(10) = "modMonthly.bas"
requiredModules(11) = "modSetup.bas"  ' 最後: 他モジュールに依存するため

Dim i
For i = 0 To UBound(requiredModules)
    If Not fso.FileExists(srcPath & requiredModules(i)) Then
        WScript.Echo "[ERROR] モジュールファイルが見つかりません: " & srcPath & requiredModules(i)
        WScript.Quit 1
    End If
Next

' ============================================================
' Excel 操作ブロック
'
' On Error Resume Next を使い、エラー発生時でも必ず xlApp.Quit を
' 呼び出してバックグラウンド Excel プロセスを残さないようにする。
' ============================================================
Dim xlApp, wb
Set xlApp = Nothing

On Error Resume Next

Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "[ERROR] Excel の起動に失敗しました: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

xlApp.Visible       = False
xlApp.DisplayAlerts = False

On Error Resume Next

' --- 新規ワークブックを作成 ---
Set wb = xlApp.Workbooks.Add
If Err.Number <> 0 Then
    WScript.Echo "[ERROR] ワークブックの作成に失敗しました: " & Err.Description
    xlApp.Quit
    WScript.Quit 1
End If

' ============================================================
' シートの作成と順序の整理
'
' 最終的なシート順序:
'   main(1) / Config(2) / all(3) / Shuukei(4→集計) /
'   Pivot(5→ピボット) / Error(6→エラー) / Monthly(7→月次サマリー)
' ============================================================
Do While wb.Sheets.Count > 1
    wb.Sheets(wb.Sheets.Count).Delete
Loop
wb.Sheets(1).Name = "main"

Dim shConfig, shAll, shShuukei, shPivot, shError, shMonthly

Set shConfig  = wb.Sheets.Add() : shConfig.Name  = "Config"
Set shAll     = wb.Sheets.Add() : shAll.Name     = "all"
Set shShuukei = wb.Sheets.Add() : shShuukei.Name = "Shuukei"
Set shPivot   = wb.Sheets.Add() : shPivot.Name   = "Pivot"
Set shError   = wb.Sheets.Add() : shError.Name   = "Error"
Set shMonthly = wb.Sheets.Add() : shMonthly.Name = "Monthly"

If Err.Number <> 0 Then
    WScript.Echo "[ERROR] シートの作成に失敗しました: " & Err.Description
    wb.Close False
    xlApp.Quit
    WScript.Quit 1
End If

' 明示的に順序を確定する
wb.Sheets("main").Move    wb.Sheets(1)
wb.Sheets("Config").Move  wb.Sheets(2)
wb.Sheets("all").Move     wb.Sheets(3)
wb.Sheets("Shuukei").Move wb.Sheets(4)
wb.Sheets("Pivot").Move   wb.Sheets(5)
wb.Sheets("Error").Move   wb.Sheets(6)
wb.Sheets("Monthly").Move wb.Sheets(7)

' ============================================================
' VBA モジュールのインポート
' ============================================================
WScript.Echo "Importing VBA modules..."
Dim vbp
Set vbp = wb.VBProject

If Err.Number <> 0 Then
    WScript.Echo "[ERROR] VBProject へのアクセスに失敗しました: " & Err.Description
    WScript.Echo "Excel の「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」を"
    WScript.Echo "有効にしているか確認してください。"
    wb.Close False
    xlApp.Quit
    WScript.Quit 1
End If

For i = 0 To UBound(requiredModules)
    vbp.VBComponents.Import srcPath & requiredModules(i)
    If Err.Number <> 0 Then
        WScript.Echo "[ERROR] モジュールのインポートに失敗しました (" & _
                     requiredModules(i) & "): " & Err.Description
        wb.Close False
        xlApp.Quit
        WScript.Quit 1
    End If
    WScript.Echo "  Imported: " & requiredModules(i)
Next
WScript.Echo "Modules imported."

' ============================================================
' ワークブック初期化の実行
' ============================================================
WScript.Echo "Running InitWorkbook..."
xlApp.Run "modSetup.InitWorkbook"
If Err.Number <> 0 Then
    WScript.Echo "[ERROR] InitWorkbook の実行に失敗しました: " & Err.Description
    wb.Close False
    xlApp.Quit
    WScript.Quit 1
End If
WScript.Echo "InitWorkbook complete."

' --- .xlsm 形式 (FileFormat 52) で保存 ---
wb.SaveAs outputFile, 52
If Err.Number <> 0 Then
    WScript.Echo "[ERROR] ファイルの保存に失敗しました: " & Err.Description
    wb.Close False
    xlApp.Quit
    WScript.Quit 1
End If
WScript.Echo "Saved: " & outputFile

On Error GoTo 0
xlApp.Quit
Set xlApp = Nothing

WScript.Echo ""
WScript.Echo "Done. Open autoSalesAggre.xlsm in Excel and enable macros."
