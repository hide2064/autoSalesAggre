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
'   有効でない場合、VBProject へのアクセスが拒否されて以下のエラーで失敗する:
'     「プログラムによるVisual Basicプロジェクトへのアクセスは信頼性に欠けます」
'
' 処理概要:
'   1. src\ フォルダの存在確認
'   2. Excel をバックグラウンドで起動する
'   3. 新規ワークブックに 5 シート (main, Config, all, Shuukei, Pivot) を作成する
'   4. src\ 配下の .bas モジュールを VBProject にインポートする
'   5. modSetup.InitWorkbook を実行してシートのレイアウト・ボタン・
'      イベントコードを設定する（Shuukei → "集計"、Pivot → "ピボット" リネームも含む）
'   6. .xlsm 形式で上書き保存する
'
' 出力ファイル:
'   autoSalesAggre.xlsm（スクリプトの1つ上のディレクトリに生成）
'
' 再ビルド手順:
'   autoSalesAggre.xlsm を削除してからこのスクリプトを再実行する。
' ============================================================

' --- パス解決 ---
' WScript.ScriptFullName: このスクリプト自身の絶対パス
' scriptDir             : setup\ フォルダのパス
' srcPath               : src\ フォルダのパス（モジュールインポート元）
' outputFile            : 生成する .xlsm ファイルの絶対パス
Dim fso, scriptDir, srcPath, outputFile
Set fso       = CreateObject("Scripting.FileSystemObject")
scriptDir     = fso.GetParentFolderName(WScript.ScriptFullName)
srcPath       = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\src")) & "\"
outputFile    = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\autoSalesAggre.xlsm"))

WScript.Echo "Setup started"
WScript.Echo "Source : " & srcPath
WScript.Echo "Output : " & outputFile

' --- src\ フォルダの存在確認 ---
' モジュールが存在しない場合は早期に終了する（後続の Excel 起動を無駄にしない）
If Not fso.FolderExists(srcPath) Then
    WScript.Echo "[ERROR] src フォルダが見つかりません: " & srcPath
    WScript.Echo "スクリプトを setup\ フォルダから実行しているか確認してください。"
    WScript.Quit 1
End If

' インポートする .bas ファイルの一覧（不足している場合に早期検出）
Dim requiredModules(8)
requiredModules(0) = "modConfig.bas"
requiredModules(1) = "modFileIO.bas"
requiredModules(2) = "modDataProcess.bas"
requiredModules(3) = "modAggregation.bas"
requiredModules(4) = "modUIControl.bas"
requiredModules(5) = "modSharePoint.bas"
requiredModules(6) = "modChart.bas"
requiredModules(7) = "modPivot.bas"
requiredModules(8) = "modSetup.bas"

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

' --- Excel をバックグラウンドで起動 ---
' Visible = False で画面に表示しない（処理中に誤操作されるのを防ぐ）
' DisplayAlerts = False で確認ダイアログを抑制する
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "[ERROR] Excel の起動に失敗しました: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0  ' エラートラップをリセット（Excel 起動成功後は通常モードで続行）

xlApp.Visible       = False
xlApp.DisplayAlerts = False

' 以降は On Error Resume Next で囲み、エラー時に xlApp.Quit を保証する
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
' 新規ワークブックはデフォルトで Sheet1〜Sheet3 などが存在するため、
' まず Sheet1 だけ残して残りを削除し、
' それを "main" にリネームしてから必要なシートを追加する。
'
' 最終的なシート順序: main(1), Config(2), all(3), Shuukei(4), Pivot(5)
' ※ Shuukei は modSetup.InitWorkbook 内で "集計" にリネームされる
' ※ Pivot   は modSetup.InitWorkbook 内で "ピボット" にリネームされる
' ============================================================
Do While wb.Sheets.Count > 1
    wb.Sheets(wb.Sheets.Count).Delete
Loop
wb.Sheets(1).Name = "main"

Dim shConfig, shAll, shShuukei, shPivot

' Add() はデフォルトで既存の最後のシートの前に挿入されるため、
' 追加後に Move で正しい順序に並べ替える
Set shConfig  = wb.Sheets.Add()
shConfig.Name = "Config"
Set shAll     = wb.Sheets.Add()
shAll.Name    = "all"
Set shShuukei = wb.Sheets.Add()
shShuukei.Name = "Shuukei"  ' modSetup.InitWorkbook で "集計" にリネームされる
Set shPivot   = wb.Sheets.Add()
shPivot.Name  = "Pivot"     ' modSetup.InitWorkbook で "ピボット" にリネームされる

If Err.Number <> 0 Then
    WScript.Echo "[ERROR] シートの作成に失敗しました: " & Err.Description
    wb.Close False
    xlApp.Quit
    WScript.Quit 1
End If

' 明示的に順序を確定する（Add の挿入位置がバージョンによって異なる場合の保険）
wb.Sheets("main").Move    wb.Sheets(1)
wb.Sheets("Config").Move  wb.Sheets(2)
wb.Sheets("all").Move     wb.Sheets(3)
wb.Sheets("Shuukei").Move wb.Sheets(4)
wb.Sheets("Pivot").Move   wb.Sheets(5)

' ============================================================
' VBA モジュールのインポート
'
' インポート順序の注意:
'   modSetup は modConfig/modFileIO 等の定数・関数を参照するため、
'   modSetup より先に依存モジュールをインポートしておく必要がある。
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
        WScript.Echo "[ERROR] モジュールのインポートに失敗しました (" & requiredModules(i) & "): " & Err.Description
        wb.Close False
        xlApp.Quit
        WScript.Quit 1
    End If
    WScript.Echo "  Imported: " & requiredModules(i)
Next
WScript.Echo "Modules imported."

' ============================================================
' ワークブック初期化の実行
'
' InitWorkbook が行う処理:
'   ・Shuukei → "集計"、Pivot → "ピボット" にリネーム
'   ・各シートのレイアウト設定（ヘッダー・ボタン・列幅など）
'   ・集計シートへの Worksheet_Change イベントコードの動的注入
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

' --- .xlsm 形式で保存 ---
' FileFormat 52 = xlOpenXMLMacroEnabled (.xlsm)
' マクロを含むため .xlsx ではなく .xlsm を使用する
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
