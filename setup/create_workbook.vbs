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
'   1. Excel をバックグラウンドで起動する
'   2. 新規ワークブックに 4 シート (main, Config, all, Shuukei) を作成する
'   3. src\ 配下の .bas モジュールを VBProject にインポートする
'   4. modSetup.InitWorkbook を実行してシートのレイアウト・ボタン・
'      イベントコードを設定する（Shuukei を "集計" にリネームも含む）
'   5. .xlsm 形式で上書き保存する
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

' --- Excel をバックグラウンドで起動 ---
' Visible = False で画面に表示しない（処理中に誤操作されるのを防ぐ）
' DisplayAlerts = False で確認ダイアログを抑制する
Dim xlApp, wb
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible      = False
xlApp.DisplayAlerts = False

' --- 新規ワークブックを作成 ---
Set wb = xlApp.Workbooks.Add

' ============================================================
' シートの作成と順序の整理
'
' 新規ワークブックはデフォルトで Sheet1〜Sheet3 などが存在するため、
' まず Sheet1 だけ残して残りを削除し、
' それを "main" にリネームしてから必要なシートを追加する。
'
' 最終的なシート順序: main(1), Config(2), all(3), Shuukei(4)
' ※ Shuukei は modSetup.InitWorkbook 内で "集計" にリネームされる
' ============================================================
Do While wb.Sheets.Count > 1
    wb.Sheets(wb.Sheets.Count).Delete
Loop
wb.Sheets(1).Name = "main"

Dim shConfig, shAll, shShuukei

' Add() はデフォルトで既存の最後のシートの前に挿入されるため、
' 追加後に Move で正しい順序に並べ替える
Set shConfig  = wb.Sheets.Add()
shConfig.Name = "Config"
Set shAll     = wb.Sheets.Add()
shAll.Name    = "all"
Set shShuukei = wb.Sheets.Add()
shShuukei.Name = "Shuukei"  ' modSetup.InitWorkbook で "集計" にリネームされる

' 明示的に順序を確定する（Add の挿入位置がバージョンによって異なる場合の保険）
wb.Sheets("main").Move    wb.Sheets(1)
wb.Sheets("Config").Move  wb.Sheets(2)
wb.Sheets("all").Move     wb.Sheets(3)
wb.Sheets("Shuukei").Move wb.Sheets(4)

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

vbp.VBComponents.Import srcPath & "modConfig.bas"       ' 定数・マスタ読み込み
vbp.VBComponents.Import srcPath & "modFileIO.bas"       ' ファイル選択・TSV読み込み
vbp.VBComponents.Import srcPath & "modDataProcess.bas"  ' all シート構築・部署収集
vbp.VBComponents.Import srcPath & "modAggregation.bas"  ' 集計・グラフ描画ロジック
vbp.VBComponents.Import srcPath & "modUIControl.bas"    ' RunAll オーケストレーター・LogMessage
vbp.VBComponents.Import srcPath & "modSharePoint.bas"   ' SharePoint アップロード
vbp.VBComponents.Import srcPath & "modChart.bas"        ' グラフ作成
vbp.VBComponents.Import srcPath & "modSetup.bas"        ' 初期化（最後にインポート: 他モジュールに依存）
WScript.Echo "Modules imported."

' ============================================================
' ワークブック初期化の実行
'
' InitWorkbook が行う処理:
'   ・Shuukei シートを "集計" にリネーム
'   ・各シートのレイアウト設定（ヘッダー・ボタン・列幅など）
'   ・集計シートへの Worksheet_Change イベントコードの動的注入
' ============================================================
WScript.Echo "Running InitWorkbook..."
xlApp.Run "modSetup.InitWorkbook"
WScript.Echo "InitWorkbook complete."

' --- .xlsm 形式で保存 ---
' FileFormat 52 = xlOpenXMLMacroEnabled (.xlsm)
' マクロを含むため .xlsx ではなく .xlsm を使用する
wb.SaveAs outputFile, 52
WScript.Echo "Saved: " & outputFile

xlApp.Quit
Set xlApp = Nothing

WScript.Echo ""
WScript.Echo "Done. Open autoSalesAggre.xlsm in Excel and enable macros."
