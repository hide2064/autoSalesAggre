Attribute VB_Name = "modSetup"
Option Explicit

' ============================================================
' modSetup — ワークブック初期化モジュール
'
' 役割:
'   ・create_workbook.vbs から呼ばれる InitWorkbook を提供する。
'     InitWorkbook は各シートのレイアウト設定とイベントコードの注入を行う
'     一回限りの初期化処理をまとめたもの。
'   ・各シートのセットアップ処理をプライベート関数に分割して管理する。
'   ・集計シートの Worksheet_Change イベントを VBA コードとして
'     動的に注入する InjectAggrEvent を提供する。
'
' 前提条件:
'   Excel の「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」が
'   有効になっていること（InjectAggrEvent の VBProject 操作に必要）。
' ============================================================

' ============================================================
' InitWorkbook — ワークブック全体の初期化を実行する（Public）
'
' create_workbook.vbs から xlApp.Run "modSetup.InitWorkbook" で呼ばれる。
' ============================================================
Public Sub InitWorkbook()
    ' create_workbook.vbs で作成した仮名シートを正式名にリネーム
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case "Shuukei", "Sheet4", "Sheet3"
                ws.Name = SH_AGGR    ' → "集計"
            Case "Pivot"
                ws.Name = SH_PIVOT   ' → "ピボット"
            Case "Error"
                ws.Name = SH_ERROR   ' → "エラー"
            Case "Monthly"
                ws.Name = SH_MONTHLY ' → "月次サマリー"
        End Select
    Next ws

    SetupMainSheet
    SetupConfigSheet
    SetupAllSheet
    SetupAggrSheet
    SetupPivotSheet
    SetupErrorSheet
    SetupMonthlySheet
    InjectAggrEvent
End Sub

' ============================================================
' SetupMainSheet — main シートのレイアウトとボタンを設定する（プライベート）
' ============================================================
Private Sub SetupMainSheet()
    Dim ws As Worksheet
    Dim btn As Object

    Set ws = ThisWorkbook.Sheets(SH_MAIN)
    ws.Cells(1, 1).Value = "実行ログ"
    ws.Cells(2, 1).Value = "日時"
    ws.Cells(2, 2).Value = "メッセージ"
    ws.Cells(1, 1).Font.Bold = True
    With ws.Rows(2)
        .Font.Bold      = True
        .Interior.Color = CLR_HEADER_BG  ' modConfig 定数: RGB(200,220,240)
    End With
    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 80

    Set btn = ws.Buttons.Add(10, 10, 160, 30)
    btn.Caption  = "ファイルを読み込む"
    btn.OnAction = "modUIControl.RunAll"

    ' エラーシート確認ボタン
    Dim errBtn As Object
    Set errBtn = ws.Buttons.Add(180, 10, 130, 30)
    errBtn.Caption  = "エラーを確認する"
    errBtn.OnAction = "modError.ActivateErrorSheet"
End Sub

' ============================================================
' SetupConfigSheet — Config シートのレイアウトとサンプルデータを設定する（プライベート）
' ============================================================
Private Sub SetupConfigSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    ' --- 製品マスタ (A列・B列) ---
    ws.Cells(1, 1).Value = "製品マスタ"
    ws.Cells(2, 1).Value = "製品コード"
    ws.Cells(2, 2).Value = "製品名"

    ' --- 口銭マスタ (D列・E列) ---
    ws.Cells(1, 4).Value = "口銭マスタ"
    ws.Cells(2, 4).Value = "売上種別"
    ws.Cells(2, 5).Value = "口銭比率%"

    ' --- ヘッダー名寄せ設定 (G列・H列) ---
    ws.Cells(1, 7).Value = "ヘッダー名寄せ設定"
    ws.Cells(2, 7).Value = "正規名"
    ws.Cells(2, 8).Value = "対応列名（カンマ区切り）"

    ' --- 集計用部署リスト (J列) ---
    ws.Cells(1, 10).Value = "集計用部署リスト"
    ws.Cells(2, 10).Value = "全部署"

    ' --- セクションヘッダーを太字に ---
    ws.Cells(1, 1).Font.Bold  = True
    ws.Cells(1, 4).Font.Bold  = True
    ws.Cells(1, 7).Font.Bold  = True
    ws.Cells(1, 10).Font.Bold = True

    ' --- 列ヘッダーを太字に ---
    ws.Range("A2:B2").Font.Bold = True
    ws.Range("D2:E2").Font.Bold = True
    ws.Range("G2:H2").Font.Bold = True
    ws.Range("J2").Font.Bold    = True

    ' --- 列幅設定 ---
    ws.Columns("A:B").ColumnWidth = 16
    ws.Columns("D:E").ColumnWidth = 14
    ws.Columns("G:H").ColumnWidth = 20
    ws.Columns("J").ColumnWidth   = 16

    ' --- 製品マスタ サンプルデータ ---
    ws.Cells(3, 1).Value = "P001" : ws.Cells(3, 2).Value = "製品A"
    ws.Cells(4, 1).Value = "P002" : ws.Cells(4, 2).Value = "製品B"

    ' --- 口銭マスタ サンプルデータ ---
    ws.Cells(3, 4).Value = "直販"   : ws.Cells(3, 5).Value = 10
    ws.Cells(4, 4).Value = "代理店" : ws.Cells(4, 5).Value = 5

    ' --- ヘッダー名寄せ サンプルデータ (HDR_* 定数を使用) ---
    ws.Cells(3,  7).Value = HDR_CLIENT:     ws.Cells(3,  8).Value = "得意先名,得意先コード,顧客名"
    ws.Cells(4,  7).Value = HDR_PROD_CODE:  ws.Cells(4,  8).Value = "品番,ProductCode"
    ws.Cells(5,  7).Value = HDR_AMOUNT:     ws.Cells(5,  8).Value = "金額,Amount,売上高"
    ws.Cells(6,  7).Value = HDR_UNIT_PRICE: ws.Cells(6,  8).Value = "単価,定価"
    ws.Cells(7,  7).Value = HDR_QTY:        ws.Cells(7,  8).Value = "数量,Qty"
    ws.Cells(8,  7).Value = HDR_DATE:       ws.Cells(8,  8).Value = "日付,売上日,Date"
    ws.Cells(9,  7).Value = HDR_SALE_TYPE:  ws.Cells(9,  8).Value = "取引区分,SaleType"
    ws.Cells(10, 7).Value = HDR_DEPT:       ws.Cells(10, 8).Value = "部門,Dept"

    ' --- SharePoint連携 (L列・M列) ---
    ws.Cells(1, CFG_PA_LABEL_COL).Value = "SharePoint連携"
    ws.Cells(1, CFG_PA_LABEL_COL).Font.Bold = True
    ws.Cells(2, CFG_PA_LABEL_COL).Value = "集計シート送信URL"
    ws.Cells(2, CFG_PA_LABEL_COL).Font.Bold = True
    ws.Cells(3, CFG_PA_LABEL_COL).Value = "全データ送信URL"
    ws.Cells(3, CFG_PA_LABEL_COL).Font.Bold = True
    ws.Cells(3, CFG_PA_LABEL_COL).Font.Color = CLR_LABEL_TEXT  ' グレー（任意設定）
    ' M3 の説明（空欄時は M2 にフォールバック）
    ws.Cells(4, CFG_PA_LABEL_COL).Value = "※M3未設定時はM2を使用"
    ws.Cells(4, CFG_PA_LABEL_COL).Font.Color = CLR_LABEL_TEXT
    ws.Columns("L").ColumnWidth = 20
    ws.Columns("M").ColumnWidth = 60

    ' --- 保存済みフィルター条件 (O列) ---
    ws.Cells(1, CFG_SAVED_FILTER_COL).Value = "フィルター条件（保存）"
    ws.Cells(1, CFG_SAVED_FILTER_COL).Font.Bold = True
    ws.Cells(2, CFG_SAVED_FILTER_COL).Value = "部署"
    ws.Cells(3, CFG_SAVED_FILTER_COL).Value = "開始日"
    ws.Cells(4, CFG_SAVED_FILTER_COL).Value = "終了日"
    ws.Cells(1, CFG_SAVED_FILTER_COL).Font.Color = CLR_LABEL_TEXT
    ws.Range(ws.Cells(2, CFG_SAVED_FILTER_COL), _
             ws.Cells(4, CFG_SAVED_FILTER_COL)).Font.Color = CLR_LABEL_TEXT
    ws.Columns("O").ColumnWidth = 16
    ws.Columns("P").ColumnWidth = 20  ' 保存値の表示領域
End Sub

' ============================================================
' SetupAllSheet — all シートのヘッダー行とボタンを設定する（プライベート）
' ============================================================
Private Sub SetupAllSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ALL)

    ws.Cells(1, ALL_COL_CLIENT).Value     = HDR_CLIENT
    ws.Cells(1, ALL_COL_PROD_CODE).Value  = HDR_PROD_CODE
    ws.Cells(1, ALL_COL_AMOUNT).Value     = HDR_AMOUNT
    ws.Cells(1, ALL_COL_UNIT_PRICE).Value = HDR_UNIT_PRICE
    ws.Cells(1, ALL_COL_QTY).Value        = HDR_QTY
    ws.Cells(1, ALL_COL_DATE).Value       = HDR_DATE
    ws.Cells(1, ALL_COL_SALE_TYPE).Value  = HDR_SALE_TYPE
    ws.Cells(1, ALL_COL_DEPT).Value       = HDR_DEPT
    ws.Cells(1, ALL_COL_PROD_NAME).Value  = HDR_PROD_NAME
    ws.Cells(1, ALL_COL_MARGIN).Value     = HDR_MARGIN
    ws.Cells(1, ALL_COL_SOURCE).Value     = HDR_SOURCE

    With ws.Rows(1)
        .Font.Bold      = True
        .Interior.Color = CLR_HEADER_BG  ' modConfig 定数: RGB(200,220,240)
    End With
    ws.Columns("A:K").AutoFit

    Dim uploadBtn As Object
    Set uploadBtn = ws.Buttons.Add(700, 5, 180, 28)
    uploadBtn.Caption  = "SharePointへアップロード"
    uploadBtn.OnAction = "modSharePoint.UploadAllToSharePoint"
End Sub

' ============================================================
' SetupAggrSheet — 集計シートのレイアウトとボタンを設定する（プライベート）
'
' ボタン配置（左から順に）:
'   グラフ作成 (330,5,120,28) → modChart.DrawAggrChart
'   エクスポート (460,5,110,28) → modExport.ExportAggrToFile
'   SharePoint (580,5,160,28) → modSharePoint.UploadToSharePoint
'   条件を保存 (750,5,90,28) → modAggregation.SaveFilter
'   条件を復元 (850,5,90,28) → modAggregation.RestoreFilter
' ============================================================
Private Sub SetupAggrSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_AGGR)

    ' --- フィルタ入力欄 ---
    ws.Cells(1, 1).Value = "部署選択"
    ws.Cells(2, 1).Value = "開始日"
    ws.Cells(3, 1).Value = "終了日"
    ws.Range("A1:A3").Font.Bold    = True
    ws.Range(AGGR_DEPT_CELL).Value = "全部署"

    ' --- 集計テーブルのヘッダー行（5行目）---
    ws.Cells(AGGR_HDR_ROW, 2).Value = "売上金額合計"
    ws.Cells(AGGR_HDR_ROW, 3).Value = "売上数量合計"
    ws.Cells(AGGR_HDR_ROW, 4).Value = "口銭総額"
    With ws.Rows(AGGR_HDR_ROW)
        .Font.Bold      = True
        .Interior.Color = CLR_HEADER_BG  ' modConfig 定数: RGB(200,220,240)
    End With

    ws.Columns("A").ColumnWidth   = 30
    ws.Columns("B:D").ColumnWidth = 15

    ' --- グラフ作成ボタン ---
    Dim chartBtn As Object
    Set chartBtn = ws.Buttons.Add(330, 5, 120, 28)
    chartBtn.Caption  = "グラフ作成"
    chartBtn.OnAction = "modChart.DrawAggrChart"

    ' --- エクスポートボタン（集計結果を .xlsx として保存）---
    Dim exportBtn As Object
    Set exportBtn = ws.Buttons.Add(460, 5, 110, 28)
    exportBtn.Caption  = "Excelへ出力"
    exportBtn.OnAction = "modExport.ExportAggrToFile"

    ' --- SharePoint アップロードボタン ---
    Dim uploadBtn As Object
    Set uploadBtn = ws.Buttons.Add(580, 5, 160, 28)
    uploadBtn.Caption  = "SharePointへアップロード"
    uploadBtn.OnAction = "modSharePoint.UploadToSharePoint"

    ' --- フィルター条件を保存するボタン ---
    Dim saveBtn As Object
    Set saveBtn = ws.Buttons.Add(750, 5, 90, 28)
    saveBtn.Caption  = "条件を保存"
    saveBtn.OnAction = "modAggregation.SaveFilter"

    ' --- 保存したフィルター条件を復元するボタン ---
    Dim restoreBtn As Object
    Set restoreBtn = ws.Buttons.Add(850, 5, 90, 28)
    restoreBtn.Caption  = "条件を復元"
    restoreBtn.OnAction = "modAggregation.RestoreFilter"
End Sub

' ============================================================
' SetupPivotSheet — ピボットシートの UI 領域を設定する（プライベート）
' ============================================================
Private Sub SetupPivotSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PIVOT)

    ws.Cells(1, 1).Value     = "売上ピボットテーブル"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(2, 1).Value     = "RunAll 実行時に自動更新されます。" & _
                               "フィールドリストで行・列・フィルター・値を自由に配置できます。"
    ws.Cells(2, 1).Font.Color = CLR_LABEL_TEXT  ' modConfig 定数: RGB(100,100,100)
    ws.Columns("A").ColumnWidth = 35

    Dim btn As Object
    Set btn = ws.Buttons.Add(400, 5, 160, 28)
    btn.Caption  = "ピボットテーブル更新"
    btn.OnAction = "modPivot.BuildPivot"
End Sub

' ============================================================
' SetupErrorSheet — エラーレポートシートのレイアウトを設定する（プライベート）
'
' 設定内容:
'   1行目: タイトル "データ処理エラーレポート"
'   2行目: 列ヘッダー (タイムスタンプ / ソースファイル / 行番号 / エラー種別 / 詳細 / 値)
'
' データ行(3行目以降)は RunAll のたびにクリアして書き直される。
' ============================================================
Private Sub SetupErrorSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ERROR)

    ws.Cells(1, 1).Value     = "データ処理エラーレポート"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Color = CLR_ERROR_ROW  ' 薄赤で注意喚起

    ' ヘッダー行
    ws.Cells(2, 1).Value = "タイムスタンプ"
    ws.Cells(2, 2).Value = "ソースファイル"
    ws.Cells(2, 3).Value = "行番号"
    ws.Cells(2, 4).Value = "エラー種別"
    ws.Cells(2, 5).Value = "詳細メッセージ"
    ws.Cells(2, 6).Value = "問題の値"

    With ws.Rows(2)
        .Font.Bold      = True
        .Interior.Color = CLR_HEADER_BG  ' modConfig 定数: RGB(200,220,240)
    End With

    ws.Columns(1).ColumnWidth = 22  ' タイムスタンプ
    ws.Columns(2).ColumnWidth = 20  ' ソースファイル
    ws.Columns(3).ColumnWidth = 8   ' 行番号
    ws.Columns(4).ColumnWidth = 20  ' エラー種別
    ws.Columns(5).ColumnWidth = 40  ' 詳細メッセージ
    ws.Columns(6).ColumnWidth = 20  ' 問題の値
End Sub

' ============================================================
' SetupMonthlySheet — 月次サマリーシートのレイアウトを設定する（プライベート）
'
' 設定内容:
'   1行目: タイトル "月次売上サマリー"
'   2行目: 列ヘッダー (年月 / 売上金額合計 / 数量合計 / 取り分合計 / 件数)
'   「月次サマリー更新」ボタン
'
' データ行(3行目以降)は RunAll のたびに更新される。
' ============================================================
Private Sub SetupMonthlySheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MONTHLY)

    ws.Cells(1, 1).Value     = "月次売上サマリー"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14

    ' ヘッダー行
    ws.Cells(2, 1).Value = "年月"
    ws.Cells(2, 2).Value = "売上金額合計"
    ws.Cells(2, 3).Value = "売上数量合計"
    ws.Cells(2, 4).Value = "部署取り分合計"
    ws.Cells(2, 5).Value = "レコード数"

    With ws.Rows(2)
        .Font.Bold      = True
        .Interior.Color = CLR_MONTHLY_HDR  ' modConfig 定数: RGB(200,240,220) 薄緑
    End With

    ws.Columns(1).ColumnWidth   = 14  ' 年月
    ws.Columns(2).ColumnWidth   = 16  ' 売上金額合計
    ws.Columns(3).ColumnWidth   = 14  ' 売上数量合計
    ws.Columns(4).ColumnWidth   = 16  ' 部署取り分合計
    ws.Columns(5).ColumnWidth   = 10  ' 件数

    Dim btn As Object
    Set btn = ws.Buttons.Add(10, 5, 140, 28)
    btn.Caption  = "月次サマリー更新"
    btn.OnAction = "modMonthly.BuildMonthly"
End Sub

' ============================================================
' InjectAggrEvent — 集計シートに Worksheet_Change イベントを動的注入する（プライベート）
' ============================================================
Private Sub InjectAggrEvent()
    Dim ws As Worksheet
    Dim codeModule As Object
    Dim code As String

    Set ws         = ThisWorkbook.Sheets(SH_AGGR)
    Set codeModule = ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule

    code = "Option Explicit" & vbNewLine & vbNewLine & _
           "Private Sub Worksheet_Change(ByVal Target As Range)" & vbNewLine & _
           "    Dim triggerRange As Range" & vbNewLine & _
           "    Set triggerRange = Me.Range(AGGR_DEPT_CELL & "","" & AGGR_FROM_CELL & "","" & AGGR_TO_CELL)" & vbNewLine & _
           "    If Intersect(Target, triggerRange) Is Nothing Then Exit Sub" & vbNewLine & _
           "    Application.ScreenUpdating = False" & vbNewLine & _
           "    Application.Calculation = xlCalculationManual" & vbNewLine & _
           "    Application.EnableEvents = False" & vbNewLine & _
           "    On Error GoTo ErrHandler" & vbNewLine & _
           "    modAggregation.Rebuild" & vbNewLine & _
           "ErrHandler:" & vbNewLine & _
           "    Application.ScreenUpdating = True" & vbNewLine & _
           "    Application.Calculation = xlCalculationAutomatic" & vbNewLine & _
           "    Application.EnableEvents = True" & vbNewLine & _
           "End Sub"

    codeModule.AddFromString code
End Sub
