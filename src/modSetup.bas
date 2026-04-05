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
' 処理順序:
'   1. 仮名シート (Shuukei → "集計"、Pivot → "ピボット") をリネーム
'   2. 各シートのレイアウト設定
'   3. 集計シートへの Worksheet_Change イベントコード注入
' ============================================================
Public Sub InitWorkbook()
    ' create_workbook.vbs で作成した仮名シートを正式名にリネーム
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case "Shuukei", "Sheet4", "Sheet3"
                ws.Name = SH_AGGR   ' "Shuukei" → "集計"
            Case "Pivot"
                ws.Name = SH_PIVOT  ' "Pivot"   → "ピボット"
        End Select
    Next ws

    SetupMainSheet
    SetupConfigSheet
    SetupAllSheet
    SetupAggrSheet
    SetupPivotSheet
    InjectAggrEvent
End Sub

' ============================================================
' SetupMainSheet — main シートのレイアウトとボタンを設定する（プライベート）
'
' 設定内容:
'   ・1行目: "実行ログ" タイトル
'   ・2行目: 列ヘッダー（日時 / メッセージ）
'   ・「ファイルを読み込む」ボタン → modUIControl.RunAll に接続
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
        .Interior.Color = RGB(200, 220, 240)
    End With
    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 80

    ' ボタン配置: (Left, Top, Width, Height) 単位はポイント
    Set btn = ws.Buttons.Add(10, 10, 160, 30)
    btn.Caption  = "ファイルを読み込む"
    btn.OnAction = "modUIControl.RunAll"
End Sub

' ============================================================
' SetupConfigSheet — Config シートのレイアウトとサンプルデータを設定する（プライベート）
'
' 設定内容:
'   A–B列: 製品マスタ（ヘッダー + サンプル2件）
'   D–E列: 口銭マスタ（ヘッダー + サンプル2件）
'   G–H列: ヘッダー名寄せ設定（ヘッダー + 全8正規名のサンプル）
'   J列  : 集計用部署リスト（ヘッダーのみ）
'   L–M列: SharePoint連携（ラベル + URL 入力欄）
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
    ws.Cells(2, 10).Value = "全部署"  ' J2 は "全部署" 固定（RunAll 後に J3〜 が更新される）

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
    ws.Cells(3, 4).Value = "直販"  : ws.Cells(3, 5).Value = 10
    ws.Cells(4, 4).Value = "代理店" : ws.Cells(4, 5).Value = 5

    ' --- ヘッダー名寄せ サンプルデータ（HDR_* 定数を使うことで正規名と一致を保証）---
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
    ws.Cells(2, CFG_PA_LABEL_COL).Value = "PowerAutomate URL"
    ws.Cells(2, CFG_PA_LABEL_COL).Font.Bold = True
    ' M2 は URL 入力セル（LoadPowerAutomateUrl が CFG_PA_URL_ROW=2, CFG_PA_URL_COL=13 で参照）
    ws.Columns("L").ColumnWidth = 20
    ws.Columns("M").ColumnWidth = 60
End Sub

' ============================================================
' SetupAllSheet — all シートのヘッダー行とボタンを設定する（プライベート）
'
' ヘッダー文字列は HDR_* 定数から書き込むことで
' modConfig との一貫性を保証する。
' アップロードボタンはヘッダー行右外側（列L付近）に配置する。
' ============================================================
Private Sub SetupAllSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ALL)

    ' --- 11列のヘッダーを設定（HDR_* 定数を使用）---
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
        .Interior.Color = RGB(200, 220, 240)
    End With
    ws.Columns("A:K").AutoFit  ' 初期列幅をデータに合わせて自動設定

    ' --- SharePoint アップロードボタン（データ列A〜Kの右外側に配置）---
    Dim uploadBtn As Object
    Set uploadBtn = ws.Buttons.Add(700, 5, 180, 28)
    uploadBtn.Caption  = "SharePointへアップロード"
    uploadBtn.OnAction = "modSharePoint.UploadAllToSharePoint"
End Sub

' ============================================================
' SetupAggrSheet — 集計シートのレイアウトとボタンを設定する（プライベート）
'
' 設定内容:
'   A1:B1  部署選択ラベル / ドロップダウン（B1）
'   A2:B2  開始日ラベル / 日付入力（B2）
'   A3:B3  終了日ラベル / 日付入力（B3）
'   5行目  集計テーブルのヘッダー（B〜D列）
'   グラフ作成ボタン / SharePoint アップロードボタン（右上エリア）
' ============================================================
Private Sub SetupAggrSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_AGGR)

    ' --- フィルタ入力欄 ---
    ws.Cells(1, 1).Value = "部署選択"
    ws.Cells(2, 1).Value = "開始日"
    ws.Cells(3, 1).Value = "終了日"
    ws.Range("A1:A3").Font.Bold     = True
    ws.Range(AGGR_DEPT_CELL).Value  = "全部署"  ' B1 の初期値

    ' --- 集計テーブルのヘッダー行（5行目）---
    ws.Cells(AGGR_HDR_ROW, 2).Value = "売上金額合計"
    ws.Cells(AGGR_HDR_ROW, 3).Value = "売上数量合計"
    ws.Cells(AGGR_HDR_ROW, 4).Value = "口銭総額"
    With ws.Rows(AGGR_HDR_ROW)
        .Font.Bold      = True
        .Interior.Color = RGB(200, 220, 240)
    End With

    ws.Columns("A").ColumnWidth   = 30
    ws.Columns("B:D").ColumnWidth = 15

    ' --- グラフ作成ボタン（フィルタ欄の右側に配置）---
    Dim chartBtn As Object
    Set chartBtn = ws.Buttons.Add(330, 5, 150, 28)
    chartBtn.Caption  = "グラフ作成"
    chartBtn.OnAction = "modChart.DrawAggrChart"

    ' --- SharePoint アップロードボタン（グラフ作成ボタンの右に配置）---
    Dim uploadBtn As Object
    Set uploadBtn = ws.Buttons.Add(490, 5, 180, 28)
    uploadBtn.Caption  = "SharePointへアップロード"
    uploadBtn.OnAction = "modSharePoint.UploadToSharePoint"
End Sub

' ============================================================
' SetupPivotSheet — ピボットシートの UI 領域（タイトル・ボタン）を設定する（プライベート）
'
' 設定内容:
'   A1: タイトル "売上ピボットテーブル"（太字）
'   A2: 使い方の説明文
'   「ピボットテーブル更新」ボタン → modPivot.BuildPivot に接続
'
' PivotTable 本体は modPivot.BuildPivot が行 4 以降に動的に作成する。
' このサブルーチンは UI 部分のみを担当する。
' ============================================================
Private Sub SetupPivotSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PIVOT)

    ' --- タイトルと使い方説明 ---
    ws.Cells(1, 1).Value    = "売上ピボットテーブル"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(2, 1).Value    = "RunAll 実行時に自動更新されます。" & _
                               "フィールドリストで行・列・フィルター・値を自由に配置できます。"
    ws.Cells(2, 1).Font.Color = RGB(100, 100, 100)
    ws.Columns("A").ColumnWidth = 35

    ' --- 更新ボタン（タイトル行の右側に配置）---
    Dim btn As Object
    Set btn = ws.Buttons.Add(400, 5, 160, 28)
    btn.Caption  = "ピボットテーブル更新"
    btn.OnAction = "modPivot.BuildPivot"
End Sub

' ============================================================
' InjectAggrEvent — 集計シートに Worksheet_Change イベントを動的注入する（プライベート）
'
' 集計シートの B1/B2/B3（部署・開始日・終了日）が変更された際に
' 自動的に modAggregation.Rebuild を呼び出すイベントハンドラを
' VBA コードとしてシートモジュールに埋め込む。
'
' 動的注入が必要な理由:
'   create_workbook.vbs は .bas ファイルをインポートする仕組みのため、
'   シートモジュール（集計シートのコードウィンドウ）への直接書き込みが
'   スクリプトからは困難。そのため InitWorkbook 実行時に VBProject 経由で
'   AddFromString によりコードを埋め込む方式を採用している。
' ============================================================
Private Sub InjectAggrEvent()
    Dim ws As Worksheet
    Dim codeModule As Object
    Dim code As String

    Set ws         = ThisWorkbook.Sheets(SH_AGGR)
    Set codeModule = ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule

    ' 注入するコード文字列を組み立てる
    ' triggerRange: B1/B2/B3 のいずれかが変更された場合に Rebuild を呼ぶ
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
