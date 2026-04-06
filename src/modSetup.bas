Attribute VB_Name = "modSetup"

Option Explicit



Public Sub InitWorkbook()

    ' Step 1: Rename placeholder sheet to 集計

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets

        If ws.Name = "Shuukei" Or ws.Name = "Sheet4" Or ws.Name = "Sheet3" Then

            ws.Name = SH_AGGR

            Exit For

        End If

    Next ws



    SetupMainSheet

    SetupConfigSheet

    SetupAllSheet

    SetupAggrSheet

    InjectAggrEvent

End Sub



Private Sub SetupMainSheet()

    Dim ws As Worksheet

    Dim btn As Object



    Set ws = ThisWorkbook.Sheets(SH_MAIN)

    ws.Cells(1, 1).Value = "実行ログ"

    ws.Cells(2, 1).Value = "日時"

    ws.Cells(2, 2).Value = "メッセージ"

    ws.Cells(1, 1).Font.Bold = True

    With ws.Rows(2)

        .Font.Bold = True

        .Interior.Color = RGB(200, 220, 240)

    End With

    ws.Columns(1).ColumnWidth = 22

    ws.Columns(2).ColumnWidth = 80



    ' Add command button

    Set btn = ws.Buttons.Add(10, 10, 160, 30)

    btn.Caption = "ファイルを読み込む"

    btn.OnAction = "modUIControl.RunAll"

End Sub



Private Sub SetupConfigSheet()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(SH_CONFIG)



    ' 製品マスタ (A1:B)

    ws.Cells(1, 1).Value = "製品マスタ"

    ws.Cells(2, 1).Value = "製品コード"

    ws.Cells(2, 2).Value = "製品名"



    ' 口銭マスタ (D1:E)

    ws.Cells(1, 4).Value = "口銭マスタ"

    ws.Cells(2, 4).Value = "売上種別"

    ws.Cells(2, 5).Value = "口銭比率%"



    ' ヘッダー名寄せ (G1:H)

    ws.Cells(1, 7).Value = "ヘッダー名寄せ設定"

    ws.Cells(2, 7).Value = "正規名"

    ws.Cells(2, 8).Value = "対応列名（カンマ区切り）"

    ws.Cells(2, 9).Value = "Allシート列名"



    ' 部署リスト (J1:J)

    ws.Cells(1, 10).Value = "集計用部署リスト"

    ws.Cells(2, 10).Value = "全部署"



    ' Bold section headers

    ws.Cells(1, 1).Font.Bold = True

    ws.Cells(1, 4).Font.Bold = True

    ws.Cells(1, 7).Font.Bold = True

    ws.Cells(1, 10).Font.Bold = True



    ' Bold column headers

    ws.Range("A2:B2").Font.Bold = True

    ws.Range("D2:E2").Font.Bold = True

    ws.Range("G2:I2").Font.Bold = True

    ws.Range("J2").Font.Bold = True



    ws.Columns("A:B").ColumnWidth = 16

    ws.Columns("D:E").ColumnWidth = 14

    ws.Columns("G:H").ColumnWidth = 20

    ws.Columns("I").ColumnWidth = 16

    ws.Columns("J").ColumnWidth = 16



    ' SharePoint連携 (L1:M)

    ws.Cells(1, CFG_PA_LABEL_COL).Value = "SharePoint連携"

    ws.Cells(1, CFG_PA_LABEL_COL).Font.Bold = True

    ws.Cells(2, CFG_PA_LABEL_COL).Value = "PowerAutomate URL"

    ws.Cells(2, CFG_PA_LABEL_COL).Font.Bold = True

    ws.Columns("L").ColumnWidth = 20

    ws.Columns("M").ColumnWidth = 60



    ' Sample 製品マスタ data

    ws.Cells(3, 1).Value = "P001": ws.Cells(3, 2).Value = "製品A"

    ws.Cells(4, 1).Value = "P002": ws.Cells(4, 2).Value = "製品B"



    ' Sample 口銭マスタ data

    ws.Cells(3, 4).Value = "直販":  ws.Cells(3, 5).Value = 10

    ws.Cells(4, 4).Value = "代理店": ws.Cells(4, 5).Value = 5



    ' Sample 名寄せ data

    ws.Cells(3, 7).Value = HDR_CLIENT:    ws.Cells(3, 8).Value = "得意先名,得意先コード,顧客名": ws.Cells(3, 9).Value = HDR_CLIENT

    ws.Cells(4, 7).Value = HDR_PROD_CODE: ws.Cells(4, 8).Value = "品番,ProductCode": ws.Cells(4, 9).Value = HDR_PROD_CODE

    ws.Cells(5, 7).Value = HDR_AMOUNT:    ws.Cells(5, 8).Value = "金額,Amount,売上高": ws.Cells(5, 9).Value = HDR_AMOUNT

    ws.Cells(6, 7).Value = HDR_UNIT_PRICE: ws.Cells(6, 8).Value = "単価,定価": ws.Cells(6, 9).Value = HDR_UNIT_PRICE

    ws.Cells(7, 7).Value = HDR_QTY:       ws.Cells(7, 8).Value = "数量,Qty": ws.Cells(7, 9).Value = HDR_QTY

    ws.Cells(8, 7).Value = HDR_DATE:      ws.Cells(8, 8).Value = "日付,売上日,Date": ws.Cells(8, 9).Value = HDR_DATE

    ws.Cells(9, 7).Value = HDR_SALE_TYPE: ws.Cells(9, 8).Value = "取引区分,SaleType": ws.Cells(9, 9).Value = HDR_SALE_TYPE

    ws.Cells(10, 7).Value = HDR_DEPT:     ws.Cells(10, 8).Value = "部門,Dept": ws.Cells(10, 9).Value = HDR_DEPT

End Sub



Private Sub SetupAllSheet()

    Dim ws As Worksheet

    Dim dictAllColDef As Object

    Dim i As Integer

    Dim k As Variant

    Dim totalCols As Integer



    Set ws = ThisWorkbook.Sheets(SH_ALL)

    Set dictAllColDef = LoadAllColDef()



    ' 動的ヘッダー書き込み

    i = 1

    For Each k In dictAllColDef.Keys

        ws.Cells(1, i).Value = dictAllColDef(k)

        i = i + 1

    Next k

    ws.Cells(1, i).Value = HDR_PROD_NAME

    ws.Cells(1, i + 1).Value = HDR_MARGIN

    ws.Cells(1, i + 2).Value = HDR_SOURCE



    totalCols = i + 2



    With ws.Rows(1)

        .Font.Bold = True

        .Interior.Color = RGB(200, 220, 240)

    End With

    ws.Range(ws.Cells(1, 1), ws.Cells(1, totalCols)).EntireColumn.AutoFit



    ' SharePointアップロードボタン

    Dim uploadBtn As Object

    Set uploadBtn = ws.Buttons.Add(700, 5, 180, 28)

    uploadBtn.Caption = "SharePointへアップロード"

    uploadBtn.OnAction = "modSharePoint.UploadAllToSharePoint"

End Sub



Private Sub SetupAggrSheet()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(SH_AGGR)



    ' Filter labels

    ws.Cells(1, 1).Value = "部署選択"

    ws.Cells(2, 1).Value = "開始日"

    ws.Cells(3, 1).Value = "終了日"

    ws.Range("A1:A3").Font.Bold = True

    ws.Range(AGGR_DEPT_CELL).Value = "全部署"



    ' Aggregate header row

    ws.Cells(AGGR_HDR_ROW, 2).Value = "売上金額合計"

    ws.Cells(AGGR_HDR_ROW, 3).Value = "売上数量合計"

    ws.Cells(AGGR_HDR_ROW, 4).Value = "口銭総額"

    With ws.Rows(AGGR_HDR_ROW)

        .Font.Bold = True

        .Interior.Color = RGB(200, 220, 240)

    End With



    ws.Columns("A").ColumnWidth = 30

    ws.Columns("B:D").ColumnWidth = 15



    ' Add chart button

    Dim chartBtn As Object

    Set chartBtn = ws.Buttons.Add(330, 5, 150, 28)

    chartBtn.Caption = "グラフ作成"

    chartBtn.OnAction = "modChart.DrawAggrChart"



    ' Add upload button

    Dim uploadBtn As Object

    Set uploadBtn = ws.Buttons.Add(490, 5, 180, 28)

    uploadBtn.Caption = "SharePointへアップロード"

    uploadBtn.OnAction = "modSharePoint.UploadToSharePoint"

End Sub



Private Sub InjectAggrEvent()

    ' 【注意】注入するコード文字列の中で modConfig の定数（AGGR_DEPT_CELL, AGGR_FROM_CELL, AGGR_TO_CELL）
    ' を直接参照している。定数名を変更した場合はここの文字列も合わせて変更すること。

    ' Requires "Trust access to the VBA project object model" to be enabled in Excel Trust Center

    Dim ws As Worksheet

    Dim codeModule As Object

    Dim code As String



    Set ws = ThisWorkbook.Sheets(SH_AGGR)

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

