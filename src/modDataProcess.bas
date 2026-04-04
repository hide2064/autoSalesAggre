Attribute VB_Name = "modDataProcess"
Option Explicit

' ============================================================
' modDataProcess — データ加工モジュール
'
' 役割:
'   ・LoadTsvToSheet で作成された個別ソースシートを走査し、
'     ヘッダー名寄せ・製品マスタ参照・口銭計算を行いながら
'     all シートへ正規化データを集約する BuildAllSheet を提供する。
'   ・all シートから部署名の一覧を収集する CollectUniqueDepts を提供する。
'
' 設計方針:
'   ・ソースシートへの逐次セルアクセスを避け、Variant 配列一括読み込み
'     → 計算 → Variant 配列一括書き込み の方式を採用する。
'   ・未登録の製品コード・売上種別はログに警告を出力し、
'     製品名は "[未登録]"、部署取り分は 0 として処理を継続する。
'   ・dictSummary のキーは「製品名 & "||" & 客先名」形式。
'     "||" はセパレータとして通常テキストに含まれにくい文字列を選択。
' ============================================================

' ============================================================
' BuildAllSheet — 全ソースシートを all シートに集約する
'
' 引数:
'   dictProduct    — LoadProductDict() が返す製品コード→製品名辞書
'   dictCommission — LoadCommissionDict() が返す売上種別→口銭率辞書
'   dictHeaderMap  — LoadHeaderMap() が返すエイリアス→正規名辞書
'
' 処理概要:
'   1. all シートの既存データ行(2行目以降)をクリアしヘッダーを再書き込み
'   2. ワークブック内の固定4シート(main/Config/all/集計)以外の全シートを
'      ソースシートとして ProcessSourceSheet に渡す
'   3. ProcessSourceSheet の戻り値(次の書き込み開始行)を累積して
'      all シートに順次追記する
' ============================================================
Public Sub BuildAllSheet(dictProduct As Object, dictCommission As Object, dictHeaderMap As Object)
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim allRowNum As Long  ' all シートへの次の書き込み行番号

    Set wsAll = ThisWorkbook.Sheets(SH_ALL)

    ' --- all シートのデータ行をクリア（ヘッダー行は保持）---
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then wsAll.Rows("2:" & lastRow).ClearContents

    ' --- ヘッダー行を再書き込み（定数を使うことでヘッダー名の一貫性を保証）---
    wsAll.Cells(1, ALL_COL_CLIENT).Value     = HDR_CLIENT
    wsAll.Cells(1, ALL_COL_PROD_CODE).Value  = HDR_PROD_CODE
    wsAll.Cells(1, ALL_COL_AMOUNT).Value     = HDR_AMOUNT
    wsAll.Cells(1, ALL_COL_UNIT_PRICE).Value = HDR_UNIT_PRICE
    wsAll.Cells(1, ALL_COL_QTY).Value        = HDR_QTY
    wsAll.Cells(1, ALL_COL_DATE).Value       = HDR_DATE
    wsAll.Cells(1, ALL_COL_SALE_TYPE).Value  = HDR_SALE_TYPE
    wsAll.Cells(1, ALL_COL_DEPT).Value       = HDR_DEPT
    wsAll.Cells(1, ALL_COL_PROD_NAME).Value  = HDR_PROD_NAME
    wsAll.Cells(1, ALL_COL_MARGIN).Value     = HDR_MARGIN
    wsAll.Cells(1, ALL_COL_SOURCE).Value     = HDR_SOURCE

    allRowNum = 2  ' データ書き込みは2行目から開始

    ' --- 固定シート以外をソースシートとして処理 ---
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case SH_MAIN, SH_CONFIG, SH_ALL, SH_AGGR
                ' 固定シートはスキップ
            Case Else
                ' ソースシートを処理し、次の書き込み行番号を受け取る
                allRowNum = ProcessSourceSheet(ws, wsAll, allRowNum, dictProduct, dictCommission, dictHeaderMap)
        End Select
    Next ws
End Sub

' ============================================================
' ProcessSourceSheet — 1枚のソースシートを処理して all シートに書き込む（プライベート）
'
' 引数:
'   wsSrc          — 処理対象のソースシート（TSV読み込みで作成）
'   wsAll          — 書き込み先の all シート
'   startRow       — all シートへの書き込み開始行番号
'   dictProduct    — 製品コード→製品名辞書
'   dictCommission — 売上種別→口銭率辞書
'   dictHeaderMap  — エイリアス→正規名辞書
'
' 戻り値: 次のソースシート処理時の書き込み開始行番号
'
' 処理概要:
'   1. ソースシートの1行目(ヘッダー)を走査し、dictHeaderMap を使って
'      各列が all シートのどの列に対応するかのマッピング配列(colMap)を作成
'   2. ソースシートの2行目以降を Variant 配列に一括読み込み
'   3. 各行について colMap に従って値をコピーし、製品名・部署取り分を計算
'   4. 結果を Variant 配列として all シートに一括書き込み
'
' colMap の構造:
'   colMap(allColIndex - 1) = ソースシートの列番号（0 はマップなし）
'   allColIndex は ALL_COL_CLIENT(1) 〜 ALL_COL_DEPT(8) の範囲
' ============================================================
Private Function ProcessSourceSheet(wsSrc As Worksheet, wsAll As Worksheet, _
    startRow As Long, dictProduct As Object, dictCommission As Object, _
    dictHeaderMap As Object) As Long

    Dim lastSrcRow As Long   ' ソースシートの最終行
    Dim lastSrcCol As Integer ' ソースシートの最終列
    Dim colMap(7) As Integer  ' all 列インデックス(0始まり) → ソース列番号 のマッピング
    Dim c As Integer
    Dim srcHeader As String   ' ソースシートのヘッダー文字列（小文字化して辞書検索）
    Dim srcData As Variant    ' ソースシートのデータ一括読み込み用
    Dim numRows As Long       ' 処理行数（ヘッダー除く）
    Dim outArr() As Variant   ' all シートへの一括書き込み用出力配列
    Dim r As Long
    Dim allCol As Integer
    Dim prodCode As String    ' 各行の製品コード（製品名計算用）
    Dim saleType As String    ' 各行の売上種別（口銭計算用）
    Dim amount As Double      ' 各行の売上金額（口銭計算用）

    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    ' データ行が1行もない場合は何もせずに開始行をそのまま返す
    If lastSrcRow < 2 Then
        ProcessSourceSheet = startRow
        Exit Function
    End If

    lastSrcCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    ' ============================================================
    ' ヘッダーマッピングの構築
    ' ソースシートの各列ヘッダーを小文字化して dictHeaderMap で検索し、
    ' 正規名に対応する all シートの列番号を colMap に記録する。
    ' dictHeaderMap にないヘッダーは colMap が 0 のままとなり、
    ' 後続処理で空文字として扱われる。
    ' ============================================================
    For c = 1 To lastSrcCol
        srcHeader = LCase(Trim(CStr(wsSrc.Cells(1, c).Value)))
        If dictHeaderMap.Exists(srcHeader) Then
            Select Case dictHeaderMap(srcHeader)
                Case HDR_CLIENT:      colMap(ALL_COL_CLIENT - 1)     = c
                Case HDR_PROD_CODE:   colMap(ALL_COL_PROD_CODE - 1)  = c
                Case HDR_AMOUNT:      colMap(ALL_COL_AMOUNT - 1)     = c
                Case HDR_UNIT_PRICE:  colMap(ALL_COL_UNIT_PRICE - 1) = c
                Case HDR_QTY:         colMap(ALL_COL_QTY - 1)        = c
                Case HDR_DATE:        colMap(ALL_COL_DATE - 1)       = c
                Case HDR_SALE_TYPE:   colMap(ALL_COL_SALE_TYPE - 1)  = c
                Case HDR_DEPT:        colMap(ALL_COL_DEPT - 1)       = c
            End Select
        End If
    Next c

    ' --- ソースシートのデータ行を Variant 配列に一括読み込み ---
    srcData = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastSrcRow, lastSrcCol)).Value

    numRows = lastSrcRow - 1
    ReDim outArr(1 To numRows, 1 To ALL_TOTAL_COLS)

    ' ============================================================
    ' 各行の処理: マッピング・製品名計算・部署取り分計算
    ' ============================================================
    For r = 1 To numRows
        ' --- 列1〜8 (客先名〜部署): colMap に従ってソース値をコピー ---
        For allCol = ALL_COL_CLIENT To ALL_COL_DEPT
            If colMap(allCol - 1) > 0 Then
                outArr(r, allCol) = srcData(r, colMap(allCol - 1))
            Else
                outArr(r, allCol) = ""  ' マップなし列は空文字
            End If
        Next allCol

        ' --- 列9: 製品名 — 製品マスタから逆引き ---
        prodCode = Trim(CStr(outArr(r, ALL_COL_PROD_CODE)))
        If dictProduct.Exists(prodCode) Then
            outArr(r, ALL_COL_PROD_NAME) = dictProduct(prodCode)
        Else
            outArr(r, ALL_COL_PROD_NAME) = "[未登録]"  ' 未登録コードは明示的に記録
            If prodCode <> "" Then
                LogMessage "警告: 製品コード未登録 [" & prodCode & "] (" & wsSrc.Name & ")"
            End If
        End If

        ' --- 列10: 部署取り分 — 売上金額 × 口銭率 / 100 ---
        saleType = Trim(CStr(outArr(r, ALL_COL_SALE_TYPE)))
        amount = 0
        If IsNumeric(outArr(r, ALL_COL_AMOUNT)) Then amount = CDbl(outArr(r, ALL_COL_AMOUNT))
        If dictCommission.Exists(saleType) Then
            outArr(r, ALL_COL_MARGIN) = amount * dictCommission(saleType) / 100
        Else
            outArr(r, ALL_COL_MARGIN) = 0  ' 未登録の売上種別は取り分 0 として処理続行
            If saleType <> "" Then
                LogMessage "警告: 売上種別未登録 [" & saleType & "] (" & wsSrc.Name & ")"
            End If
        End If

        ' --- 列11: ソースファイル名（シート名 = 拡張子なしファイル名）---
        outArr(r, ALL_COL_SOURCE) = wsSrc.Name
    Next r

    ' --- 出力配列を all シートに一括書き込み ---
    wsAll.Range(wsAll.Cells(startRow, 1), wsAll.Cells(startRow + numRows - 1, ALL_TOTAL_COLS)).Value = outArr

    ' 次のソースシートの書き込み開始行を返す
    ProcessSourceSheet = startRow + numRows
End Function

' ============================================================
' CollectUniqueDepts — all シートから部署名の一覧を収集する
'
' 戻り値: Dictionary(部署名 As String → 1) の重複なし辞書
'         部署取り分ドロップダウンの更新(RefreshDeptList)に渡す。
' ============================================================
Public Function CollectUniqueDepts() As Object
    Dim dict As Object
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim deptArr As Variant  ' 部署列の一括読み込み用
    Dim r As Long
    Dim dept As String

    Set dict = NewDict()
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)

    lastRow = wsAll.Cells(wsAll.Rows.Count, ALL_COL_DEPT).End(xlUp).Row
    If lastRow < 2 Then
        Set CollectUniqueDepts = dict  ' データなしの場合は空の辞書を返す
        Exit Function
    End If

    ' 部署列のみを一括読み込み（全列読み込みより効率的）
    deptArr = wsAll.Range(wsAll.Cells(2, ALL_COL_DEPT), wsAll.Cells(lastRow, ALL_COL_DEPT)).Value

    For r = 1 To UBound(deptArr, 1)
        dept = Trim(CStr(deptArr(r, 1)))
        ' 空欄・重複を除外して部署名を辞書に登録
        If dept <> "" And Not dict.Exists(dept) Then dict(dept) = 1
    Next r

    Set CollectUniqueDepts = dict
End Function
