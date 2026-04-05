Attribute VB_Name = "modDataProcess"
Option Explicit

' ============================================================
' modDataProcess — データ加工モジュール
'
' 役割:
'   ・LoadFileToSheet で作成された個別ソースシートを走査し、
'     ヘッダー名寄せ・製品マスタ参照・口銭計算を行いながら
'     all シートへ正規化データを集約する BuildAllSheet を提供する。
'   ・all シートから部署名の一覧を収集する CollectUniqueDepts を提供する。
'
' 設計方針:
'   ・ソースシートへの逐次セルアクセスを避け、Variant 配列一括読み込み
'     → 計算 → Variant 配列一括書き込み の方式を採用する。
'   ・未登録の製品コード・売上種別はログに警告を出力し、
'     modError.LogError でエラーシートにも記録する。
'     製品名は "[未登録]"、部署取り分は 0 として処理を継続する。
'   ・重複行の検出: 同一ソースファイル内で全フィールドが一致する行を
'     dictDedupKeys で追跡する。重複行はエラーシートに記録して除外する。
'     （同じファイルを誤って2回選択した場合の保護）
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
'   2. ワークブック内の固定シート以外の全シートを
'      ソースシートとして ProcessSourceSheet に渡す
'   3. 重複排除辞書(dictDedupKeys)を全ソースシート間で共有し、
'      同一ファイルの二重読み込みを検出する
' ============================================================
Public Sub BuildAllSheet(dictProduct As Object, dictCommission As Object, dictHeaderMap As Object)
    Dim wsAll As Worksheet
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim allRowNum As Long
    Dim dictDedupKeys As Object  ' 重複排除用: キー=ソース名+全フィールド結合, 値=1

    Set wsAll = ThisWorkbook.Sheets(SH_ALL)

    ' --- all シートのデータ行をクリア（ヘッダー行は保持）---
    lastRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then wsAll.Rows("2:" & lastRow).ClearContents

    ' --- ヘッダー行を再書き込み（HDR_* 定数で一貫性を保証）---
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

    allRowNum = 2

    ' 重複排除辞書: 全ソースシートにまたがって共有する
    ' キー = ソースファイル名 & DICT_KEY_SEP & 主要8フィールドの結合文字列
    Set dictDedupKeys = NewDict()

    ' --- 固定シート以外をソースシートとして処理 ---
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case SH_MAIN, SH_CONFIG, SH_ALL, SH_AGGR, SH_PIVOT, SH_ERROR, SH_MONTHLY
                ' システム用固定シートはスキップ
            Case Else
                allRowNum = ProcessSourceSheet( _
                    ws, wsAll, allRowNum, _
                    dictProduct, dictCommission, dictHeaderMap, dictDedupKeys)
        End Select
    Next ws
End Sub

' ============================================================
' ProcessSourceSheet — 1枚のソースシートを処理して all シートに書き込む（プライベート）
'
' 引数:
'   wsSrc          — 処理対象のソースシート
'   wsAll          — 書き込み先の all シート
'   startRow       — all シートへの書き込み開始行番号
'   dictProduct    — 製品コード→製品名辞書
'   dictCommission — 売上種別→口銭率辞書
'   dictHeaderMap  — エイリアス→正規名辞書
'   dictDedupKeys  — 重複排除用辞書（全ソースシートで共有）
'
' 戻り値: 次のソースシート処理時の書き込み開始行番号
'
' colMap の構造:
'   colMap(allColIndex - 1) = ソースシートの列番号（0 はマップなし）
'   allColIndex は ALL_COL_CLIENT(1) 〜 ALL_COL_DEPT(8) の範囲
' ============================================================
Private Function ProcessSourceSheet(wsSrc As Worksheet, wsAll As Worksheet, _
    startRow As Long, dictProduct As Object, dictCommission As Object, _
    dictHeaderMap As Object, dictDedupKeys As Object) As Long

    Dim lastSrcRow As Long
    Dim lastSrcCol As Integer
    Dim colMap(COL_MAP_COUNT - 1) As Integer
    Dim c As Integer
    Dim srcHeader As String
    Dim srcData As Variant
    Dim numRows As Long
    Dim outArr() As Variant
    Dim writeRow As Long    ' 実際に書き込む行数カウンター
    Dim r As Long
    Dim allCol As Integer
    Dim prodCode As String
    Dim saleType As String
    Dim amount As Double
    Dim dedupKey As String  ' 重複排除判定用のキー文字列

    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    If lastSrcRow < 2 Then
        ProcessSourceSheet = startRow
        Exit Function
    End If

    lastSrcCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    ' ============================================================
    ' ヘッダーマッピングの構築
    ' ソースシートの各列ヘッダーを小文字化して dictHeaderMap で検索し、
    ' 正規名に対応する all シートの列番号を colMap に記録する。
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
    writeRow = 0  ' outArr に実際に格納した行数

    ' ============================================================
    ' 各行の処理: マッピング・重複排除・製品名計算・部署取り分計算
    ' ============================================================
    For r = 1 To numRows

        ' --- 列1〜8 (客先名〜部署): colMap に従ってソース値をコピー ---
        For allCol = ALL_COL_CLIENT To ALL_COL_DEPT
            If colMap(allCol - 1) > 0 Then
                outArr(writeRow + 1, allCol) = srcData(r, colMap(allCol - 1))
            Else
                outArr(writeRow + 1, allCol) = ""
            End If
        Next allCol

        ' ============================================================
        ' 重複行の検出
        ' キー = ソースファイル名 & DICT_KEY_SEP & 客先名 & ... & 部署
        ' 同じファイルが2回選択された場合、2回目の同一行を除外する。
        ' ============================================================
        dedupKey = wsSrc.Name & DICT_KEY_SEP & _
                   CStr(outArr(writeRow + 1, ALL_COL_CLIENT))     & DICT_KEY_SEP & _
                   CStr(outArr(writeRow + 1, ALL_COL_PROD_CODE))  & DICT_KEY_SEP & _
                   CStr(outArr(writeRow + 1, ALL_COL_AMOUNT))     & DICT_KEY_SEP & _
                   CStr(outArr(writeRow + 1, ALL_COL_UNIT_PRICE)) & DICT_KEY_SEP & _
                   CStr(outArr(writeRow + 1, ALL_COL_QTY))        & DICT_KEY_SEP & _
                   CStr(outArr(writeRow + 1, ALL_COL_DATE))       & DICT_KEY_SEP & _
                   CStr(outArr(writeRow + 1, ALL_COL_SALE_TYPE))  & DICT_KEY_SEP & _
                   CStr(outArr(writeRow + 1, ALL_COL_DEPT))

        If dictDedupKeys.Exists(dedupKey) Then
            ' 重複行: エラーシートに記録してスキップ
            LogError wsSrc.Name, r + 1, "重複行", _
                     "同一ファイルの行が既に読み込まれています", _
                     CStr(outArr(writeRow + 1, ALL_COL_CLIENT)) & " / " & _
                     CStr(outArr(writeRow + 1, ALL_COL_DATE))
            LogMessage "警告: 重複行を除外しました (" & wsSrc.Name & " 行" & (r + 1) & ")"
            ' outArr のこの行はクリアして次の行へ
            Dim clrCol As Integer
            For clrCol = 1 To ALL_TOTAL_COLS
                outArr(writeRow + 1, clrCol) = ""
            Next clrCol
            GoTo NextRow
        End If
        dictDedupKeys(dedupKey) = 1

        ' --- 列9: 製品名 — 製品マスタから逆引き ---
        prodCode = Trim(CStr(outArr(writeRow + 1, ALL_COL_PROD_CODE)))
        If dictProduct.Exists(prodCode) Then
            outArr(writeRow + 1, ALL_COL_PROD_NAME) = dictProduct(prodCode)
        Else
            outArr(writeRow + 1, ALL_COL_PROD_NAME) = "[未登録]"
            If prodCode <> "" Then
                LogMessage "警告: 製品コード未登録 [" & prodCode & "] (" & wsSrc.Name & " 行" & (r + 1) & ")"
                LogError wsSrc.Name, r + 1, "製品コード未登録", _
                         "製品マスタに存在しないコードです", prodCode
            End If
        End If

        ' --- 列10: 部署取り分 — 売上金額 × 口銭率 / 100 ---
        saleType = Trim(CStr(outArr(writeRow + 1, ALL_COL_SALE_TYPE)))
        amount = 0
        If IsNumeric(outArr(writeRow + 1, ALL_COL_AMOUNT)) Then
            amount = CDbl(outArr(writeRow + 1, ALL_COL_AMOUNT))
        End If
        If dictCommission.Exists(saleType) Then
            outArr(writeRow + 1, ALL_COL_MARGIN) = amount * dictCommission(saleType) / 100
        Else
            outArr(writeRow + 1, ALL_COL_MARGIN) = 0
            If saleType <> "" Then
                LogMessage "警告: 売上種別未登録 [" & saleType & "] (" & wsSrc.Name & " 行" & (r + 1) & ")"
                LogError wsSrc.Name, r + 1, "売上種別未登録", _
                         "口銭マスタに存在しない売上種別です", saleType
            End If
        End If

        ' --- 列11: ソースファイル名 ---
        outArr(writeRow + 1, ALL_COL_SOURCE) = wsSrc.Name
        writeRow = writeRow + 1

NextRow:
    Next r

    ' 書き込む行がある場合のみ all シートに一括書き込み
    If writeRow > 0 Then
        wsAll.Range( _
            wsAll.Cells(startRow, 1), _
            wsAll.Cells(startRow + writeRow - 1, ALL_TOTAL_COLS)).Value = _
            outArr
    End If

    ProcessSourceSheet = startRow + writeRow
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
    Dim deptArr As Variant
    Dim r As Long
    Dim dept As String

    Set dict = NewDict()
    Set wsAll = ThisWorkbook.Sheets(SH_ALL)

    lastRow = wsAll.Cells(wsAll.Rows.Count, ALL_COL_DEPT).End(xlUp).Row
    If lastRow < 2 Then
        Set CollectUniqueDepts = dict
        Exit Function
    End If

    deptArr = wsAll.Range(wsAll.Cells(2, ALL_COL_DEPT), wsAll.Cells(lastRow, ALL_COL_DEPT)).Value

    For r = 1 To UBound(deptArr, 1)
        dept = Trim(CStr(deptArr(r, 1)))
        If dept <> "" And Not dict.Exists(dept) Then dict(dept) = 1
    Next r

    Set CollectUniqueDepts = dict
End Function
