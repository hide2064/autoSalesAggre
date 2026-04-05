Attribute VB_Name = "modFileIO"
Option Explicit

' ============================================================
' modFileIO — ファイル入出力モジュール
'
' 役割:
'   ・ユーザーにファイルを複数選択させる SelectFiles を提供する。
'   ・選択されたファイルを拡張子に応じて適切なローダーに振り分ける
'     LoadFileToSheet を提供する。
'     対応形式: TSV (.tsv/.txt) / CSV (.csv) / Excel (.xlsx/.xls/.xlsm)
'
' ローダーの設計方針:
'   ・TSV / CSV : LoadDelimitedToSheet で一括処理（区切り文字のみ異なる）
'       2パス方式: パス1で行数・列数確定 → パス2で2D配列一括読み込み
'       NumberFormat="@"を事前設定して先頭ゼロの数値変換を防止
'   ・Excel (.xlsx等): LoadXlsxSheetToSheet でブックを読み取り専用で開き
'       第1シートのデータを Variant 配列経由でコピーして閉じる
'   ・同名シートが既に存在する場合は削除して再作成（再読み込み対応）
'   ・新シートは 集計 シートの直前に挿入してシート順を維持する
' ============================================================

' ============================================================
' SelectFiles — ファイルの複数選択ダイアログを表示する
'
' 対応拡張子: .tsv .txt .csv .xlsx .xls .xlsm
'
' 戻り値:
'   ・ファイルが選択された場合: 選択ファイルパスの Variant 配列 (1始まりインデックス)
'   ・キャンセルされた場合    : Boolean の False
'
' 呼び出し元では VarType(戻り値) = vbBoolean でキャンセル判定する。
' ============================================================
Public Function SelectFiles() As Variant
    SelectFiles = Application.GetOpenFilename( _
        FileFilter:="全対応ファイル (*.tsv;*.txt;*.csv;*.xlsx;*.xls;*.xlsm)," & _
                    "*.tsv;*.txt;*.csv;*.xlsx;*.xls;*.xlsm," & _
                    "テキスト/CSVファイル (*.tsv;*.txt;*.csv),*.tsv;*.txt;*.csv," & _
                    "Excelファイル (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm", _
        Title:="読み込むファイルを選択してください（TSV / CSV / Excel）", _
        MultiSelect:=True)
End Function

' ============================================================
' LoadFileToSheet — ファイルを拡張子に応じた方法でシートに読み込む
'
' 引数:
'   filePath — 読み込む対象ファイルの絶対パス
'
' 戻り値:
'   True  : 読み込み成功
'   False : 非対応拡張子またはエラー発生
'
' 拡張子判定:
'   .tsv / .txt → タブ区切り (LoadDelimitedToSheet)
'   .csv        → カンマ区切り (LoadDelimitedToSheet)
'   .xlsx / .xls / .xlsm → Excel ブックの第1シート (LoadXlsxSheetToSheet)
' ============================================================
Public Function LoadFileToSheet(filePath As String) As Boolean
    Dim ext As String
    ext = LCase(Mid(filePath, InStrRev(filePath, ".") + 1))

    Select Case ext
        Case "tsv", "txt"
            LoadFileToSheet = LoadDelimitedToSheet(filePath, vbTab)
        Case "csv"
            LoadFileToSheet = LoadDelimitedToSheet(filePath, ",")
        Case "xlsx", "xls", "xlsm"
            LoadFileToSheet = LoadXlsxSheetToSheet(filePath)
        Case Else
            LogMessage "警告: 非対応の拡張子です [" & ext & "] " & filePath
            LoadFileToSheet = False
    End Select
End Function

' ============================================================
' LoadDelimitedToSheet — 区切り文字テキストファイルをシートに読み込む（プライベート）
'
' 引数:
'   filePath  — 読み込む対象ファイルの絶対パス
'   delimiter — 区切り文字（タブ: vbTab、カンマ: ","）
'
' 戻り値:
'   True  : 読み込み成功
'   False : エラー発生
'
' シート名: FilePathToSheetName で変換したファイル名（拡張子なし・31文字以内）
' ============================================================
Private Function LoadDelimitedToSheet(filePath As String, delimiter As String) As Boolean
    On Error GoTo ErrHandler

    Dim sheetName As String
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim fileNum As Integer
    Dim lineText As String
    Dim cols() As String
    Dim lineCount As Long
    Dim maxCols As Integer
    Dim r As Long
    Dim c As Integer
    Dim dataArr() As Variant  ' 2D Variant 配列（一括書き込み用）

    sheetName = FilePathToSheetName(filePath)

    ' --- 同名シートが既に存在する場合は削除（再読み込み対応）---
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws

    ' --- 新しいシートを 集計 シートの直前に挿入 ---
    Set newSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(SH_AGGR))
    newSheet.Name = sheetName

    ' ============================================================
    ' パス1: 全行を走査して総行数と最大列数を確定する
    ' ============================================================
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    lineCount = 0
    maxCols = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineCount = lineCount + 1
        cols = Split(lineText, delimiter)
        If UBound(cols) + 1 > maxCols Then maxCols = UBound(cols) + 1
    Loop
    Close #fileNum
    fileNum = 0

    ' 空ファイルの場合は正常終了（シートは作成済み）
    If lineCount = 0 Or maxCols = 0 Then
        LoadDelimitedToSheet = True
        Exit Function
    End If

    ' ============================================================
    ' パス2: 2D Variant 配列に全データを読み込んで一括書き込み
    ' ============================================================
    ReDim dataArr(1 To lineCount, 1 To maxCols)

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    r = 1
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        cols = Split(lineText, delimiter)
        For c = 0 To UBound(cols)
            dataArr(r, c + 1) = cols(c)
        Next c
        r = r + 1
    Loop
    Close #fileNum
    fileNum = 0

    ' NumberFormat="@"(テキスト書式)を先に設定して先頭ゼロの自動変換を防止
    With newSheet.Range(newSheet.Cells(1, 1), newSheet.Cells(lineCount, maxCols))
        .NumberFormat = "@"
        .Value = dataArr
    End With

    LoadDelimitedToSheet = True
    Exit Function

ErrHandler:
    Application.DisplayAlerts = True
    If fileNum > 0 Then Close #fileNum
    LoadDelimitedToSheet = False
End Function

' ============================================================
' LoadXlsxSheetToSheet — Excel ブックの第1シートをワークブックに読み込む（プライベート）
'
' 引数:
'   filePath — 読み込む .xlsx / .xls / .xlsm ファイルの絶対パス
'
' 戻り値:
'   True  : 読み込み成功
'   False : エラー発生
'
' 動作:
'   対象ブックを読み取り専用で開き、第1シートのデータを Variant 配列で
'   取得してからブックを閉じる。その後このワークブックの新規シートに書き込む。
'   ブックを開いたまま操作しないことで元ファイルへの意図しない変更を防ぐ。
' ============================================================
Private Function LoadXlsxSheetToSheet(filePath As String) As Boolean
    Dim sheetName As String
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Integer
    Dim srcData As Variant

    On Error GoTo ErrHandler
    Set srcWb = Nothing

    sheetName = FilePathToSheetName(filePath)

    ' --- 同名シートが既に存在する場合は削除 ---
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            ws.Delete
            Exit For
        End If
    Next ws
    Application.DisplayAlerts = True

    ' --- コピー先の新シートを作成（集計シートの直前）---
    Set destWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(SH_AGGR))
    destWs.Name = sheetName

    ' --- ソースブックを読み取り専用で開く ---
    ' UpdateLinks:=False で外部リンクの更新ダイアログを抑制する
    Set srcWb = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = srcWb.Sheets(1)

    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

    If lastRow >= 1 And lastCol >= 1 Then
        ' データを Variant 配列に一括取得してから元ブックを閉じる
        srcData = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol)).Value
        srcWb.Close SaveChanges:=False
        Set srcWb = Nothing

        ' コピー先シートに書き込み（テキスト書式で先頭ゼロを保護）
        With destWs.Range(destWs.Cells(1, 1), destWs.Cells(lastRow, lastCol))
            .NumberFormat = "@"
            .Value = srcData
        End With
    Else
        srcWb.Close SaveChanges:=False
        Set srcWb = Nothing
    End If

    LoadXlsxSheetToSheet = True
    Exit Function

ErrHandler:
    Application.DisplayAlerts = True
    If Not srcWb Is Nothing Then
        On Error Resume Next
        srcWb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    LoadXlsxSheetToSheet = False
End Function

' ============================================================
' FilePathToSheetName — ファイルパスを有効なシート名に変換する（プライベート）
'
' 引数:
'   filePath — 変換元のファイルパス（絶対・相対どちらでも可）
'
' 戻り値:
'   拡張子なしのファイル名から Excel のシート名として無効な文字
'   (\ / ? * [ ] :) を "_" に置換し、31文字に切り詰めた文字列。
' ============================================================
Private Function FilePathToSheetName(filePath As String) As String
    Dim fileName As String
    Dim dotPos As Integer
    Dim invalids As String
    Dim i As Integer

    ' パス区切り文字(\)の最後の位置から後ろを取り出してファイル名を取得
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    ' 拡張子を除去（最後の "." 以降を削除）
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then fileName = Left(fileName, dotPos - 1)

    ' Excel シート名として使用できない文字を "_" に置換
    invalids = "\/?*[]:"
    For i = 1 To Len(invalids)
        fileName = Join(Split(fileName, Mid(invalids, i, 1)), "_")
    Next i

    ' Excel のシート名上限(31文字)に切り詰め
    If Len(fileName) > 31 Then fileName = Left(fileName, 31)
    FilePathToSheetName = fileName
End Function
