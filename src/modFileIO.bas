Attribute VB_Name = "modFileIO"
Option Explicit

' ============================================================
' modFileIO — ファイル入出力モジュール
'
' 役割:
'   ・ユーザーに TSV ファイルを複数選択させる SelectFiles を提供する。
'   ・選択された TSV ファイルを1ファイルにつき1シートとして
'     ワークブックに読み込む LoadTsvToSheet を提供する。
'
' 設計方針:
'   ・TSV 読み込みは2パス方式を採用する。
'     パス1: 全行を走査して行数・最大列数を取得（配列サイズ確定）
'     パス2: 2D Variant 配列に一括格納し、レンジに一括書き込み
'     → セルへの逐次書き込みを避けることで大量データの処理を高速化する。
'   ・NumberFormat = "@"（テキスト書式）を書き込み前に設定することで、
'     "00123" のような先頭ゼロを持つ製品コードが数値変換されるのを防ぐ。
'   ・同名シートが既に存在する場合は削除してから再作成する（再読み込み対応）。
'   ・新シートは 集計 シートの直前に挿入し、シート順を一定に保つ。
' ============================================================

' ============================================================
' SelectFiles — TSV ファイルの複数選択ダイアログを表示する
'
' 戻り値:
'   ・ファイルが選択された場合: 選択ファイルパスの Variant 配列 (1始まりインデックス)
'   ・キャンセルされた場合    : Boolean の False
'
' 呼び出し元では VarType(戻り値) = vbBoolean でキャンセル判定する。
' ============================================================
Public Function SelectFiles() As Variant
    SelectFiles = Application.GetOpenFilename( _
        FileFilter:="テキストファイル (*.txt;*.tsv),*.txt;*.tsv", _
        Title:="読み込むTSVファイルを選択してください", _
        MultiSelect:=True)
End Function

' ============================================================
' LoadTsvToSheet — TSV ファイルをワークブックの1シートに読み込む
'
' 引数:
'   filePath — 読み込む TSV ファイルの絶対パス
'
' 戻り値:
'   True  : 読み込み成功
'   False : エラー発生（ファイルが開けない場合など）
'
' シート名:
'   ファイル名（拡張子なし）から無効文字を除去し、31文字以内に切り詰めた文字列。
'   FilePathToSheetName 関数が変換を担当する。
' ============================================================
Public Function LoadTsvToSheet(filePath As String) As Boolean
    On Error GoTo ErrHandler

    Dim sheetName As String
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim fileNum As Integer
    Dim lineText As String
    Dim cols() As String
    Dim lines() As String
    Dim lineCount As Long   ' TSV ファイルの総行数
    Dim maxCols As Integer  ' TSV ファイルの最大列数
    Dim r As Long
    Dim c As Integer
    Dim dataArr() As Variant  ' 2D Variant 配列（一括書き込み用）

    sheetName = FilePathToSheetName(filePath)

    ' --- 同名シートが既に存在する場合は削除（再読み込み対応）---
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            Application.DisplayAlerts = False  ' 削除確認ダイアログを抑制
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws

    ' --- 新しいシートを 集計 シートの直前に挿入 ---
    ' Before:=集計 を指定することでシート順(main,Config,all,[TSVシート群],集計)を維持する
    Set newSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(SH_AGGR))
    newSheet.Name = sheetName

    ' ============================================================
    ' パス1: 全行を走査して総行数と最大列数を確定する
    ' 先に配列サイズを確定することで、後の ReDim を1回に抑える。
    ' ============================================================
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    lineCount = 0
    maxCols = 0
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineCount = lineCount + 1
        cols = Split(lineText, vbTab)
        If UBound(cols) + 1 > maxCols Then maxCols = UBound(cols) + 1
    Loop
    Close #fileNum
    fileNum = 0  ' ErrHandler での多重クローズを防ぐためにリセット

    ' 空ファイルの場合は正常終了（シートは作成済み）
    If lineCount = 0 Or maxCols = 0 Then
        LoadTsvToSheet = True
        Exit Function
    End If

    ' ============================================================
    ' パス2: 2D Variant 配列に全データを読み込む
    ' ============================================================
    ReDim dataArr(1 To lineCount, 1 To maxCols)

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    r = 1
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        cols = Split(lineText, vbTab)
        For c = 0 To UBound(cols)
            dataArr(r, c + 1) = cols(c)
        Next c
        r = r + 1
    Loop
    Close #fileNum
    fileNum = 0

    ' --- 書き込み前にテキスト書式を設定（先頭ゼロの数値変換を防止）---
    ' --- 2D 配列を一括書き込みでセルへの逐次アクセスを回避 ---
    With newSheet.Range(newSheet.Cells(1, 1), newSheet.Cells(lineCount, maxCols))
        .NumberFormat = "@"  ' テキスト書式: "00123" が 123 に変換されるのを防ぐ
        .Value = dataArr     ' 一括書き込み
    End With

    LoadTsvToSheet = True
    Exit Function

ErrHandler:
    Application.DisplayAlerts = True
    If fileNum > 0 Then Close #fileNum  ' エラー時もファイルを確実にクローズ
    LoadTsvToSheet = False
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
