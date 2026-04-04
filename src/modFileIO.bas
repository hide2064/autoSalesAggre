Attribute VB_Name = "modFileIO"
Option Explicit

Public Function SelectFiles() As Variant
    ' Returns Variant array of paths, or Boolean False if cancelled
    SelectFiles = Application.GetOpenFilename( _
        FileFilter:="テキストファイル (*.txt;*.tsv),*.txt;*.tsv", _
        Title:="読み込むTSVファイルを選択してください", _
        MultiSelect:=True)
End Function

Public Function LoadTsvToSheet(filePath As String) As Boolean
    ' Creates or replaces a sheet named after the file. Returns True on success.
    On Error GoTo ErrHandler

    Dim sheetName As String
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim fileNum As Integer
    Dim lineText As String
    Dim cols() As String
    Dim lines() As String
    Dim lineCount As Long
    Dim maxCols As Integer
    Dim r As Long
    Dim c As Integer
    Dim dataArr() As Variant

    sheetName = FilePathToSheetName(filePath)

    ' Delete existing sheet with same name
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws

    ' Insert before 集計 sheet to keep sheet order consistent
    Set newSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(SH_AGGR))
    newSheet.Name = sheetName

    ' Pass 1: read all lines and find dimensions
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
    fileNum = 0

    If lineCount = 0 Or maxCols = 0 Then
        LoadTsvToSheet = True
        Exit Function
    End If

    ' Pass 2: read into 2D Variant array
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

    ' Set text format on entire range first, then bulk-write values
    With newSheet.Range(newSheet.Cells(1, 1), newSheet.Cells(lineCount, maxCols))
        .NumberFormat = "@"
        .Value = dataArr
    End With

    LoadTsvToSheet = True
    Exit Function

ErrHandler:
    Application.DisplayAlerts = True
    If fileNum > 0 Then Close #fileNum
    LoadTsvToSheet = False
End Function

Private Function FilePathToSheetName(filePath As String) As String
    ' Extract filename without extension; strip invalid sheet name chars; truncate to 31 chars
    Dim fileName As String
    Dim dotPos As Integer
    Dim invalids As String
    Dim i As Integer

    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then fileName = Left(fileName, dotPos - 1)

    ' Remove characters invalid for sheet names: \ / ? * [ ] :
    invalids = "\/?*[]:"
    For i = 1 To Len(invalids)
        fileName = Join(Split(fileName, Mid(invalids, i, 1)), "_")
    Next i

    If Len(fileName) > 31 Then fileName = Left(fileName, 31)
    FilePathToSheetName = fileName
End Function
