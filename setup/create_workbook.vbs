Option Explicit

' ============================================================
' autoSalesAggre workbook setup script
' Usage: cscript setup\create_workbook.vbs
' Prereq: Enable "Trust access to the VBA project object model"
'         in Excel Trust Center settings before running.
' ============================================================

Dim fso, scriptDir, srcPath, outputFile
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
srcPath   = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\src")) & "\"
outputFile = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, "..\autoSalesAggre.xlsm"))

WScript.Echo "Setup started"
WScript.Echo "Source : " & srcPath
WScript.Echo "Output : " & outputFile

Dim xlApp, wb
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set wb = xlApp.Workbooks.Add

' ---- Create 4 sheets: main, Config, all, Shuukei (placeholder for 集計) ----
' Remove extra default sheets first
Do While wb.Sheets.Count > 1
    wb.Sheets(wb.Sheets.Count).Delete
Loop
wb.Sheets(1).Name = "main"
wb.Sheets.Add(After:=wb.Sheets("main")).Name = "Config"
wb.Sheets.Add(After:=wb.Sheets("Config")).Name = "all"
wb.Sheets.Add(After:=wb.Sheets("all")).Name = "Shuukei"

' ---- Import VBA modules ----
WScript.Echo "Importing VBA modules..."
Dim vbp
Set vbp = wb.VBProject

vbp.VBComponents.Import srcPath & "modConfig.bas"
vbp.VBComponents.Import srcPath & "modFileIO.bas"
vbp.VBComponents.Import srcPath & "modDataProcess.bas"
vbp.VBComponents.Import srcPath & "modAggregation.bas"
vbp.VBComponents.Import srcPath & "modUIControl.bas"
vbp.VBComponents.Import srcPath & "modSetup.bas"
WScript.Echo "Modules imported."

' ---- Run one-time setup (renames sheets, sets layouts, injects event) ----
WScript.Echo "Running InitWorkbook..."
xlApp.Run "modSetup.InitWorkbook"
WScript.Echo "InitWorkbook complete."

' ---- Save as .xlsm (52 = xlOpenXMLMacroEnabled) ----
wb.SaveAs outputFile, 52
WScript.Echo "Saved: " & outputFile

xlApp.Quit
Set xlApp = Nothing

WScript.Echo ""
WScript.Echo "Done. Open autoSalesAggre.xlsm in Excel and enable macros."
