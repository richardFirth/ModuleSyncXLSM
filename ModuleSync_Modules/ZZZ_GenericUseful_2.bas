Attribute VB_Name = "ZZZ_GenericUseful_2"
'$VERSIONCONTROL
'$*MINOR_VERSION*2.4
'$*DATE*3/8/2018*xxx
'$*ID*GenericUseful
'$*CharCount*7859*xxxx
'$*RowCount*244*xxxx

'/T--ZZZ_GenericUseful_2------------------------------------------------------------------------------------\
' Function Name         | Return  |  Description                                                            |
'-----------------------|---------|-------------------------------------------------------------------------|
'openFolder             | Void    |  open and display a folder                                              |
'clearWorkSpace         | Void    | clears columns beween startpoint and stop point                         |
'saveWorkbookToDesktop  | Void    | saves the given workbook to the desktop                                 |
'addSheetWithName       | Void    | adds a sheet with the specific name                                     |
'complexRoutineStart    | Void    | switches off stuff to allow for faster execution                        |
'endComplex             | Void    |  visible version                                                        |
'complexRoutineEnd      | Void    |  turns stuff back on                                                    |
'hideExcelStuff         | Void    | Hides the ribbon, formula bar, statusbar, tabs, heading, and gridlines  |
'showExcelStuff         | Void    | Shows the ribbon, formula bar, sratusbar, tabs, heading, and gridlines  |
'GenerateIDcode         | String  | generates a random ID code                                              |
'readableDate           | String  | gives todays date as a string in format 7Feb18                          |
'----- Logging----------------------------------------------------------------------------------------------|
'addToDebugLog          | Void    | appends strings to a debug log                                          |
'reportError            | Void    | appends an error log to a CSV in this folder                            |
'\----------------------------------------------------------------------------------------------------------/

Option Explicit

'  Dim x As Integer
'  For x = LBound(tErrData) To UBound(tErrData)
'
'  Next x

' Dim x As Integer
' Dim n As Integer: n = 1
' For x = LBound(tErrData) To UBound(tErrData)
'     ReDim Preserve something(1 To n) As String
'     something(n) = ""
'     n = n + 1
' Next x
'
'
' Dim theResult As Boolean
' If MsgBox("Select Yes Or No", vbYesNo, "TITLE 1") = vbYes Then theResult = True

Dim logFunc() As String
Dim logMod() As String

Public Sub openFolder(tFolderPath As String)
' open and display a folder
    Shell "explorer.exe" & " " & tFolderPath, vbNormalFocus
End Sub

Public Sub clearWorkSpace(aSheet As Worksheet, startP As Integer, stopP As Integer)
'clears columns beween startpoint and stop point
Dim x As Integer
For x = startP To stopP

With aSheet.Columns(x)
.ClearContents
.Interior.Pattern = xlNone
.Interior.TintAndShade = 0
.Interior.PatternTintAndShade = 0
End With

Next x
End Sub

Public Sub saveWorkbookToDesktop(theWKBK As Workbook, theName As String)
'saves the given workbook to the desktop
theWKBK.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & theName)
End Sub

Public Sub addSheetWithName(theName As String, theBook As Workbook)
'adds a sheet with the specific name
theBook.Sheets.Add(After:=theBook.Worksheets(theBook.Worksheets.Count)).Name = theName
End Sub

Public Sub complexRoutineStart(notUsed As String)
'switches off stuff to allow for faster execution
With Application
.ShowWindowsInTaskbar = False
.ScreenUpdating = False
.Calculation = xlCalculationManual
End With
End Sub

Public Sub endComplex()
' visible version
Call complexRoutineEnd("")
End Sub

Public Sub complexRoutineEnd(notUsed As String)
' turns stuff back on
With Application
.ShowWindowsInTaskbar = True
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
End With
End Sub

Public Sub hideExcelStuff(nused As String)
'Hides the ribbon, formula bar, statusbar, tabs, heading, and gridlines

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = Not Application.DisplayStatusBar
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayGridlines = False

End Sub

Public Sub showExcelStuff(nused As String)
'Shows the ribbon, formula bar, sratusbar, tabs, heading, and gridlines
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayGridlines = True
End Sub

Public Function GenerateIDcode() As String
'generates a random ID code
Randomize

Dim a As Integer:   a = Int(9 * Rnd) + 1
Dim b As Integer:   b = Int(9 * Rnd) + 1
Dim c As Integer:   c = Int(9 * Rnd) + 1
Dim d As Integer:   d = Int(9 * Rnd) + 1

Dim y As Integer:   y = Int(26 * Rnd) + 1

Select Case y

Case 1: GenerateIDcode = "A" & a & b & c & d
Case 2: GenerateIDcode = "B" & a & b & c & d
Case 3: GenerateIDcode = "C" & a & b & c & d
Case 4: GenerateIDcode = "D" & a & b & c & d
Case 5: GenerateIDcode = "E" & a & b & c & d
Case 6: GenerateIDcode = "F" & a & b & c & d
Case 7: GenerateIDcode = "G" & a & b & c & d
Case 8: GenerateIDcode = "H" & a & b & c & d
Case 9: GenerateIDcode = "I" & a & b & c & d
Case 10: GenerateIDcode = "J" & a & b & c & d
Case 11: GenerateIDcode = "K" & a & b & c & d
Case 12: GenerateIDcode = "L" & a & b & c & d
Case 13: GenerateIDcode = "M" & a & b & c & d
Case 14: GenerateIDcode = "N" & a & b & c & d
Case 15: GenerateIDcode = "O" & a & b & c & d
Case 16: GenerateIDcode = "P" & a & b & c & d
Case 17: GenerateIDcode = "Q" & a & b & c & d
Case 18: GenerateIDcode = "R" & a & b & c & d
Case 19: GenerateIDcode = "S" & a & b & c & d
Case 20: GenerateIDcode = "T" & a & b & c & d
Case 21: GenerateIDcode = "U" & a & b & c & d
Case 22: GenerateIDcode = "V" & a & b & c & d
Case 23: GenerateIDcode = "W" & a & b & c & d
Case 24: GenerateIDcode = "X" & a & b & c & d
Case 25: GenerateIDcode = "Y" & a & b & c & d
Case 26: GenerateIDcode = "Z" & a & b & c & d

End Select

End Function

Public Function readableDate(Optional aDate As Date = 0) As String
'gives todays date as a string in format 7Feb18
If aDate = 0 Then aDate = Now()

Dim currentDay As Integer: currentDay = Day(aDate)
Dim currentMonthNumber As Integer: currentMonthNumber = month(aDate)
Dim currentYear As Integer: currentYear = Year(aDate)
Dim currentMonth As String: currentMonth = ""

Select Case currentMonthNumber
Case 1
currentMonth = "Jan"
Case 2
currentMonth = "Feb"
Case 3
currentMonth = "Mar"
Case 4
currentMonth = "Apr"
Case 5
currentMonth = "May"
Case 6
currentMonth = "Jun"
Case 7
currentMonth = "Jul"
Case 8
currentMonth = "Aug"
Case 9
currentMonth = "Sep"
Case 10
currentMonth = "Oct"
Case 11
currentMonth = "Nov"
Case 12
currentMonth = "Dec"

End Select

readableDate = currentDay & currentMonth & currentYear

End Function

'# Logging

Public Sub addToDebugLog(functionName As String, moduleName As String)
'appends strings to a debug log
If Not FileThere(ThisWorkbook.Path & "\DebugLog.csv") Then
    Dim hDat(1 To 3) As String
    hDat(1) = "Function"
    hDat(2) = "Module"
    hDat(3) = ThisWorkbook.Name
    Call appendAsRowToCSV(ThisWorkbook.Path & "\DebugLog.csv", hDat)
End If

If stringInArray(functionName, logFunc) Then
    Exit Sub
Else
    logFunc = AddToStringArray(logFunc, functionName)
End If
Dim repDat(1 To 2) As String
repDat(1) = functionName
repDat(2) = moduleName
Call appendAsRowToCSV(ThisWorkbook.Path & "\DebugLog.csv", repDat)

End Sub

Public Sub reportError(functionName As String, tSTR() As String)
'appends an error log to a CSV in this folder
Dim myErr() As String
Dim repDat(1 To 2) As String
repDat(1) = functionName
repDat(2) = Now
myErr = ConcatenateArrays(repDat, tSTR)

Call appendAsRowToCSV(ThisWorkbook.Path & "\errLog.csv", myErr)
End Sub
