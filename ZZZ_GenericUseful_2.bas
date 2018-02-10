Attribute VB_Name = "ZZZ_GenericUseful_2"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.4
'$*DATE*10Feb18
'$*ID*GenericUseful

'/---ZZZ_GenericUseful_2--------------------------------------------------------------------------\
'  Function Name         | Return          |   Description                                        |
'------------------------|-----------------|------------------------------------------------------|
' clearWorkSpace         | void            | clears columns beween startpoint and stop point      |
' saveWorkbookToDesktop  | void            | saves the given workbook to the desktop              |
' addSheetWithName       | void            | adds a sheet with the specific name                  |
' complexRoutineStart    | void            | switches off stuff to allow for faster execution     |
' complexRoutineEnd      | void            | switches the stuff back on                           |
' GenerateIDcode         | string          | generates a random ID code.                          |
' readableDate           | string          | gives todays date as a string in format 7Feb18       |
'\------------------------------------------------------------------------------------------------/



Option Explicit



Public Sub reportError(functionName As String, tStr() As String)
    Dim myERR() As String
    Dim repDat(1 To 2) As String
    repDat(1) = functionName
    repDat(2) = Now
    myERR = ConcatenateArrays(repDat, tStr)

    Call appendAsRowToCSV(ThisWorkbook.Path & "\errLog.csv", myERR)
End Sub



Sub clearWorkSpace(aSheet As Worksheet, startP As Integer, stopP As Integer)
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
    theWKBK.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & theName)
End Sub


Public Sub addSheetWithName(theName As String, theBook As Workbook)
   theBook.Sheets.Add(After:=theBook.Worksheets(theBook.Worksheets.Count)).name = theName
End Sub



Public Sub complexRoutineStart(notUsed As String)
    With Application
        .ShowWindowsInTaskbar = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Sub

Public Sub endComplex()
    Call complexRoutineEnd("")
End Sub


Public Sub complexRoutineEnd(notUsed As String)
    With Application
        .ShowWindowsInTaskbar = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Public Function GenerateIDcode() As String

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










