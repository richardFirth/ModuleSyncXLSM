Attribute VB_Name = "ZZZ_StringArrayFunctions_6"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.9
'$*DATE*3/9/2018*xxx
'$*ID*StringArrayFunctions
'$*CharCount*10518*xxx
'$*RowCount*327*xxxx

'/T--ZZZ_StringArrayFunctions_6-------------------------------------------------------------------\
' Function Name            | Return    |  Description                                             |
'--------------------------|-----------|----------------------------------------------------------|
'ConcatenateArrays         | String()  | concatentates two string arrays                          |
'insertArray               | String()  | concatentates two string arrays                          |
'AddToStringArray          | String()  | add to string array                                      |
'arrayHasStuff             | Boolean   | returns false when array is not initialized with data    |
'stringInArray             | Boolean   | is the string found in the array                         |
'stringInArray_StartsWith  | Boolean   |  is the left half of the string found in the array       |
'printStringArrToAlert     | Void      | prints all values in array to alert                      |
'showStrArr                | Void      |  prints all values in array to new workbook in column 1  |
'printStringArrToColumn    | Void      | prints all values in array to column                     |
'getArrayFromColumn        | String()  | puts all values in column to array                       |
'----- Set functions------------------------------------------------------------------------------|
'isSubset                  | Boolean   | returns true if all values are found                     |
'newStrings                | String()  | returns strings in newset that are not in oldset         |
'DifferenceBetweenSets     | String()  | returns strings in newset that are not in oldset         |
'intersectionOfStrings     | String()  | returns strings present in both arrays                   |
'----- Formatting strings-------------------------------------------------------------------------|
'TrimAndCleanArray         | String()  | trims and cleans all elements in array                   |
'CleanArray                | String()  |  cleans all elements in array                            |
'TrimArray                 | String()  |  trims all elements in array                             |
'removeBlanksFromArray     | String()  |  removes Blanks From an Array                            |
'removeDupesStringArray    | String()  | removes all identical values                             |
'\------------------------------------------------------------------------------------------------/

Option Explicit

Public Function ConcatenateArrays(theArray1() As String, theArray2() As String) As String()
'concatentates two string arrays
Dim newArr() As String
Dim n As Long: n = 1
Dim x As Long
If arrayHasStuff(theArray1) Then
    For x = LBound(theArray1) To UBound(theArray1)
        ReDim Preserve newArr(1 To n) As String
        newArr(n) = theArray1(x)
        n = n + 1
    Next x
End If

If arrayHasStuff(theArray2) Then
    For x = LBound(theArray2) To UBound(theArray2)
        ReDim Preserve newArr(1 To n) As String
        newArr(n) = theArray2(x)
        n = n + 1
    Next x
End If
ConcatenateArrays = newArr
End Function

Public Function insertArray(theTarget() As String, toInsert() As String, afterPosition As Long) As String()
'concatentates two string arrays
Dim newArr() As String
Dim n As Long: n = 1

If Not arrayHasStuff(theTarget) Then insertArray = toInsert: Exit Function
If Not arrayHasStuff(toInsert) Then insertArray = theTarget: Exit Function

If afterPosition < 0 Then GoTo insertArrayERR
If afterPosition > UBound(theTarget) Then GoTo insertArrayERR
Dim x As Long

For x = LBound(theTarget) To UBound(theTarget)
    
    ReDim Preserve newArr(1 To n) As String
    newArr(n) = theTarget(x)
    n = n + 1
    
    If x = afterPosition Then
    Dim y As Long
    For y = LBound(toInsert) To UBound(toInsert)
        ReDim Preserve newArr(1 To n) As String
        newArr(n) = toInsert(y)
        n = n + 1
    Next y
    End If

Next x

insertArray = newArr
Exit Function
insertArrayERR:
MsgBox "position out of bounds!"
End
End Function

Public Function AddToStringArray(theArray1() As String, theString As String) As String()
'add to string array
Dim newArr(1 To 1) As String
newArr(1) = theString

AddToStringArray = ConcatenateArrays(theArray1, newArr)
End Function

Public Function arrayHasStuff(theArr() As String) As Boolean
'returns false when array is not initialized with data
'https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
If (Not Not theArr) <> 0 Then arrayHasStuff = True
End Function

Public Function stringInArray(theString As String, theArray() As String) As Boolean
'is the string found in the array
If Not arrayHasStuff(theArray) Then stringInArray = False: Exit Function

Dim x As Long
For x = LBound(theArray) To UBound(theArray)
If theArray(x) = theString Then stringInArray = True: Exit Function
Next x

End Function

Public Function stringInArray_StartsWith(theString As String, theArray() As String) As Boolean
' is the left half of the string found in the array
If Not arrayHasStuff(theArray) Then stringInArray_StartsWith = False: Exit Function
Dim x As Long
For x = LBound(theArray) To UBound(theArray)
    If theArray(x) = Left(theString, Len(theArray(x))) Then stringInArray_StartsWith = True: Exit Function
Next x
End Function

Public Sub printStringArrToAlert(theArr() As String)
'prints all values in array to alert
Dim x As Long: Dim theSTR As String
If arrayHasStuff(theArr) Then
    For x = LBound(theArr) To UBound(theArr)
        theSTR = theSTR + theArr(x) & ","
    Next x
End If
MsgBox theSTR
End Sub

Public Sub showStrArr(theArr() As String, title As String)
' prints all values in array to new workbook in column 1
Dim WKBK As Workbook
Set WKBK = Workbooks.Add
Call printStringArrToColumn(theArr, WKBK.Sheets(1), 1, title)
End Sub

Public Sub printStringArrToColumn(theArr() As String, theSheet As Worksheet, theCol As Integer, theTitle As String)
'prints all values in array to column
Dim x As Long: Dim n As Long: n = 2

With theSheet
    theSheet.Cells(1, theCol).Value = theTitle
    If Not arrayHasStuff(theArr) Then Exit Sub
    For x = LBound(theArr) To UBound(theArr)
        theSheet.Cells(n, theCol).Value = theArr(x)
        n = n + 1
    Next x
End With

End Sub

Public Function getArrayFromColumn(theSheet As Worksheet, theColumn As Integer) As String()
'puts all values in column to array
Dim locSTR() As String

With theSheet
    Dim x As Long
    For x = 1 To .Cells.SpecialCells(xlCellTypeLastCell).Row
        If .Cells(x, theColumn).Value = "" Then Exit For
        ReDim Preserve locSTR(1 To x) As String
        locSTR(x) = .Cells(x, theColumn).Value
    Next x
End With

getArrayFromColumn = locSTR
End Function

'# Set functions

Public Function isSubset(subSET() As String, superset() As String) As Boolean
'returns true if all values are found
Dim x As Integer
For x = LBound(subSET) To UBound(subSET)
    If Not stringInArray(subSET(x), superset) Then isSubset = False: Exit Function
Next x
isSubset = True
End Function

Public Function newStrings(oldSet() As String, newSet() As String) As String()
'returns strings in newset that are not in oldset
MsgBox "Depreciated"
newStrings = setDifference(newSet, oldSet)
End Function

Public Function DifferenceBetweenSets(mainSet() As String, subtractThis() As String) As String()
'returns strings in newset that are not in oldset
Dim outSTR() As String
Dim n As Integer: n = 1

If Not arrayHasStuff(subtractThis) Then DifferenceBetweenSets = mainSet: Exit Function
If Not arrayHasStuff(mainSet) Then Exit Function

Dim x As Integer
For x = LBound(mainSet) To UBound(mainSet)
    If Not stringInArray(mainSet(x), subtractThis) Then
        ReDim Preserve outSTR(1 To n) As String
        outSTR(n) = mainSet(x)
        n = n + 1
    End If
Next x

DifferenceBetweenSets = outSTR

End Function

Public Function intersectionOfStrings(set1() As String, set2() As String) As String()
'returns strings present in both arrays

If Not arrayHasStuff(set1) Then Exit Function
If Not arrayHasStuff(set2) Then Exit Function

Dim fullThing() As String
fullThing = removeDupesStringArray(ConcatenateArrays(set1, set2))

Dim theIntersect() As String
Dim n As Integer: n = 1

Dim x As Integer
For x = LBound(fullThing) To UBound(fullThing)
    If stringInArray(fullThing(x), set1) And stringInArray(fullThing(x), set2) Then
    ReDim Preserve theIntersect(1 To n) As String
    theIntersect(n) = fullThing(x)
    n = n + 1
    End If
Next x

intersectionOfStrings = theIntersect
End Function

'# Formatting strings

Public Function TrimAndCleanArray(Arr() As String) As String()
'trims and cleans all elements in array
Dim loc() As String
Dim n As String: n = 1
Dim x As Integer
If Not arrayHasStuff(Arr) Then Exit Function

For x = LBound(Arr) To UBound(Arr)
ReDim Preserve loc(1 To n) As String
loc(n) = Trim(Application.WorksheetFunction.Clean(Arr(x)))
n = n + 1
Next x
TrimAndCleanArray = loc
End Function

Public Function CleanArray(Arr() As String) As String()
' cleans all elements in array
Dim loc() As String
Dim n As String: n = 1
Dim x As Integer
If Not arrayHasStuff(Arr) Then Exit Function

For x = LBound(Arr) To UBound(Arr)
ReDim Preserve loc(1 To n) As String
loc(n) = Application.WorksheetFunction.Clean(Arr(x))
n = n + 1
Next x
CleanArray = loc
End Function

Public Function TrimArray(Arr() As String) As String()
' trims all elements in array
Dim loc() As String
Dim n As String: n = 1
Dim x As Integer
If Not arrayHasStuff(Arr) Then Exit Function

For x = LBound(Arr) To UBound(Arr)
ReDim Preserve loc(1 To n) As String
loc(n) = Trim(Arr(x))
n = n + 1
Next x
TrimArray = loc
End Function

Public Function removeBlanksFromArray(Arr() As String) As String()
' removes Blanks From an Array
Dim loc() As String
Dim n As String: n = 1
Dim x As Integer
If Not arrayHasStuff(Arr) Then Exit Function

For x = LBound(Arr) To UBound(Arr)
    If Arr(x) <> "" Then
    ReDim Preserve loc(1 To n) As String
    loc(n) = Arr(x)
    n = n + 1
    End If
Next x
removeBlanksFromArray = loc
End Function

Public Function removeDupesStringArray(theArray() As String) As String()
'removes all identical values
Dim newStringArr() As String
Dim x As Long

If Not arrayHasStuff(theArray) Then Exit Function

ReDim Preserve newStringArr(1 To 1) As String
newStringArr(1) = theArray(LBound(theArray))

Dim n As Long
n = 2

For x = LBound(theArray) To UBound(theArray)
    If Not stringInArray(theArray(x), newStringArr) Then
    ReDim Preserve newStringArr(1 To n) As String
    newStringArr(n) = theArray(x)
    n = n + 1
    End If
Next x
removeDupesStringArray = newStringArr
End Function
