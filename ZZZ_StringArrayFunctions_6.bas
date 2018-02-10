Attribute VB_Name = "ZZZ_StringArrayFunctions_6"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.5
'$*DATE*8Feb18
'$*ID*StringArrayFunctions


Option Explicit

'/----ZZZ_StringArrayFunctions_6-----------------------------------------------------------------------\
'  Function Name          | Return          |   Description                                            |
'-------------------------|-----------------|----------------------------------------------------------|
' ConcatenateArrays       | String()        | concatentates two string arrays                          |
' AddToStringArray        | String()        | adds a single string to an array                         |
' removeDupesStringArray  | String()        | removes all identical values                             |
' stringInArray           | boolean         | is the string found in the array                         |
' stringInArray_StartsWith| boolean         | string must match left values of a value in array        |
' printStringArrToAlert   | void            | prints all values in array to alert                      |
' showStrArr              | void            | prints all values in array to new workbook in column 1   |
' printStringArrToColumn  | void            | prints all values in array to column                     |
' arrayHasStuff           | boolean         | returns false when array has nothing                     |
' getArrayFromColumn      | String()        | puts all values in column to array                       |
' isSubset                | boolean         | returns true if all values are found                     |
' DifferenceBetweenSets   | String()        | subtracts a set from another set                         |
' intersectionOfStrings   | String()        | returns strings present in both arrays                   |
' TrimAndCleanArray       | String()        | trim and clean each element in array                     |
' removeBlanksFromArray   | String()        | remove Blanks From Array                                 |
'\----------------------------------------------------------------------------------------------------/


 ' /----------------------------------------\
 ' |concatentates two string arrays         |
 ' \----------------------------------------/
Public Function ConcatenateArrays(theArray1() As String, theArray2() As String) As String()

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


 ' /----------------------------------------\
 ' |add to string array                     |
 ' \----------------------------------------/
Public Function AddToStringArray(theArray1() As String, theString As String) As String()
    Dim newArr(1 To 1) As String
    newArr(1) = theString
    
    AddToStringArray = ConcatenateArrays(theArray1, newArr)
End Function




 ' /-------------------------------------\
 ' |removes all identical values         |
 ' \-------------------------------------/
Public Function removeDupesStringArray(theArray() As String) As String()
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
    Next
removeDupesStringArray = newStringArr
End Function


 ' /-----------------------------------------\
 ' |is the string found in the array         |
 ' \-----------------------------------------/
Public Function stringInArray(theString As String, theArray() As String) As Boolean
    If Not arrayHasStuff(theArray) Then stringInArray = False: Exit Function

    Dim x As Long
    For x = LBound(theArray) To UBound(theArray)
        If theArray(x) = theString Then stringInArray = True: Exit Function
    Next x

End Function


 ' /----------------------------------------------------------\
 ' |is the left half of the string found in the array         |
 ' \----------------------------------------------------------/
Public Function stringInArray_StartsWith(theString As String, theArray() As String) As Boolean
    If Not arrayHasStuff(theArray) Then stringInArray_StartsWith = False: Exit Function

    Dim x As Long
    For x = LBound(theArray) To UBound(theArray)
        If theArray(x) = Left(theString, Len(theArray(x))) Then stringInArray_StartsWith = True: Exit Function
    Next x

End Function




 ' /----------------------------------------\
 ' |prints all values in array to alert     |
 ' \----------------------------------------/
Public Sub printStringArrToAlert(theArr() As String)
    Dim x As Long
    Dim theSTR As String
    
    If arrayHasStuff(theArr) Then
        For x = LBound(theArr) To UBound(theArr)
            theSTR = theSTR + theArr(x) & ","
        Next x
    End If
    
    MsgBox theSTR
End Sub

 ' /---------------------------------------------------------------\
 ' |prints all values in array to new workbook in column 1         |
 ' \---------------------------------------------------------------/
Public Sub showStrArr(theArr() As String, title As String)

Dim WKBK As Workbook
Set WKBK = Workbooks.Add
Call printStringArrToColumn(theArr, WKBK.Sheets(1), 1, title)

End Sub

 ' /---------------------------------------------\
 ' |prints all values in array to column         |
 ' \---------------------------------------------/
Public Sub printStringArrToColumn(theArr() As String, theSheet As Worksheet, theCol As Integer, theTitle As String)
    
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



 ' /-----------------------------------------\
 ' |returns false when array has nothing     |
 ' \-----------------------------------------/
Public Function arrayHasStuff(theArr() As String) As Boolean
'https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
    If (Not Not theArr) <> 0 Then arrayHasStuff = True
End Function


 ' /----------------------------------------\
 ' |puts all values in column to array      |
 ' \----------------------------------------/
Public Function getArrayFromColumn(theSheet As Worksheet, theColumn As Integer) As String()
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


' set operations




 ' /---------------------------------------------\
 ' |returns true if all values are found         |
 ' \---------------------------------------------/
Public Function isSubset(subSET() As String, superset() As String) As Boolean
Dim x As Integer

For x = LBound(subSET) To UBound(subSET)
    If Not stringInArray(subSET(x), superset) Then isSubset = False: Exit Function
Next x

isSubset = True
End Function


 ' /--------------------------------------------------\
 ' |returns strings in newset that are not in oldset  |
 ' \--------------------------------------------------/
Public Function newStrings(oldSet() As String, newSet() As String) As String()
    MsgBox "Depreciated"
newStrings = setDifference(newSet, oldSet)
End Function


 ' /--------------------------------------------------\
 ' |returns strings in newset that are not in oldset  |
 ' \--------------------------------------------------/
Public Function DifferenceBetweenSets(mainSet() As String, subtractThis() As String) As String()
Dim outSTR() As String
Dim n As Integer: n = 1

If Not arrayHasStuff(subtractThis) Then DifferenceBetweenSets = mainSet: Exit Function
If Not arrayHasStuff(mainSet) Then Exit Function

Dim x As Integer
    For x = LBound(mainSet) To UBound(mainSet)
        If stringInArray(mainSet(x), subtractThis) Then
        Else
            ReDim Preserve outSTR(1 To n) As String
            outSTR(n) = mainSet(x)
            n = n + 1
        End If
    Next x

DifferenceBetweenSets = outSTR

End Function

 ' /--------------------------------------------------\
 ' |returns strings present in both arrays            |
 ' \--------------------------------------------------/
Public Function intersectionOfStrings(set1() As String, set2() As String) As String()

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


 ' /--------------------------------------------------\
 ' |trims all elements in arraa            |
 ' \--------------------------------------------------/

Public Function TrimAndCleanArray(Arr() As String) As String()
    Dim LOC() As String
    Dim n As String: n = 1
    Dim x As Integer
    If Not arrayHasStuff(Arr) Then Exit Function
    
    For x = LBound(Arr) To UBound(Arr)
        ReDim Preserve LOC(1 To n) As String
        LOC(n) = Trim(Application.WorksheetFunction.Clean(Arr(x)))
        n = n + 1
    Next x
    TrimAndCleanArray = LOC
End Function


 ' /--------------------------------------------------\
 ' |removes Blanks From an Array                      |
 ' \--------------------------------------------------/

Public Function removeBlanksFromArray(Arr() As String) As String()
    Dim LOC() As String
    Dim n As String: n = 1
    Dim x As Integer
    If Not arrayHasStuff(Arr) Then Exit Function
    
    For x = LBound(Arr) To UBound(Arr)
        If Arr(x) <> "" Then
            ReDim Preserve LOC(1 To n) As String
            LOC(n) = Arr(x)
            n = n + 1
        End If
    Next x
    removeBlanksFromArray = LOC
End Function
