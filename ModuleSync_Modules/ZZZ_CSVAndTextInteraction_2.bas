Attribute VB_Name = "ZZZ_CSVAndTextInteraction_2"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.7
'$*DATE*21Feb18
'$*ID*CSVAndTextInteraction
'$*CharCount*4625*xxxx
'$*RowCount*212*xxxx

Option Explicit

'/---ZZZ_CSVAndTextInteraction_2---------updated 21Feb18------------------------------------------\
'  Function Name                   | Return          |   Description                             |
'----------------------------------|-----------------|-------------------------------------------|
' appendAsRowToCSV                 | Boolean         | appends a string to a CSV                 |
' appendWrappedDataToCSV           | Boolean         |
' getCSVFromFile                   | String()        | gets a CSV file into a string             |
' getTxTDocumentAsString           | String()        | gets a text document into a string        |
' createTextFromStringArr          | void            | saves a string into a text document       |
' createFile                       | Boolean         | creates a file                            |
'\-----------------------------------------------------------------------------------------------/


'Function getTxTDocumentAsString(thePath As String) As String()
'Sub createTextFromStringArr(theContents() As String, theFullPathName As String)
'Function createFile(fileobject As Object, FSO As Object, fileName As String) As Boolean


Public Type StringWrapper
theSTR() As String
End Type


Public seenError As Boolean

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0




Public Function string2wrapper(mySTR() As String) As StringWrapper
Dim mySW As StringWrapper
mySW.theSTR = mySTR
string2wrapper = mySW
End Function



' each string array is appended as a row
Function appendToCSV(filePath As String, contents() As String) As Boolean
If Not seenError Then MsgBox "Function deprectiated: appendToCSV"
seenError = True
appendToCSV = appendAsRowToCSV(filePath, contents)
End Function



Function appendAsRowToCSV(filePath As String, theRow() As String) As Boolean

Dim x As Integer
On Error GoTo BADappendtoCSV
Open filePath For Append As #1
For x = LBound(theRow) To UBound(theRow)
If x = UBound(theRow) Then
Write #1, theRow(x)
Else
Write #1, theRow(x),
End If
Next x
Close #1
appendAsRowToCSV = True

Exit Function
BADappendtoCSV:
appendAsRowToCSV = False
End Function


Function appendWrappedDataToCSV(filePath As String, theRowDat() As StringWrapper) As Boolean

Dim x As Integer
Dim y As Long
On Error GoTo BADappendDCVToCSV
Open filePath For Append As #1

For y = LBound(theRowDat) To UBound(theRowDat)

With theRowDat(y)
For x = LBound(.theSTR) To UBound(.theSTR)
If x = UBound(.theSTR) Then
Write #1, .theSTR(x)
Else
Write #1, .theSTR(x),
End If
Next x
End With

Next y


Close #1



appendWrappedDataToCSV = True

Exit Function
BADappendDCVToCSV:
appendWrappedDataToCSV = False
End Function






Function getCSVFromFile(filePath As String) As String()

Dim csvRow() As String
Dim dataIN As String
Dim x As Integer: x = 1

Open filePath For Input As #4

Do While True
On Error GoTo BADgetCSVFromFile
'Input #4, dataIN
Line Input #4, dataIN
ReDim Preserve csvRow(1 To x) As String
csvRow(x) = dataIN
x = x + 1
Loop

finishedA:

Close #4

getCSVFromFile = csvRow

Exit Function
BADgetCSVFromFile:

Resume finishedA

End Function



Function getTxTDocumentAsString(thePath As String) As String()
Dim theFileContents As String

On Error GoTo getTXTerr
theFileContents = CreateObject("Scripting.FileSystemObject").GetFile(thePath).OpenAsTextStream(ForReading, TristateUseDefault).ReadAll

Dim docFeed() As String
docFeed = Split(theFileContents, Chr(10))
getTxTDocumentAsString = TrimAndCleanArray(docFeed) ' otherwise weird stuff messes up your array
'getTxTDocumentAsString = docFeed
Exit Function
getTXTerr:

End Function


Sub createTextFromStringArr(theContents() As String, theFullPathName As String)

Dim fs As Object, f As Object
Set fs = CreateObject("Scripting.FileSystemObject") ' this creates the fileSystemObject object for all file operations

Call createFile(f, fs, theFullPathName)

Dim n As Integer
If arrayHasStuff(theContents) Then
For n = LBound(theContents) To UBound(theContents)
f.WriteLine theContents(n)
Next n
End If

End Sub




Sub createTextFromString(theContents As String, theFullPathName As String)

Dim fs As Object, f As Object
Set fs = CreateObject("Scripting.FileSystemObject") ' this creates the fileSystemObject object for all file operations

Call createFile(f, fs, theFullPathName)

f.WriteLine theContents

End Sub



Function createFile(fileobject As Object, FSO As Object, fileName As String) As Boolean
On Error GoTo createFileError

Set fileobject = FSO.CreateTextFile(fileName, True)

Exit Function
createFileError:
MsgBox "Problem: " & fileName
End Function






