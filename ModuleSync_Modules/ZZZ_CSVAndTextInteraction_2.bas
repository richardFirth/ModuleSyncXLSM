Attribute VB_Name = "ZZZ_CSVAndTextInteraction_2"
'$VERSIONCONTROL
'$*MINOR_VERSION*2.0
'$*DATE*3/2/2018*xxx
'$*ID*CSVAndTextInteraction
'$*CharCount*6695*xxxx
'$*RowCount*220*xxxx

'/T--ZZZ_CSVAndTextInteraction_2----------------------------------------------------------------------------------------------------\
' Function Name                 | Return         |  Description                                                                     |
'-------------------------------|----------------|----------------------------------------------------------------------------------|
'string2wrapper                 | StringWrapper  |  feed in string arrays to get an array of wrappers that hold the string arrays.  |
'appendToCSV                    | Boolean        |  appends a string array to CSV file as a row                                     |
'appendAsRowToCSV               | Boolean        |                                                                                  |
'appendWrappedDataToCSV         | Boolean        |  feed stringrappers in. each stringwrapper is written as a row                   |
'getCSVFromFile                 | String()       |  get a csv file as a string array                                                |
'convertTXTDocumentToStringArr  | String()       |  get a text document as a string array                                           |
'convertTXTDocumentToString     | String         |  get a text document as a string                                                 |
'getTxTDocumentAsString         | String()       |  get a text document as a string                                                 |
'createTextFromStringArr        | Void           |  spawns a text file containing the string array                                  |
'createTextFromString           | Void           |  creates a text document containing a single string                              |
'createFile                     | Boolean        |  creates a file                                                                  |
'\----------------------------------------------------------------------------------------------------------------------------------/

Option Explicit

Public Type StringWrapper
theSTR() As String
End Type

Public seenError As Boolean

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Public Function string2wrapper(mySTR() As String) As StringWrapper
' feed in string arrays to get an array of wrappers that hold the string arrays.
Dim mySW As StringWrapper
mySW.theSTR = mySTR
string2wrapper = mySW
End Function

Function appendToCSV(filePath As String, contents() As String) As Boolean
' appends a string array to CSV file as a row
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
' feed stringrappers in. each stringwrapper is written as a row
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
' get a csv file as a string array
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

Function convertTXTDocumentToStringArr(thePath As String) As String()
' get a text document as a string array
Dim theFileContents As String

On Error GoTo convTXTerr
theFileContents = CreateObject("Scripting.FileSystemObject").GetFile(thePath).OpenAsTextStream(ForReading, TristateUseDefault).ReadAll

Dim docFeed() As String
docFeed = Split(theFileContents, Chr(10))
convertTXTDocumentToStringArr = docFeed

Exit Function
convTXTerr:

End Function

Function convertTXTDocumentToString(thePath As String) As String
' get a text document as a string
Dim theFileContents As String

On Error GoTo convTXTstrerr
theFileContents = CreateObject("Scripting.FileSystemObject").GetFile(thePath).OpenAsTextStream(ForReading, TristateUseDefault).ReadAll

convertTXTDocumentToString = theFileContents

Exit Function
convTXTstrerr:

End Function

Function getTxTDocumentAsString(thePath As String) As String()
' get a text document as a string
Dim theFileContents As String

On Error GoTo getTXTerr
theFileContents = CreateObject("Scripting.FileSystemObject").GetFile(thePath).OpenAsTextStream(ForReading, TristateUseDefault).ReadAll

Dim docFeed() As String
docFeed = Split(theFileContents, Chr(10))
getTxTDocumentAsString = TrimAndCleanArray(docFeed)

' otherwise weird stuff messes up your array
' getTxTDocumentAsString = docFeed
Dim ERRSTR(1 To 2) As String
ERRSTR(1) = "OLD METHOD!"
Call reportError("getTxTDocumentAsString", ERRSTR)

Exit Function
getTXTerr:

End Function

Sub createTextFromStringArr(theContents() As String, theFullPathName As String)
' spawns a text file containing the string array
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
' creates a text document containing a single string
Dim fs As Object, f As Object
Set fs = CreateObject("Scripting.FileSystemObject") ' this creates the fileSystemObject object for all file operations

Call createFile(f, fs, theFullPathName)

f.WriteLine theContents

End Sub

Function createFile(fileobject As Object, FSO As Object, fileName As String) As Boolean
' creates a file
On Error GoTo createFileError

Set fileobject = FSO.CreateTextFile(fileName, True)

Exit Function
createFileError:
MsgBox "Problem: " & fileName
End Function

