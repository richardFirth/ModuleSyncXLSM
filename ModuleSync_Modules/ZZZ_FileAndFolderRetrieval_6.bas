Attribute VB_Name = "ZZZ_FileAndFolderRetrieval_6"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.5
'$*DATE*3/8/2018*xxx
'$*ID*FileAndFolderRetrieval
'$*CharCount*12284*xxx
'$*RowCount*385*xxxx

'/T--ZZZ_FileAndFolderRetrieval_6----------------------------------------------------------------------------------------------------\
' Function Name                 | Return         |  Description                                                                      |
'-------------------------------|----------------|-----------------------------------------------------------------------------------|
'getFilesAndFoldersArray        | String()       | gets files and folders as string array                                            |
'getFilesInFolder2Array         | String()       | gets files as string array                                                        |
'getFilePathsInFolder2Array     | String()       | gets file path as string array                                                    |
'getFolderInFolder2Array        | String()       | gets folders as string array                                                      |
'getDetailFolderTree            | PathAndName()  | gets the pathsAndNames of all levels of folders to an array of pathandName   |    |
'getNextLevel                   | PathAndName()  |  gets one level of pathsAndNames of all items in the folders in thecurrent level  |
'DetailgetFilesAndFoldersArray  | PathAndName()  | gets all files and folders in the folder in thecurrent level                      |
'DetailFilesInFolder2Array      | PathAndName()  |  gets all files in the folder                                                     |
'DetailFolderInFolder2Array     | PathAndName()  |  gets all folders in the folder                                                   |
'ConcatenatePathAndName         | PathAndName()  |   concatentaes two pathAndName arrays                                             |
'PathAndNameArrayHasStuff       | Boolean        | true when there is data in array                                                  |
'printPathAndNameToColumn       | Void           | prints detail to sheet                                                            |
'namesFromPaths                 | String()       | returns filenames given paths                                                     |
'nameFromPath                   | String         | returns filename given path                                                       |
'pathFromName                   | String         | returns path given full path                                                      |
'getPathFromPandN               | String()       |  retrieve the paths from the path and name type                                   |
'ShowHideFilesInFolder          | Void           |  shows or hides files in folder                                                   |
'~~showHideFileFolder           | Void           |  helper                                                                           |
'\-----------------------------------------------------------------------------------------------------------------------------------/

Option Explicit

Type PathAndName
A_Path As String
B_Name As String
C_DateMod As Date
D_isFolder As Boolean
End Type

Public Enum ShowHideConfig
A_ShowAll
B_HideAll
C_ShowAllExcept
D_HideAllExcept
E_HideThese
End Enum

Public Function getFilesAndFoldersArray(theFolder As String) As String()
'gets files and folders as string array
getFilesAndFoldersArray = ConcatenateArrays(getFilesInFolder2Array(theFolder), getFolderInFolder2Array(theFolder))
End Function

Public Function getFilesInFolder2Array(theFolder As String) As String()
'gets files as string array
getFilesInFolder2Array = namesFromPaths(getFilePathsInFolder2Array(theFolder))
End Function

Public Function getFilePathsInFolder2Array(theFolder As String) As String()
'gets file path as string array
Dim findToolsFSO As Object
Set findToolsFSO = CreateObject("Scripting.FileSystemObject") 'Create an instance of the FileSystemObject
Dim objFolder_CM As Object
Set objFolder_CM = findToolsFSO.GetFolder(theFolder) 'Get the folder object

'loops through each file in the directory and prints their names and path

Dim x As Integer

x = 1
Dim locFiles() As String
Dim objFile_CM As Object

For Each objFile_CM In objFolder_CM.Files
ReDim Preserve locFiles(1 To x) As String
locFiles(x) = objFile_CM.Path
x = x + 1
Next objFile_CM

getFilePathsInFolder2Array = locFiles

End Function

Public Function getFolderInFolder2Array(theFolder As String) As String()
'gets folders as string array
Dim findToolsFSO As Object
Set findToolsFSO = CreateObject("Scripting.FileSystemObject") 'Create an instance of the FileSystemObject

Dim objFolder_CM As Object
Set objFolder_CM = findToolsFSO.GetFolder(theFolder) 'Get the folder object

'loops through each file in the directory and prints their names and path

Dim x As Integer

x = 1
Dim locFolders() As String
Dim objFolderB_CM As Object
For Each objFolderB_CM In objFolder_CM.SubFolders

ReDim Preserve locFolders(1 To x) As String
locFolders(x) = objFolderB_CM.Name

x = x + 1

Next objFolderB_CM

getFolderInFolder2Array = locFolders

End Function

Public Function getDetailFolderTree(theFolder As String) As PathAndName()
'gets the pathsAndNames of all levels of folders to an array of pathandName   |

Dim thisLevel() As PathAndName
Dim nextLevel() As PathAndName

Dim totalFilesAndFolders() As PathAndName

thisLevel = DetailgetFilesAndFoldersArray(theFolder)

If Not Not thisLevel Then
Else
Exit Function
End If

nextLevel = getNextLevel(thisLevel)

Dim n As Integer: n = 1

Do While (n < 19)
nextLevel = getNextLevel(thisLevel)
totalFilesAndFolders = ConcatenatePathAndName(totalFilesAndFolders, thisLevel)
If PathAndNameArrayHasStuff(nextLevel) Then
thisLevel = nextLevel
Else
n = 20
End If
' Cells(n, 1).Value = UBound(totalFilesAndFolders)
n = n + 1
Loop

getDetailFolderTree = totalFilesAndFolders

End Function

Public Function getNextLevel(currentLevel() As PathAndName) As PathAndName()
' gets one level of pathsAndNames of all items in the folders in thecurrent level
Dim nextLvL() As PathAndName
Dim x As Integer
For x = LBound(currentLevel) To UBound(currentLevel)
If currentLevel(x).D_isFolder Then
nextLvL = ConcatenatePathAndName(nextLvL, DetailgetFilesAndFoldersArray(currentLevel(x).A_Path))
End If
Next x

getNextLevel = nextLvL

End Function

Public Function DetailgetFilesAndFoldersArray(theFolder As String) As PathAndName()
'gets all files and folders in the folder in thecurrent level
DetailgetFilesAndFoldersArray = ConcatenatePathAndName(DetailFilesInFolder2Array(theFolder), DetailFolderInFolder2Array(theFolder))
End Function

Public Function DetailFilesInFolder2Array(theFolder As String) As PathAndName()
' gets all files in the folder
Dim findToolsFSO As Object
Set findToolsFSO = CreateObject("Scripting.FileSystemObject") 'Create an instance of the FileSystemObject
Dim objFolder_CM As Object
Set objFolder_CM = findToolsFSO.GetFolder(theFolder) 'Get the folder object

'loops through each file in the directory and prints their names and path

Dim x As Integer

x = 1
Dim locFiles() As PathAndName
Dim objFile_CM As Object

' fileModDate = f.DateLastModified

For Each objFile_CM In objFolder_CM.Files

ReDim Preserve locFiles(1 To x) As PathAndName
locFiles(x).A_Path = objFile_CM.Path
locFiles(x).B_Name = objFile_CM.Name
locFiles(x).C_DateMod = objFile_CM.DateLastModified
x = x + 1

Next objFile_CM

DetailFilesInFolder2Array = locFiles

End Function

Public Function DetailFolderInFolder2Array(theFolder As String) As PathAndName()
' gets all folders in the folder
Dim findToolsFSO As Object
Set findToolsFSO = CreateObject("Scripting.FileSystemObject") 'Create an instance of the FileSystemObject

Dim objFolder_CM As Object
Set objFolder_CM = findToolsFSO.GetFolder(theFolder) 'Get the folder object

'loops through each file in the directory and prints their names and path

Dim x As Integer

x = 1
Dim locFolders() As PathAndName
Dim objFolderB_CM As Object
For Each objFolderB_CM In objFolder_CM.SubFolders

ReDim Preserve locFolders(1 To x) As PathAndName
locFolders(x).A_Path = objFolderB_CM.Path
locFolders(x).B_Name = objFolderB_CM.Name
locFolders(x).C_DateMod = objFolderB_CM.DateLastModified
locFolders(x).D_isFolder = True
x = x + 1

Next objFolderB_CM

DetailFolderInFolder2Array = locFolders

End Function

Public Function ConcatenatePathAndName(theArray1() As PathAndName, theArray2() As PathAndName) As PathAndName()
'  concatentaes two pathAndName arrays
Dim newArr() As PathAndName

Dim n As Integer: n = 1

Dim x As Integer
If PathAndNameArrayHasStuff(theArray1) Then
For x = LBound(theArray1) To UBound(theArray1)
ReDim Preserve newArr(1 To n) As PathAndName
newArr(n) = theArray1(x)
n = n + 1
Next x
End If
If PathAndNameArrayHasStuff(theArray2) Then

For x = LBound(theArray2) To UBound(theArray2)
ReDim Preserve newArr(1 To n) As PathAndName
newArr(n) = theArray2(x)
n = n + 1
Next x

End If

ConcatenatePathAndName = newArr

End Function

Public Function PathAndNameArrayHasStuff(theArr() As PathAndName) As Boolean
'true when there is data in array
'https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
If (Not Not theArr) <> 0 Then PathAndNameArrayHasStuff = True

End Function

Public Sub printPathAndNameToColumn(theArr() As PathAndName, theSheet As Worksheet, theCol As Integer)
'prints detail to sheet

Dim x As Integer: Dim n As Integer
n = 2
With theSheet
.Cells(1, theCol).Value = "Path"
.Cells(1, theCol + 1).Value = "Name"
.Cells(1, theCol + 2).Value = "Date"

If Not PathAndNameArrayHasStuff(theArr) Then Exit Sub

For x = LBound(theArr) To UBound(theArr)
.Cells(n, theCol).Value = theArr(x).A_Path
.Cells(n, theCol + 1).Value = theArr(x).B_Name
.Cells(n, theCol + 2).Value = theArr(x).C_DateMod
If theArr(x).D_isFolder Then Range(.Cells(n, theCol), .Cells(n, theCol + 2)).Interior.Color = RGB(225, 225, 225)
n = n + 1
Next x
End With

End Sub

Public Function namesFromPaths(thePaths() As String) As String()
'returns filenames given paths
Dim locPaths() As String
Dim x As Integer
Dim n As Integer: n = 1

For x = LBound(thePaths) To UBound(thePaths)
ReDim Preserve locPaths(1 To n) As String
locPaths(n) = nameFromPath(thePaths(x))
n = n + 1
Next x
namesFromPaths = locPaths

End Function

Public Function nameFromPath(thePath As String) As String
'returns filename given path
Dim SplitPath() As String
SplitPath = Split(thePath, "\")
nameFromPath = SplitPath(UBound(SplitPath))
End Function

Public Function pathFromName(thePath As String) As String
'returns path given full path
Dim SplitPath() As String
SplitPath = Split(thePath, "\")
    Dim tpath As String
    Dim x As Integer
    For x = LBound(SplitPath) To UBound(SplitPath) - 1
        tpath = tpath & SplitPath(x) & "\"
    Next x
pathFromName = tpath
End Function

Public Function getPathFromPandN(thePandN() As PathAndName) As String()
' retrieve the paths from the path and name type
Dim something() As String
Dim x As Integer
Dim n As Integer: n = 1
 
 For x = LBound(thePandN) To UBound(thePandN)
     ReDim Preserve something(1 To n) As String
     something(n) = thePandN(x).A_Path
     n = n + 1
 Next x
 
 getPathFromPandN = something

End Function

Public Sub ShowHideFilesInFolder(theFolder As String, theFiles() As String, theConfig As ShowHideConfig)
' shows or hides files in folder
Dim findToolsFSO As Object
Set findToolsFSO = CreateObject("Scripting.FileSystemObject") 'Create an instance of the FileSystemObject

Dim objFolder_CM As Object
Set objFolder_CM = findToolsFSO.GetFolder(theFolder) 'Get the folder object

'loops through each file in the directory and prints their names and path

Dim objFile_CM As Object
For Each objFile_CM In objFolder_CM.Files

Call showHideFileFolder(objFile_CM, theFiles, theConfig)

Next objFile_CM

For Each objFile_CM In objFolder_CM.SubFolders

Call showHideFileFolder(objFile_CM, theFiles, theConfig)

Next objFile_CM

End Sub

Private Sub showHideFileFolder(theItem As Object, theFiles() As String, theConfig As ShowHideConfig)
' helper
If theConfig = A_ShowAll Then theItem.Attributes = 0 ' show
If theConfig = B_HideAll Then theItem.Attributes = 2 ' hide
If theConfig = C_ShowAllExcept Then
If Not stringInArray(theItem.Name, theFiles) Then
theItem.Attributes = 0
Else
theItem.Attributes = 2
End If
End If

If theConfig = D_HideAllExcept Then
If Not stringInArray(theItem.Name, theFiles) Then
theItem.Attributes = 2
Else
theItem.Attributes = 0
End If
End If
If theConfig = E_HideThese Then
If stringInArray(theItem.Name, theFiles) Then theItem.Attributes = 2
End If

End Sub

