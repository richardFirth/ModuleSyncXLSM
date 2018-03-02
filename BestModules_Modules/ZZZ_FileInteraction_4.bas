Attribute VB_Name = "ZZZ_FileInteraction_4"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.4
'$*DATE*2/28/2018*xx
'$*ID*FileInteraction
'$*CharCount*7070*xxxx
'$*RowCount*235*xxxx

'/T--ZZZ_FileInteraction_4--------------------------------------------------------\
' Function Name          | Return    |  Description                               |
'------------------------|-----------|--------------------------------------------|
'----- Directory interactions-----------------------------------------------------|
'CopyFileRF              | Boolean   | copies a file                              |
'MoveFileRF              | Boolean   | moves a file                               |
'DeleteFileRF            | Boolean   | deletes a file                             |
'DeleteFolderRF          | Boolean   | deletes a folder and its contents          |
'DeleteFolderTreeRF      | Boolean   | Delete all files and subfolders            |
'RenameFileRF            | Boolean   | renames a file                             |
'createFolderOnDesktop   | Boolean   | creates a directory on desktop             |
'createDirectoryRF       | Boolean   | creates a directory                        |
'FolderThere             | Boolean   | checks if a folder is present              |
'FileThere               | Boolean   | checks if a file is present                |
'----- Browsing to files----------------------------------------------------------|
'setDefaultDirToOpen     | String)   |  changes where browse starts               |
'BrowseToMacro           | Workbook  |  browse to an excel macro                  |
'BrowseFilePath          | String    | Gets path of text file for importing data  |
'BrowseFilePaths         | String()  | Browse to many paths                       |
'browse4type             | String    | Browse to one paths                        |
'ConvertVariantToSTRArr  | String()  |  Converts variants to string arr           |
'\--------------------------------------------------------------------------------/

Option Explicit

Public Enum getFileType
A_CSV
B_EXCEL
C_EXCEL_OLD
D_EXCEL_MACRO
E_InoFile
F_Proc3File
G_VBAModule
H_Text
I_Lib
J_BRD
K_SCH
End Enum

'# Directory interactions

Public Function CopyFileRF(source As String, destination As String) As Boolean
'copies a file
On Error GoTo CopyFileRF_Fail
FileCopy source, destination
CopyFileRF = True
Exit Function

CopyFileRF_Fail:
CopyFileRF = False
End Function

Public Function MoveFileRF(source As String, destination As String) As Boolean
'moves a file
On Error GoTo MoveFileRF_Fail

If CopyFileRF(source, destination) Then
Kill source
MoveFileRF = True
End If

Exit Function
MoveFileRF_Fail:
MoveFileRF = False
End Function

Public Function DeleteFileRF(thePath As String) As Boolean
'deletes a file
On Error GoTo DeleteFileRF_Fail
Kill thePath
DeleteFileRF = True
Exit Function
DeleteFileRF_Fail:
DeleteFileRF = False
End Function

Public Function DeleteFolderRF(thePath As String) As Boolean
'deletes a folder and its contents
On Error GoTo DeleteFolderRF_Fail
Kill thePath & "\*.*"    ' delete all files in the folder
RmDir thePath  ' delete folder
DeleteFolderRF = True
Exit Function
DeleteFolderRF_Fail:
DeleteFolderRF = False
End Function

Public Function DeleteFolderTreeRF(thePath As String) As Boolean
'Delete all files and subfolders
'Be sure that no file is open in the folder
Dim FSO As Object

Set FSO = CreateObject("scripting.filesystemobject")

If Right(thePath, 1) = "\" Then
thePath = Left(thePath, Len(thePath) - 1)
End If

If FSO.FolderExists(thePath) = False Then DeleteFolderTreeRF = False:  Exit Function

On Error Resume Next
FSO.deletefile thePath & "\*.*", True 'Delete files
FSO.deletefolder thePath & "\*.*", True 'Delete subfolders
FSO.deletefolder thePath
Call DeleteFolderRF(thePath)
On Error GoTo 0
DeleteFolderTreeRF = True
End Function

Public Function RenameFileRF(source As String, destination As String) As Boolean
'renames a file
RenameFileRF = MoveFileRF(source, destination)
End Function

Public Function createFolderOnDesktop(ByVal dirName As String) As Boolean
'creates a directory on desktop
createFolderOnDesktop = createDirectoryRF(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & dirName)
End Function

Public Function createDirectoryRF(ByVal dirName As String) As Boolean
'creates a directory
On Error GoTo noFolder
MkDir dirName
On Error GoTo 0
createDirectoryRF = True

Exit Function

noFolder:
createDirectoryRF = False
End Function

Public Function FolderThere(folderPathToTest As String) As Boolean
'checks if a folder is present
If folderPathToTest = "" Then FolderThere = False: Exit Function
If Len(Dir(folderPathToTest, vbDirectory)) = 0 Then
FolderThere = False
Else
FolderThere = True
End If

End Function

Public Function FileThere(theFileNameToTest As String) As Boolean
'checks if a file is present
If theFileNameToTest = "" Then FileThere = False: Exit Function
If Len(Dir(theFileNameToTest)) = 0 Then
FileThere = False
Else
FileThere = True
End If

End Function

'# Browsing to files

Public Function setDefaultDirToOpen(tDir As String)
' changes where browse starts
ChDir tDir
'CreateObject("WScript.Shell").SpecialFolders("Desktop")
End Function

Public Function BrowseToMacro() As Workbook
' browse to an excel macro
Set BrowseToMacro = Workbooks.Open(BrowseFilePath(D_EXCEL_MACRO))
End Function

Public Function BrowseFilePath(theType As getFileType) As String
'Gets path of text file for importing data
Dim sFullName As String
Dim sFileName As String
sFullName = Application.GetOpenFilename(browse4type(theType))
If sFullName = "False" Then End
BrowseFilePath = sFullName

End Function

Public Function BrowseFilePaths(theType As getFileType) As String()
'Browse to many paths
On Error GoTo BrowseFilePathsError
Dim sFullName() As Variant
sFullName() = Application.GetOpenFilename(browse4type(theType), , , , True)
BrowseFilePaths = ConvertVariantToSTRArr(sFullName)
Exit Function

BrowseFilePathsError:
End
Dim errorStr(1 To 1) As String
'errorStr = -1
errorStr(1) = "No Selection"
BrowseFilePaths = errorStr

End Function

Private Function browse4type(theType As getFileType) As String
'Browse to one paths
If theType = A_CSV Then browse4type = "*.csv,*.csv"
If theType = B_EXCEL Then browse4type = "*.xlsx,*.xlsx"
If theType = C_EXCEL_OLD Then browse4type = "*.xls,*.xls"
If theType = D_EXCEL_MACRO Then browse4type = "*.xlsm,*.xlsm"
If theType = E_InoFile Then browse4type = "*.ino,*.ino"
If theType = F_Proc3File Then browse4type = "*.pde,*.pde"
If theType = G_VBAModule Then browse4type = "*.bas,*.bas"
If theType = H_Text Then browse4type = "*.txt,*.txt"
If theType = I_Lib Then browse4type = "*.lbr,*.lbr"
If theType = J_BRD Then browse4type = "*.brd,*.brd"
If theType = K_SCH Then browse4type = "*.sch,*.sch"

' "Visual Basic Files (.bas; *.txt),.bas;*.txt"
End Function

' /--------------------------------------------\
' |converts a variant array to a string array  |
' \--------------------------------------------/
Private Function ConvertVariantToSTRArr(theVariant() As Variant) As String()
' Converts variants to string arr
Dim strARR() As String
Dim x As Integer
For x = LBound(theVariant) To UBound(theVariant)
ReDim Preserve strARR(1 To x) As String
strARR(x) = theVariant(x)
Next x
ConvertVariantToSTRArr = strARR

End Function

