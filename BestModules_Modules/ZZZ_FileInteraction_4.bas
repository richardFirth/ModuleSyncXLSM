Attribute VB_Name = "ZZZ_FileInteraction_4"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.3
'$*DATE*8Feb18
'$*ID*FileInteraction
'$*CharCount*7748*xxxx
'$*RowCount*298*xxxx


Option Explicit

'/---ZZZ_FileInteraction_4-----------------updated 6Feb18-------------------------------------\
'  Function Name         | Return          |   Description                                    |
'------------------------|-----------------|--------------------------------------------------|
' CopyFileRF             | Boolean         | copies a file                                    |
' MoveFileRF             | Boolean         | moves a file                                     |
' DeleteFileRF           | Boolean         | deletes a file                                   |
' DeleteFolderRF         | Boolean         | deletes a folder (with all its files)            |
' DeleteFolderTreeRF     | Boolean         | deletes a folder tree                            |
' RenameFileRF           | Boolean         | renames a file                                   |
' createFolderOnDesktop  | Boolean         | creates a directory on desktop                   |
' createDirectoryRF      | Boolean         | creates a directory                              |
' FolderThere            | Boolean         | checks if a folder is present                    |
' FileThere              | Boolean         | checks if a file is present                      |
' setDefaultDirToOpen    | void            | set default folder for browser                   |
' BrowseToMacro          | Workbook        | browse to, open, and return a workbook object    |
' BrowseFilePath         | String          | gets a single file path                          |
' BrowseFilePaths        | String()        | gets multiple file paths                         |
' Private browse4type    | String          | internal, gets the type to browse for            |
' ConvertVariantToSTRArr | String()        | turns variant array into string array            |
'\--------------------------------------------------------------------------------------------/




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



' /------------------\
' |copies a file     |
' \------------------/
Public Function CopyFileRF(source As String, destination As String) As Boolean

On Error GoTo CopyFileRF_Fail
FileCopy source, destination
CopyFileRF = True
Exit Function

CopyFileRF_Fail:
CopyFileRF = False
End Function


' /------------------\
' |moves a file      |
' \------------------/
Public Function MoveFileRF(source As String, destination As String) As Boolean

On Error GoTo MoveFileRF_Fail

If CopyFileRF(source, destination) Then
Kill source
MoveFileRF = True
End If

Exit Function
MoveFileRF_Fail:
MoveFileRF = False
End Function

' /-------------------\
' |delete a file      |
' \-------------------/
Public Function DeleteFileRF(thePath As String) As Boolean
'You can use this to delete all the files in the folder Test
On Error GoTo DeleteFileRF_Fail
Kill thePath
DeleteFileRF = True
Exit Function
DeleteFileRF_Fail:
DeleteFileRF = False
End Function

' /---------------------\
' |delete a folder      |
' \---------------------/
Public Function DeleteFolderRF(thePath As String) As Boolean
'You can use this to delete the whole folder
'Note: RmDir delete only a empty folder
On Error GoTo DeleteFolderRF_Fail
Kill thePath & "\*.*"    ' delete all files in the folder
RmDir thePath  ' delete folder
DeleteFolderRF = True
Exit Function
DeleteFolderRF_Fail:
DeleteFolderRF = False
End Function



' /-------------------------\
' |delete a folder tree     |
' \-------------------------/
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





' /-----------------------\
' |renames a file         |
' \-----------------------/
Public Function RenameFileRF(source As String, destination As String) As Boolean
RenameFileRF = MoveFileRF(source, destination)
End Function



' /----------------------------------\
' |creates a directory on desktop    |
' \----------------------------------/
Public Function createFolderOnDesktop(ByVal dirName As String) As Boolean
createFolderOnDesktop = createDirectoryRF(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & dirName)
End Function



' /----------------------\
' |creates a directory   |
' \----------------------/
Public Function createDirectoryRF(ByVal dirName As String) As Boolean

On Error GoTo noFolder
MkDir dirName
On Error GoTo 0
createDirectoryRF = True

Exit Function

noFolder:
createDirectoryRF = False
End Function






' /--------------------------------\
' |checks if a folder is present   |
' \--------------------------------/
Public Function FolderThere(folderPathToTest As String) As Boolean

If folderPathToTest = "" Then FolderThere = False: Exit Function
If Len(Dir(folderPathToTest, vbDirectory)) = 0 Then
FolderThere = False
Else
FolderThere = True
End If


End Function



' /-------------------------------\
' |checks if a file is present    |
' \-------------------------------/
Public Function FileThere(theFileNameToTest As String) As Boolean

If theFileNameToTest = "" Then FileThere = False: Exit Function
If Len(Dir(theFileNameToTest)) = 0 Then
FileThere = False
Else
FileThere = True
End If

End Function



Public Function setDefaultDirToOpen(tDir As String)
ChDir tDir
'CreateObject("WScript.Shell").SpecialFolders("Desktop")
End Function


Public Function BrowseToMacro() As Workbook
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
Function ConvertVariantToSTRArr(theVariant() As Variant) As String()

Dim strARR() As String
Dim x As Integer
For x = LBound(theVariant) To UBound(theVariant)
ReDim Preserve strARR(1 To x) As String
strARR(x) = theVariant(x)
Next x
ConvertVariantToSTRArr = strARR

End Function






