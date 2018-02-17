Attribute VB_Name = "ZZZ_FileAndFolderRetrieval_6"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.1
'$*DATE*5Feb18
'$*ID*FileAndFolderRetrieval

Option Explicit

'/---ZZZ_FileAndFolderRetrieval_6--------updated 5Feb18-------------------------------------------------\
'  Function Name                   | Return          |   Description                                    |
'----------------------------------|-----------------|--------------------------------------------------|
' getFilesAndFoldersArray          | String()        | gets files and folders as string array           |
' getFilesInFolder2Array           | String()        | gets files as string array                       |
' getFilePathsInFolder2Array       | String()        | gets filepaths as string array                   |
' getFolderInFolder2Array          | String()        | gets folders as string array                     |
' getDetailFolderTree              | PathAndName()   | gets everything in folder structure              |
' getNextLevel                     | PathAndName()   | gets next level paths and names for each folder  |
' DetailgetFilesAndFoldersArray    | PathAndName()   | gets pathandname for all files and folders       |
' DetailFilesInFolder2Array        | PathAndName()   | gets pathandname for all files                   |
' DetailFolderInFolder2Array       | PathAndName()   | gets pathandname for all folders                 |
' ConcatenatePathAndName           | PathAndName()   | concatentaes two pathAndName arrays              |
' PathAndNameArrayHasStuff         | boolean         | true when there is data in array                 |
' printPathAndNameToColumn         | void            | prints detail to sheet                           |
' namesFromPaths                   | String()        | returns filenames given paths                    |
' nameFromPath                     | String          | returns filename given path                      |
' ShowHideFilesInFolder            | void            | shows or hides files in a folder                 |
' private showHideFileFolder       | void            | used by ShowHideFilesInFolder for looping        |
'\------------------------------------------------------------------------------------------------------/

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@@ Depends on ZZZ_StringArrayFunctions_5 @@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


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


 ' /------------------------------------------------------\
 ' |gets files and folders as string array                |
 ' \------------------------------------------------------/
Public Function getFilesAndFoldersArray(theFolder As String) As String()
    getFilesAndFoldersArray = ConcatenateArrays(getFilesInFolder2Array(theFolder), getFolderInFolder2Array(theFolder))
End Function



 ' /------------------------------------------\
 ' |gets files as string array                |
 ' \------------------------------------------/
Public Function getFilesInFolder2Array(theFolder As String) As String()
    getFilesInFolder2Array = namesFromPaths(getFilePathsInFolder2Array(theFolder))
End Function

 ' /----------------------------------------------\
 ' |gets file path as string array                |
 ' \----------------------------------------------/
Public Function getFilePathsInFolder2Array(theFolder As String) As String()

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



 ' /--------------------------------------------\
 ' |gets folders as string array                |
 ' \--------------------------------------------/
Public Function getFolderInFolder2Array(theFolder As String) As String()

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



' /++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\
' | gets the pathsAndNames of all levels of folders to an array of pathandName   |
' |                                                                              |
' \++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++/
Public Function getDetailFolderTree(theFolder As String) As PathAndName()


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

' /------------------------------------------------------------------------------------\
' | gets one level of pathsAndNames of all items in the folders in thecurrent level    |
' \------------------------------------------------------------------------------------/
Public Function getNextLevel(currentLevel() As PathAndName) As PathAndName()

Dim nextLvL() As PathAndName
Dim x As Integer
    For x = LBound(currentLevel) To UBound(currentLevel)
        If currentLevel(x).D_isFolder Then
            nextLvL = ConcatenatePathAndName(nextLvL, DetailgetFilesAndFoldersArray(currentLevel(x).A_Path))
        End If
    Next x
    
getNextLevel = nextLvL

End Function



' /-----------------------------------------------------------------\
' | gets all files and folders in the folder in thecurrent level    |
' \-----------------------------------------------------------------/
Public Function DetailgetFilesAndFoldersArray(theFolder As String) As PathAndName()
    DetailgetFilesAndFoldersArray = ConcatenatePathAndName(DetailFilesInFolder2Array(theFolder), DetailFolderInFolder2Array(theFolder))
End Function



' /----------------------------------\
' | gets all files in the folder     |
' \----------------------------------/
Public Function DetailFilesInFolder2Array(theFolder As String) As PathAndName()

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



' /------------------------------------\
' | gets all folders in the folder     |
' \------------------------------------/
Public Function DetailFolderInFolder2Array(theFolder As String) As PathAndName()

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



' /-------------------------------------\
' | concatentaes two pathAndName arrays |
' \-------------------------------------/
Public Function ConcatenatePathAndName(theArray1() As PathAndName, theArray2() As PathAndName) As PathAndName()

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




' /------------------------------------------------\
' |true when there is data in array                |
' \------------------------------------------------/
Public Function PathAndNameArrayHasStuff(theArr() As PathAndName) As Boolean

'https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
    If (Not Not theArr) <> 0 Then PathAndNameArrayHasStuff = True

End Function


' /--------------------------------------\
' |prints detail to sheet                |
' \--------------------------------------/
Public Sub printPathAndNameToColumn(theArr() As PathAndName, theSheet As Worksheet, theCol As Integer)
    
    
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


' /--------------------------------------\
' |returns filenames given paths         |
' \--------------------------------------/
Public Function namesFromPaths(thePaths() As String) As String()

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



' /--------------------------------------\
' |returns filename given path           |
' \--------------------------------------/
Public Function nameFromPath(thePath As String) As String

Dim SplitPath() As String
SplitPath = Split(thePath, "\")

nameFromPath = SplitPath(UBound(SplitPath))

End Function




' /----------------------------------------------\
' |shows and hides files and folders for a path  |
' \----------------------------------------------/
Public Sub ShowHideFilesInFolder(theFolder As String, theFiles() As String, theConfig As ShowHideConfig)

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







