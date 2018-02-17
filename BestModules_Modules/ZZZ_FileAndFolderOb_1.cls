VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZZZ_FileAndFolderOb_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'$VERSIONCONTROL
'$*MINOR_VERSION*1.4
'$*DATE*15Feb18
'$*ID*FileAndFolder


' object for wrapper functions for file system object.
'Depends on ZZZ_StringArrayFunctions_6 and ZZZ_FileAndFolderRetrieval_5


'/---ZZZ_FileAndFolderOb_1--------updated 15Feb18----------------------------------------------------------------------------------\
'  Function Name                         | Return          |   Description                                                         |
'----------------------------------------|-----------------|-----------------------------------------------------------------------|
' initializeWithPaths                    | void            | pass an array of strings representing paths. this will set            |
'                                                          |  "pathsAndNames" to all the files and folders in that tree            |
' getAllPathsAndNames                    | PathAndName()   | getter for private variable PathsAndNames                             |
' getAllPathsAndNamesWithExtentions      | PathAndName()   | pass in an array of file extentions to get only the stuff with those  |
'                                                          | extentions                                                            |
'\---------------------------------------------------------------------------------------------------------------------------------/

Option Explicit

Private filePaths() As String
Private PathsAndNames() As PathAndName




Public Sub initializeWithPaths(Tpaths() As String)
    filePaths = Tpaths
    Dim x As Integer
    For x = LBound(filePaths) To UBound(filePaths)
        PathsAndNames = ConcatenatePathAndName(PathsAndNames, getDetailFolderTree(filePaths(x)))
    Next x

End Sub

Public Function getAllPathsAndNames() As PathAndName()
   getAllPathsAndNames = PathsAndNames
End Function


Public Function getAllPathsAndNamesWithExtention(anExtention As String) As PathAndName()
    Dim theExtention(1 To 1) As String: theExtention(1) = anExtention
    getAllPathsAndNamesWithExtention = getAllPathsAndNamesWithExtentions(theExtention)
End Function

Public Function getAllPathsAndNamesWithExtentions(theExtentions() As String) As PathAndName()

Dim locRetP() As PathAndName
Dim x As Integer: Dim n As Integer: n = 1
For x = LBound(PathsAndNames) To UBound(PathsAndNames)
    
    If hasValidExtention(theExtentions, PathsAndNames(x)) Then
        ReDim Preserve locRetP(1 To n) As PathAndName
        locRetP(n) = PathsAndNames(x)
        n = n + 1
    End If
    
Next x

getAllPathsAndNamesWithExtentions = locRetP

End Function


    
    Private Function hasValidExtention(validExtentions() As String, FileN As PathAndName) As Boolean
    
        Dim x As Integer
        For x = LBound(validExtentions) To UBound(validExtentions)
            If UCase(validExtentions(x)) = UCase(Right(FileN.B_Name, Len(validExtentions(x)))) Then hasValidExtention = True: Exit Function
        Next x
    
    End Function

