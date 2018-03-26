Attribute VB_Name = "BBB_HeaderDataFactory"

'/T--BBB_HeaderDataFactory-------------------------------------------------------------------------------------------------------\
' Function Name                    | Return                       |  Description                                                 |
'----------------------------------|------------------------------|--------------------------------------------------------------|
'createModuleObjectsCollection     | ModuleVersionDataObjectsCol  | factory function for theModuleVersionDataObjects collection  |
'createModuleHeaderObjectFromWKBK  | ModuleVersionDataObject      |  factory function for ModuleVersionDataObject                |
'ConcatenateModuleVersionData      | ModuleVersionData()          |  concatenates two moduleversiondata type arrays              |
'ModuleVersionDataHasStuff         | Boolean                      |  checks if array is initialized                              |
'\-------------------------------------------------------------------------------------------------------------------------------/

Option Explicit

Function createModuleObjectsCollection(tPaths() As String) As ModuleVersionDataObjectsCol
'factory function for theModuleVersionDataObjects collection
Dim locHDat As New ModuleVersionDataObjectsCol
Dim locCol As New Collection
Dim aWKBK As Workbook
    Call createDirectoryRF(folderToPlaceData & "\ModuleSyncOutput") ' to allow output to all me in the module sync
    Dim theFolderPath As String: theFolderPath = folderToPlaceData & "\ModuleSyncOutput\Modules_"
    Dim x As Integer
    For x = LBound(tPaths) To UBound(tPaths)
        Dim otherWKBKHeaders As ModuleVersionDataObject
        If FileThere(tPaths(x)) Then
            Application.EnableEvents = False ' stop macros triggering when the workbook opens
            Set aWKBK = Workbooks.Open(tPaths(x))
            Dim theFileName As String: theFileName = Left(nameFromPath(tPaths(x)), Len(nameFromPath(tPaths(x))) - 5)
            Set otherWKBKHeaders = createModuleHeaderObjectFromWKBK(aWKBK, theFolderPath & theFileName) ' factory function for ModuleVersionDataObject
            Call locCol.Add(otherWKBKHeaders)
            Call aWKBK.Close(False)
            Application.EnableEvents = True ' stop macros triggering when the workbook opens
        End If
    Next x
Call locHDat.setModuleVersionDataObjects(locCol)
Set createModuleObjectsCollection = locHDat ' return newly created object
End Function

Function createModuleHeaderObjectFromWKBK(theWKBK As Workbook, theFolderName As String) As ModuleVersionDataObject
' factory function for ModuleVersionDataObject
    Dim myHData As New ModuleVersionDataObject
    Dim thisCollection As Collection
    Set thisCollection = getAllModules(theWKBK)
    Dim modulePaths() As String
    If Not FolderThere(theFolderName) Then ' make sure folder creation is possible
        If Not createDirectoryRF(theFolderName) Then GoTo createModuleHeaderObjectFromWKBKErr
    End If
    modulePaths() = ExportVBAModulesToPaths(thisCollection, theFolderName)
    Dim singleModcollection As New Collection
    Dim x As Integer
    Dim singleModObj As X_SingleModuleObject_1
    For x = LBound(modulePaths) To UBound(modulePaths)
        Set singleModObj = New X_SingleModuleObject_1
        Call singleModObj.initializeModule(modulePaths(x))
        Call singleModcollection.Add(singleModObj)
    Next x
    Call myHData.setSingleModules(singleModcollection)
    Call myHData.setWKBKPath(theWKBK.Path & "\" & theWKBK.Name)
    Call myHData.setModulesFolderPath(theFolderName)
    Call myHData.setModulePaths(modulePaths)
    Set createModuleHeaderObjectFromWKBK = myHData
Exit Function
createModuleHeaderObjectFromWKBKErr:
    Dim tERR(1 To 2) As String
    tERR(1) = theFolderName
    tERR(2) = "Folder create error"
   Call reportError("createModuleHeaderObjectFromWKBK", tERR)
End Function

Public Function ConcatenateModuleVersionData(theArray1() As ModuleVersionData, theArray2() As ModuleVersionData) As ModuleVersionData()
' concatenates two moduleversiondata type arrays
Dim newArr() As ModuleVersionData
Dim n As Long: n = 1
Dim x As Long
If ModuleVersionDataHasStuff(theArray1) Then
    For x = LBound(theArray1) To UBound(theArray1)
        ReDim Preserve newArr(1 To n) As ModuleVersionData
        newArr(n) = theArray1(x)
        n = n + 1
    Next x
End If
If ModuleVersionDataHasStuff(theArray2) Then
    For x = LBound(theArray2) To UBound(theArray2)
        ReDim Preserve newArr(1 To n) As ModuleVersionData
        newArr(n) = theArray2(x)
        n = n + 1
    Next x
End If
ConcatenateModuleVersionData = newArr
End Function

Public Function ModuleVersionDataHasStuff(theArr() As ModuleVersionData) As Boolean
' checks if array is initialized
'https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
    If (Not Not theArr) <> 0 Then ModuleVersionDataHasStuff = True
End Function

