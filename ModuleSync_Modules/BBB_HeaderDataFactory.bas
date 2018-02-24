Attribute VB_Name = "BBB_HeaderDataFactory"
Option Explicit





'/---BBB_ModuleVersionDataFactory-------------------------updated 21Feb18--------------------------------------------------------------------------------------------\
'  Function Name                     | Return                  |   Description                                                                                |
'------------------------------------|-------------------------|----------------------------------------------------------------------------------------------|
' createHeaderObjectsCollection      | ModuleVersionDataObjectsCol    | factory function for theModuleVersionDataObjects collection                                         |
' createModuleHeaderObjectFromWKBK   | ModuleVersionDataObject        | factory function for ModuleVersionDataObject                                                        |




' this creates the ModuleVersionDataObjectCol, and also puts the text files of all the modules onto the desktop

Function createHeaderObjectsCollection(tPaths() As String) As ModuleVersionDataObjectsCol

Dim locHDat As New ModuleVersionDataObjectsCol
Dim locCol As New Collection

Dim aWKBK As Workbook
    Call createDirectoryRF(folderToPlaceData & "\ModuleSyncOutput") ' to allow output to all me in the module sync
    Dim theFolderPath As String:
    
    theFolderPath = folderToPlaceData & "\ModuleSyncOutput\Modules_"

    Dim x As Integer
    For x = LBound(tPaths) To UBound(tPaths)
        Dim otherWKBKHeaders As ModuleVersionDataObject
        If FileThere(tPaths(x)) Then
            Set aWKBK = Workbooks.Open(tPaths(x))
            Dim theFileName As String: theFileName = Left(nameFromPath(tPaths(x)), Len(nameFromPath(tPaths(x))) - 5)
            Set otherWKBKHeaders = createModuleHeaderObjectFromWKBK(aWKBK, theFolderPath & theFileName) ' factory function for ModuleVersionDataObject
            Call locCol.Add(otherWKBKHeaders)
            Call aWKBK.Close(False)
        Else
            Dim tErr(1 To 2) As String
            tErr(1) = tPaths(x)
            tErr(2) = Left(nameFromPath(tPaths(x)), Len(nameFromPath(tPaths(x))) - 5)
            Call reportError("createHeaderObjectsCollection", tErr)
        End If
    Next x

Call locHDat.setModuleVersionDataObjects(locCol)
Set createHeaderObjectsCollection = locHDat ' return newly created object
End Function


Function createModuleHeaderObjectFromWKBK(theWKBK As Workbook, theFolderName As String) As ModuleVersionDataObject

Dim myHData As New ModuleVersionDataObject

    Dim thisCollection As Collection
    Set thisCollection = getAllModules(theWKBK)
    
    Dim modulePaths() As String
    
    If Not FolderThere(theFolderName) Then ' make sure folder creation is possible
        If Not createDirectoryRF(theFolderName) Then GoTo createModuleHeaderObjectFromWKBKErr
    End If
    
    modulePaths() = ExportVBAModulesToPaths(thisCollection, theFolderName)
    
    Call myHData.setModVData(extractModVerData(modulePaths))
    Call myHData.setWKBKPath(theWKBK.Path & "\" & theWKBK.Name)
    Call myHData.setModulesFolderPath(theFolderName)
    Call myHData.setModulePaths(modulePaths)
    Set createModuleHeaderObjectFromWKBK = myHData
    
Exit Function
createModuleHeaderObjectFromWKBKErr:
    Dim tErr(1 To 2) As String
    tErr(1) = theFolderName
    tErr(2) = "Folder create error"
   Call reportError("createModuleHeaderObjectFromWKBK", tErr)
    
End Function



' ** NOT USED CURRENTLY **
Function getModuleHeaderObjectFromFolder(theFolderName As String) As ModuleVersionDataObject

MsgBox "is this used?"

    Dim myD As New ModuleVersionDataObject
   
    Dim modulePaths() As String
    modulePaths() = getFilePathsInFolder2Array(theFolderName)
    Call myD.setModVData(extractModVerData(modulePaths))
    Call myD.setWKBKPath("N/A")
    Call myD.setModulePaths(modulePaths)
    Call myD.setModulesFolderPath(theFolderName)
    Set getModuleHeaderObjectFromFolder = myD
    
End Function




' gets an array of header data given the paths to the text files of that header data.
' this represents the header data all modules in a single workbook
Function extractModVerData(modulePaths() As String) As ModuleVersionData()
    Dim locModuleVersionData() As ModuleVersionData
    Dim x As Integer
    Dim n As Integer: n = 1
    
    If arrayHasStuff(modulePaths) Then
    
        For x = LBound(modulePaths) To UBound(modulePaths)
            ReDim Preserve locModuleVersionData(1 To n) As ModuleVersionData
            locModuleVersionData(n) = extractSingleModVerData(modulePaths(x))
            n = n + 1
        Next x
        extractModVerData = locModuleVersionData
    
    End If
End Function



Function extractSingleModVerData(modulePath As String) As ModuleVersionData

Dim locModuleVersionData As ModuleVersionData
locModuleVersionData.A_Name = "N/A"

If Not FileThere(modulePath) Then GoTo extractSingleModVerDataError

Dim VBACODE() As String: VBACODE = getTxTDocumentAsString(modulePath)

locModuleVersionData.F_ModulePath = modulePath
locModuleVersionData.A_Name = extractVersionName(modulePath)

locModuleVersionData.E_Vcontrol = underVersionControl(VBACODE)

    If locModuleVersionData.E_Vcontrol Then
        locModuleVersionData.J_CodeChange = detectCodeChange(VBACODE, modulePath)
        locModuleVersionData.B_MajorVersion = extractMajorVersion(locModuleVersionData.A_Name)
        locModuleVersionData.C_MinorVersion = extractMinorVer(VBACODE)
        locModuleVersionData.H_ID = extractModuleID(VBACODE)
        locModuleVersionData.D_date = getDateVer(VBACODE)
        
    End If

extractSingleModVerData = locModuleVersionData

Exit Function
extractSingleModVerDataError:
     Dim tErr(1 To 3) As String
     tErr(1) = modulePath
     
    Call reportError("extractSingleModVerData", tErr)
 
End Function


'$VERSIONCONTROL
'$*MINOR_VERSION*1.0
'$*DATE*18Jan18
'$*NAME*StringLookupTables

Function underVersionControl(theCode() As String) As Boolean
  Dim x As Integer
  For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 16) = "'$VERSIONCONTROL" Then underVersionControl = True: Exit Function
        If x = 30 Then Exit Function ' versioncontrol won't be this far down
  Next x
End Function

Function extractVersionName(fullFilePath As String) As String
    Dim zz() As String
    zz = Split(fullFilePath, "\")
    extractVersionName = Left(zz(UBound(zz)), Len(zz(UBound(zz))) - 4)
End Function

Function extractMajorVersion(aName As String) As String
    Dim zz() As String
    zz = Split(aName, "_")
    extractMajorVersion = zz(UBound(zz))

End Function

Function extractMinorVer(theCode() As String) As String
  Dim x As Integer
  For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 8) = "'$*MINOR" Then
            Dim loc() As String
            loc = Split(theCode(x), "*")
            If loc(1) = "MINOR_VERSION" Then extractMinorVer = loc(2): Exit Function
        End If
        If x = 30 Then Exit For ' extractMinorVer won't be this far down
  Next x
   
  extractMinorVer = "NA"
    
End Function



Function getDateVer(theCode() As String) As String
  Dim x As Integer
  For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 7) = "'$*DATE" Then
            Dim loc() As String
            loc = Split(theCode(x), "*")
            If loc(1) = "ID" Then getDateVer = loc(2): Exit Function
        End If
        If x = 30 Then Exit For ' extractMinorVer won't be this far down
  Next x
  
  getDateVer = "NA"
    
End Function

Function extractModuleID(theCode() As String) As String
  Dim x As Integer
  For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 5) = "'$*ID" Then
            Dim loc() As String
            loc = Split(theCode(x), "*")
            If loc(1) = "ID" Then extractModuleID = loc(2): Exit Function
        End If
        If x = 30 Then Exit For ' extractMinorVer won't be this far down
  Next x
  
  extractModuleID = "NA"
    
End Function



Public Function ConcatenateModuleVersionData(theArray1() As ModuleVersionData, theArray2() As ModuleVersionData) As ModuleVersionData()

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
'https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
    If (Not Not theArr) <> 0 Then ModuleVersionDataHasStuff = True
End Function


