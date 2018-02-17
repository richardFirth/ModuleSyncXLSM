Attribute VB_Name = "GGA_HeaderDataFactory"
Option Explicit



Public Type HeaderData
    A_Name As String
    B_MajorVersion As String
    C_MinorVersion As String
    D_date As String
    
    E_Vcontrol As Boolean
    F_ModulePath As String
    G_OldVersion As Boolean
    H_ID As String
    
    I_newModule As Boolean
End Type



Function createHeaderObjectsCollection(tPaths() As String) As HeaderDataObjectsCol

Dim locHDat As New HeaderDataObjectsCol
Dim locCol As New Collection

Dim aWKBK As Workbook

    Dim x As Integer
    For x = LBound(tPaths) To UBound(tPaths)
        Dim otherWKBKHeaders As HeaderDataObject
        Call createFolderOnDesktop("ModuleSyncOutput") ' to allow output to all me in the module sync
        If FileThere(tPaths(x)) Then
            Set aWKBK = Workbooks.Open(tPaths(x))
            Dim theFileName As String: theFileName = Left(nameFromPath(tPaths(x)), Len(nameFromPath(tPaths(x))) - 5)

            Dim theFolderPath As String: theFolderPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\ModuleSyncOutput\Modules_"
            
            Set otherWKBKHeaders = getModuleHeaderObjectFromWKBK(aWKBK, theFolderPath & theFileName)
               
            Call locCol.Add(otherWKBKHeaders)
            Call aWKBK.Close(False)
        End If
    Next x

Call locHDat.setHeaderDataObjects(locCol)
Call locHDat.setTotalPaths(tPaths)
Set createHeaderObjectsCollection = locHDat

End Function






Function getModuleHeaderObjectFromWKBK(theWKBK As Workbook, theFolderName As String) As HeaderDataObject

Dim myD As New HeaderDataObject

    Dim thisCollection As Collection
    Set thisCollection = combineCollections(getModulesByType(theWKBK, A_NORMAL), getModulesByType(theWKBK, B_CLASS))
    Set thisCollection = combineCollections(thisCollection, getModulesByType(theWKBK, C_FORM))
      
    Dim modulePaths() As String
    
    If Not FolderThere(theFolderName) Then
        If Not createDirectoryRF(theFolderName) Then GoTo getModuleHeaderObjectFromWKBKErr
    End If
    
    modulePaths() = ExportVBAModulesToPaths(thisCollection, theFolderName)
    
    Call myD.setHeader(getHeaders(modulePaths))
    Call myD.setWKBKPath(theWKBK.Path & "\" & theWKBK.Name)
    Call myD.setModulesFolderPath(theFolderName)
     
    Set getModuleHeaderObjectFromWKBK = myD
    
    
Exit Function
getModuleHeaderObjectFromWKBKErr:
    Dim tErr(1 To 2) As String
    tErr(1) = theFolderName
    tErr(2) = "Folder create error"
   Call reportError("getModuleHeaderObjectFromWKBK", tErr)
    
End Function


Function getModuleHeaderObjectFromFolder(theFolderName As String) As HeaderDataObject

    Dim myD As New HeaderDataObject
   
    Dim modulePaths() As String
    modulePaths() = getFilePathsInFolder2Array(theFolderName)
    Call myD.setHeader(getHeaders(modulePaths))
    Call myD.setWKBKPath("N/A")
    Call myD.setModulesFolderPath(theFolderName)
    Set getModuleHeaderObjectFromFolder = myD
    
End Function







Function getHeaders(modulePaths() As String) As HeaderData()
    Dim locHeaderData() As HeaderData
    
    Dim x As Integer
    Dim n As Integer: n = 1
    
    If arrayHasStuff(modulePaths) Then
    
        For x = LBound(modulePaths) To UBound(modulePaths)
            ReDim Preserve locHeaderData(1 To n) As HeaderData
            locHeaderData(n) = getHeader(modulePaths(x))
            n = n + 1
        Next x
        getHeaders = locHeaderData
    
    End If
End Function





Function getHeader(modulePath As String) As HeaderData


Dim locHeaderData As HeaderData

locHeaderData.A_Name = "N/A"

If Not FileThere(modulePath) Then GoTo getHeaderError

Dim VBACode() As String: VBACode = getTxTDocumentAsString(modulePath)

locHeaderData.F_ModulePath = modulePath
locHeaderData.A_Name = getName(modulePath)


If UBound(VBACode) > 20 Then

    If UCase(Right(modulePath, 3)) = "BAS" Then
            locHeaderData.E_Vcontrol = underVersionControl(VBACode(2))
            If locHeaderData.E_Vcontrol Then
                locHeaderData.B_MajorVersion = getMajorVersion(locHeaderData.A_Name)
                locHeaderData.C_MinorVersion = getMinorVer(VBACode(3))
                locHeaderData.D_date = getDateVer(VBACode(4))
                locHeaderData.H_ID = getModuleID(VBACode(5))
            End If
    End If
    If UCase(Right(modulePath, 3)) = "CLS" Then
            locHeaderData.E_Vcontrol = underVersionControl(VBACode(2 + 8))
           ' MsgBox VBACode(10) & "|" & VBACode(11) & "|" & VBACode(12) & "|" & VBACode(13) & "|" & VBACode(14)
            If locHeaderData.E_Vcontrol Then
                locHeaderData.B_MajorVersion = getMajorVersion(locHeaderData.A_Name)
                locHeaderData.C_MinorVersion = getMinorVer(VBACode(3 + 8))
                locHeaderData.D_date = getDateVer(VBACode(4 + 8))
                locHeaderData.H_ID = getModuleID(VBACode(5 + 8))
            End If
    End If
    
    If UCase(Right(modulePath, 3)) = "FRM" Then
    
            Dim addVal As Integer: addVal = 14
    
            locHeaderData.E_Vcontrol = underVersionControl(VBACode(2 + addVal))
            
            If Not locHeaderData.E_Vcontrol Then
                addVal = 15
                locHeaderData.E_Vcontrol = underVersionControl(VBACode(2 + addVal))
            End If
            
            If locHeaderData.E_Vcontrol Then
                locHeaderData.B_MajorVersion = getMajorVersion(locHeaderData.A_Name)
                locHeaderData.C_MinorVersion = getMinorVer(VBACode(3 + addVal))
                locHeaderData.D_date = getDateVer(VBACode(4 + addVal))
                locHeaderData.H_ID = getModuleID(VBACode(5 + addVal))
            End If
  
    End If

End If

    getHeader = locHeaderData
Exit Function
getHeaderError:
 Dim tErr(1 To 3) As String
 tErr(1) = modulePath
 
Call reportError("getHeader", tErr)
 
End Function


'$VERSIONCONTROL
'$*MINOR_VERSION*1.0
'$*DATE*18Jan18
'$*NAME*StringLookupTables

Function getName(fullFilePath As String) As String
    Dim zz() As String
    zz = Split(fullFilePath, "\")
    getName = Left(zz(UBound(zz)), Len(zz(UBound(zz))) - 4)
End Function

Function getMajorVersion(aName As String) As String
    Dim zz() As String
    zz = Split(aName, "_")
    getMajorVersion = zz(UBound(zz))

End Function

Function underVersionControl(firstSTR As String) As Boolean
    If Left(firstSTR, 16) = "'$VERSIONCONTROL" Then underVersionControl = True
End Function


Function getMinorVer(theRaw As String) As String
    
    If Mid(theRaw, 2, 1) <> "$" Then getMinorVer = "NA": Exit Function
    Dim LOC() As String
    LOC = Split(theRaw, "*")
    If LOC(1) = "MINOR_VERSION" Then getMinorVer = LOC(2): Exit Function
    getMinorVer = "NA"
    
End Function

Function getDateVer(theRaw As String) As String
    If Mid(theRaw, 2, 1) <> "$" Then getDateVer = "NA": Exit Function
    Dim LOC() As String
    LOC = Split(theRaw, "*")
     If LOC(1) = "DATE" Then getDateVer = LOC(2): Exit Function
    getDateVer = "NA"
    
End Function

Function getModuleID(theRaw As String) As String
    If Mid(theRaw, 2, 1) <> "$" Then getModuleID = "NA": Exit Function
    Dim LOC() As String
    LOC = Split(theRaw, "*")
     If LOC(1) = "ID" Then getModuleID = LOC(2): Exit Function
    getModuleID = "NA"
    
End Function






