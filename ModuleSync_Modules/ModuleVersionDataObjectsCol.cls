VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModuleVersionDataObjectsCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Purpose:
' The ModuleVersionDataObjectsCol contains a collection of ModuleVersionDataObject that represent a number of workbooks
' that can by synced



'/------ModuleVersionDataObjectsCol-------------------------------------------------------------------------------------------------------------------------\
' Factory Function: createHeaderObjectsCollection(tPaths() As String) As ModuleVersionDataObjectsCol                                                        |
'-----------------------------------------------------------------------------------------------------------------------------------------------------------|
'  Function Name                     | Return          |   Description                                                                                      |
'------------------------------------|-----------------|----------------------------------------------------------------------------------------------------|
' setModuleVersionDataObjects               | void            | sets the theModuleVersionDataObjects collection                                             |
' getTotalPaths                      | String()        | gets the workbook paths from each HeaderDataObject                                                 |
' getTotalNames                      | String()        | gets workbook names                                                                                |
' getPathsWithData                   | String()        | for use with UI                                                                                    |
' identifyAllOldModules          | void            | gets newestData() and calls identifyOldModuleVersions in each ModuleVersionDataObject                     |
' displayHeaderObjectData            | void            | populates Data to excel sheet by calling printHeaderToColumn in each ModuleVersionDataObject       |
' updateToLatestVersions             | void            | calls updateToNewest in each ModuleVersionDataObject                                               |



' setModVData                  | void                    | sets myData ModuleVersionData                                    |
' getHeader                  | ModuleVersionData()    | gets myData ModuleVersionData                                    |

' getSize                    | Integer         | gets size of ModuleVersionData() array                           |
' getIDS                     | String()        | gets IDs from within ModuleVersionData() array                   |
' FlagNewIDS                 | void            | flags ModuleVersionData that has new IDS                         |
' updateToNewest             | void            | opens the workbook, and updates modules to latest versions|
' identifyOldModuleVersions         | void            | checks if headers are the newest version                  |
' printHeaderToColumn        | void            | Prints data to excel sheet                                |
'-----Private----------------|-----------------|-----------------------------------------------------------|
' populateNewestVersions     | void            | fills newestData() with one of each ModuleVersionData  that contain latest versions  |
' getNewModules               | String          | adds an '!' for each modules not found in myNewestData    |                |
' NewestModulePathGivenID    | String          | given an ID, finds a filepath in myNewestData for the ID  |
' testForNewest              | Boolean         | tests if a module is the newesr                           |
'\---------------------------------------------------------------------------------------------------------/


Private theModuleVersionDataObjects As Collection
Private newestData() As ModuleVersionData ' represents the newest version of each module



' /===========================\
' |Setters and getters        |
' \===========================/


Public Sub setModuleVersionDataObjects(theHDataCol As Collection):    Set theModuleVersionDataObjects = theHDataCol: End Sub


Public Function getTotalNames() As String():    getTotalNames = namesFromPaths(getTotalPaths): End Function

Public Function getTotalPaths() As String()
    Dim totalPaths() As String
    Dim aHeaderObject As ModuleVersionDataObject
    
    For Each aHeaderObject In theModuleVersionDataObjects
        totalPaths = AddToStringArray(totalPaths, aHeaderObject.getWKBKPath)
    Next aHeaderObject

    getTotalPaths = totalPaths
End Function


Public Function getModuleVDataObjByName(tName) As ModuleVersionDataObject
    Dim aHeaderObject As ModuleVersionDataObject
    
    For Each aHeaderObject In theModuleVersionDataObjects
       If nameFromPath(aHeaderObject.getWKBKPath) = tName Then Set getModuleVDataObjByName = aHeaderObject: Exit Function
    Next aHeaderObject
    
End Function


' /===========================\
' |New Module Identification  |
' \===========================/

Public Sub identifyAllOldModules()
    Dim myHD As ModuleVersionDataObject
    Dim mainV() As String ' main versions
    Dim processToNewestData() As ModuleVersionData
        For Each myHD In theModuleVersionDataObjects
            processToNewestData = ConcatenateModuleVersionData(processToNewestData, myHD.getHeader)
            If nameFromPath(myHD.getWKBKPath) = BestModules Then mainV = myHD.getIDS
        Next myHD
        
    Call populateNewestVersions(processToNewestData) 'this generates the newest data
    
        For Each myHD In theModuleVersionDataObjects
            Call myHD.identifyOldModuleVersions(newestData) ' set the new
            Call myHD.FlagNewIDS(mainV)
        Next myHD
End Sub

    Private Sub populateNewestVersions(processToNewestData() As ModuleVersionData)
        Dim ids() As String:  ids = getUniqueID(processToNewestData)
        Dim x As Integer
        Dim bestVersions() As ModuleVersionData
        Dim n As Integer: n = 1
        For x = LBound(ids) To UBound(ids)
            ReDim Preserve bestVersions(1 To n) As ModuleVersionData
            bestVersions(n) = getNewestVersion(ids(x), processToNewestData)
            n = n + 1
        Next x
        newestData = bestVersions
    End Sub
    
        
        Private Function getUniqueID(Arr() As ModuleVersionData) As String()
        Dim x As Integer:        Dim n As Integer: n = 1
        Dim midArr() As String

        If Not ModuleVersionDataHasStuff(Arr) Then GoTo getUniqueIDError:
        
        For x = LBound(Arr) To UBound(Arr) ' error here
        If Arr(x).H_ID <> "" Then
            ReDim Preserve midArr(1 To n) As String
            midArr(n) = Arr(x).H_ID
            n = n + 1
        End If
        Next x

        getUniqueID = removeDupesStringArray(midArr)
    Exit Function
getUniqueIDError:
Dim errSTR(1 To 2) As String
errSTR(1) = "ERR"
       Call reportError("getUniqueID", errSTR)
    
    End Function

    
    ' for a given ID, find the newest version of it among versions
    Private Function getNewestVersion(id As String, Arr() As ModuleVersionData) As ModuleVersionData
        Dim x As Integer
        
        Dim latestMajor As Integer: latestMajor = 0
        Dim latestMinor As Double: latestMinor = 0
        
        Dim newV As ModuleVersionData
        
        For x = LBound(Arr) To UBound(Arr)
            If Arr(x).H_ID = id Then
                If newerVersion(Arr(x).B_MajorVersion, Arr(x).C_MinorVersion, latestMajor, latestMinor) Then
                    latestMajor = Arr(x).B_MajorVersion
                    latestMinor = Arr(x).C_MinorVersion
                    newV = Arr(x)
                End If
            End If
        Next x
        getNewestVersion = newV
    End Function
    ' check if a version is newer than another version
    Private Function newerVersion(tMajor As String, tMinor As String, bestMajor As Integer, bestMinor As Double) As Boolean
        On Error GoTo BADNewerVersion
        Dim testMajor As Integer: testMajor = CInt(tMajor)
        Dim testMinor As Double: testMinor = CDbl(tMinor)
        If testMajor > bestMajor Then newerVersion = True: Exit Function
        If testMajor < bestMajor Then newerVersion = False: Exit Function
        If testMinor > bestMinor Then newerVersion = True
    Exit Function
BADNewerVersion:
            newerVersion = False
    End Function




' /===========================\
' | Module Modification       |
' \===========================/

Public Sub updateToLatestVersions()

    Dim tPaths() As String
    Dim n As Integer: n = 1
    Dim aHDataObj As ModuleVersionDataObject

  If Not theModuleVersionDataObjects Is Nothing Then
    
    For Each aHDataObj In theModuleVersionDataObjects
        aHDataObj.updateToNewest
        ReDim Preserve tPaths(1 To n) As String
        tPaths(n) = aHDataObj.getWKBKPath
        n = n + 1
    Next aHDataObj
       
  Else
    MsgBox "Not Initialised"
  End If
  
End Sub


' /===================================\
' | Aceept Changes Modification       |
' \===================================/


Public Sub acceptRejectChangesInWKBK(tName As String, accept As Boolean)
    Dim theOB As ModuleVersionDataObject
    Set theOB = getModuleVDataObjByName(tName)
    If theOB Is Nothing Then GoTo acceptRejectChangeInWKBKERR
    
    Call theOB.acceptRejectAllModuleChanges(accept)
    refrestAndReprintAll

Exit Sub
acceptRejectChangeInWKBKERR:
    
Dim tErr(1 To 2) As String
    tErr(1) = tName
Call reportError("acceptRejectChangesInWKBK", tErr)
    
End Sub




Public Sub acceptRejectChangesInModule(tName As String, tModName As String, accept As Boolean)
    Dim theOB As ModuleVersionDataObject
    Set theOB = getModuleVDataObjByName(tName)
    If theOB Is Nothing Then GoTo acceptRejectChangesInModuleERR
    
    Call theOB.acceptRejectSingleModuleChanges(tModName, accept)
    
    refrestAndReprintAll
Exit Sub
acceptRejectChangesInModuleERR:
    
Dim tErr(1 To 2) As String
    tErr(1) = tName
    tErr(2) = tModName
Call reportError("acceptRejectChangesInModule", tErr)
    
End Sub








' /===========================\
' |UI/UX methods              |
' \===========================/

    
Public Function makeModuleDisplayByWKBK(wkbkName As String) As String()

Dim aHDataObj As ModuleVersionDataObject

If theModuleVersionDataObjects Is Nothing Then Exit Function
    
For Each aHDataObj In theModuleVersionDataObjects
    If nameFromPath(aHDataObj.getWKBKPath) = wkbkName Then makeModuleDisplayByWKBK = aHDataObj.makeModuleDisplay: Exit Function
Next aHDataObj

End Function
    

' /===========================\
' |Print data to sheet        |
' \===========================/


Public Sub refrestAndReprintAll()
 Dim theHobj As ModuleVersionDataObject

    For Each theHobj In theModuleVersionDataObjects
       Call theHobj.refreshModVerData
    Next theHobj
    Call identifyAllOldModules
    Call displayHeaderObjectData(ThisWorkbook.Sheets(1))
    
End Sub

Public Sub displayHeaderObjectData(theSheet As Worksheet)
    Call clearWorkSpace(theSheet, 1, 6)
    Dim theHobj As ModuleVersionDataObject
    Dim n As Integer: n = 10
    For Each theHobj In theModuleVersionDataObjects
       Call theHobj.printHeaderToColumn(theSheet, 1, n)
       n = n + theHobj.getSize + 2
    Next theHobj
    
End Sub
