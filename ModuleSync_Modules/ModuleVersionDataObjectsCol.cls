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

'/T--ModuleVersionDataObjectsCol-----------------------------------------------------------------------------------------------------\
' Function Name               | Return                   |  Description                                                              |
'-----------------------------|--------------------------|---------------------------------------------------------------------------|
'----- Setters and getters-----------------------------------------------------------------------------------------------------------|
'setModuleVersionDataObjects  | Void                     |  sets the module version data ojects                                      |
'getTotalNames                | String()                 |  gives the names of the workbooks                                         |
'getTotalPaths                | String()                 |  gets the workbook paths from each HeaderDataObject                       |
'getModuleVDataObjByName      | ModuleVersionDataObject  |  gets a module version data object by it's name                           |
'----- New Module Identification-----------------------------------------------------------------------------------------------------|
'identifyAllOldModules        | Void                     |  identify all old modules                                                 |
'~~populateNewestVersions     | Void                     |  populates Newest Versions of modules                                     |
'~~getUniqueID                | String()                 |  gets the total number of unique ID's from an array of moduleVersionData  |
'~~getNewestVersion           | ModuleVersionData        |  for a given ID, find the newest version of it among versions             |
'~~newerVersion               | Boolean                  |  check if a version is newer than another version                         |
'----- Module Modification-----------------------------------------------------------------------------------------------------------|
'updateToLatestVersions       | Void                     |  update modules in each workbook to latest versions                       |
'----- Accept Changes Modification---------------------------------------------------------------------------------------------------|
'acceptRejectChangesInWKBK    | Void                     |  accepts or rejects the changes for all modules in a workbook             |
'acceptRejectChangesInModule  | Void                     |  accepts or rejects the changes for a single modules in a workbook        |
'UpdateTablesInWKBK           | Void                     |  updates all tables in wkbk                                               |
'----- UI/UX methods-----------------------------------------------------------------------------------------------------------------|
'makeModuleDisplayByWKBK      | String()                 | im aHDataObj As ModuleVersionDataObject                                   |
'-----Print data to sheet------------------------------------------------------------------------------------------------------------|
'refrestAndReprintAll         | Void                     | im theHobj As ModuleVersionDataObject                                     |
'displayHeaderObjectData      | Void                     | all clearWorkSpace(theSheet, 1, 6)                                        |
'\-----------------------------------------------------------------------------------------------------------------------------------/

Private theModuleVersionDataObjects As Collection
Private newestData() As ModuleVersionData ' represents the newest version of each module

'/===========================\
'# Setters and getters
'\===========================/

Public Sub setModuleVersionDataObjects(theHDataCol As Collection)
' sets the module version data ojects
Set theModuleVersionDataObjects = theHDataCol
End Sub

Public Function getTotalNames() As String()
' gives the names of the workbooks
getTotalNames = namesFromPaths(getTotalPaths)
End Function

Public Function getTotalPaths() As String()
' gets the workbook paths from each HeaderDataObject
    Dim totalPaths() As String
    Dim aHeaderObject As ModuleVersionDataObject
    For Each aHeaderObject In theModuleVersionDataObjects
        totalPaths = AddToStringArray(totalPaths, aHeaderObject.getWKBKPath)
    Next aHeaderObject
    getTotalPaths = totalPaths
End Function

Public Function getModuleVDataObjByName(tName) As ModuleVersionDataObject
' gets a module version data object by it's name
    Dim aHeaderObject As ModuleVersionDataObject
    For Each aHeaderObject In theModuleVersionDataObjects
       If nameFromPath(aHeaderObject.getWKBKPath) = tName Then Set getModuleVDataObjByName = aHeaderObject: Exit Function
    Next aHeaderObject
End Function

'/============================\
'# New Module Identification
'\============================/

Public Sub identifyAllOldModules()
' identify all old modules
    Dim myHD As ModuleVersionDataObject
    Dim processToNewestData() As ModuleVersionData
        For Each myHD In theModuleVersionDataObjects
            processToNewestData = ConcatenateModuleVersionData(processToNewestData, myHD.getLatestModuleData)
        Next myHD
    Call populateNewestVersions(processToNewestData) 'this generates the newest data
        For Each myHD In theModuleVersionDataObjects
            Call myHD.identifyOldModuleVersions(newestData) ' set the new
        Next myHD
End Sub

Private Sub populateNewestVersions(processToNewestData() As ModuleVersionData)
' populates Newest Versions of modules
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
        ' gets the total number of unique ID's from an array of moduleVersionData
        Dim x As Integer:        Dim n As Integer: n = 1
        Dim midArr() As String
        If Not ModuleVersionDataHasStuff(Arr) Then GoTo getUniqueIDError:
        For x = LBound(Arr) To UBound(Arr) ' error here
        If Arr(x).D_ID <> "" Then
            ReDim Preserve midArr(1 To n) As String
            midArr(n) = Arr(x).D_ID
            n = n + 1
        End If
        Next x
        getUniqueID = removeDupesStringArray(midArr)
    Exit Function
getUniqueIDError:
Dim ERRSTR(1 To 2) As String
ERRSTR(1) = "ERR"
       Call reportError("getUniqueID", ERRSTR)
    End Function

    
    
    Private Function getNewestVersion(id As String, Arr() As ModuleVersionData) As ModuleVersionData
    ' for a given ID, find the newest version of it among versions
        Dim x As Integer
        Dim latestMajor As Integer: latestMajor = 0
        Dim latestMinor As Double: latestMinor = 0
        Dim bestDate As Date: bestDate = #1/31/2018#
        Dim newV As ModuleVersionData
        For x = LBound(Arr) To UBound(Arr)
            If Arr(x).D_ID = id Then
                If newerVersion(Arr(x), latestMajor, latestMinor, bestDate) Then
                    latestMajor = Arr(x).E_MajorVersion
                    latestMinor = Arr(x).F_MinorVersion
                    bestDate = Arr(x).G_date
                    newV = Arr(x)
                End If
            End If
        Next x
        getNewestVersion = newV
    End Function
    
    
    
    Private Function newerVersion(tVer As ModuleVersionData, bestMajor As Integer, bestMinor As Double, bestDate As Date) As Boolean
    ' check if a version is newer than another version
        On Error GoTo BADNewerVersion
        Dim testMajor As Integer: testMajor = CInt(tVer.E_MajorVersion)
        Dim testMinor As Double: testMinor = CDbl(tVer.F_MinorVersion)
        Dim testdate As Date: testdate = CDate(tVer.G_date)
        If testMajor > bestMajor Then newerVersion = True: Exit Function
        If testMajor < bestMajor Then newerVersion = False: Exit Function
        If testMinor > bestMinor Then newerVersion = True
        If testdate > bestDate Then newerVersion = True
    Exit Function
BADNewerVersion:
            newerVersion = False
    End Function

'/===========================\
'# Module Modification
'\===========================/

Public Sub updateToLatestVersions()
' update modules in each workbook to latest versions
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
    refrestAndReprintAll
  Else
    MsgBox "Not Initialised"
  End If
End Sub

'/===================================\
'# Accept Changes Modification
'\===================================/

Public Sub acceptRejectChangesInWKBK(tName As String, accept As Boolean)
' accepts or rejects the changes for all modules in a workbook
    Dim theOB As ModuleVersionDataObject
    Set theOB = getModuleVDataObjByName(tName)
    If theOB Is Nothing Then GoTo acceptRejectChangeInWKBKERR
    Call theOB.acceptRejectAllModuleChanges(accept)
    refrestAndReprintAll
Exit Sub
acceptRejectChangeInWKBKERR:
Dim tERR(1 To 2) As String
    tERR(1) = tName
    tERR(2) = "Object is nothing"
Call reportError("acceptRejectChangesInWKBK", tERR)
End Sub

Public Sub acceptRejectChangesInModule(tName As String, tModName As String, accept As Boolean)
' accepts or rejects the changes for a single modules in a workbook
    Dim theOB As ModuleVersionDataObject
    Set theOB = getModuleVDataObjByName(tName)
    If theOB Is Nothing Then GoTo acceptRejectChangesInModuleERR
    Call theOB.acceptRejectSingleModuleChanges(tModName, accept)
    refrestAndReprintAll
Exit Sub
acceptRejectChangesInModuleERR:
Dim tERR(1 To 2) As String
    tERR(1) = tName
    tERR(2) = tModName
Call reportError("acceptRejectChangesInModule", tERR)
End Sub

Public Sub UpdateTablesInWKBK(tName As String)
' updates all tables in wkbk
    Dim theOB As ModuleVersionDataObject
    Set theOB = getModuleVDataObjByName(tName)
    If theOB Is Nothing Then GoTo UpdateTablesInWKBK
    Call theOB.updateAllTables
    refrestAndReprintAll
Exit Sub
UpdateTablesInWKBK:
Dim tERR(1 To 2) As String
    tERR(1) = tName
    tERR(2) = "Object is nothing"
Call reportError("UpdateTablesInWKBK", tERR)
End Sub

'/===========================\
'# UI/UX methods
'\===========================/

    
Public Function makeModuleDisplayByWKBK(wkbkName As String) As String()
Dim aHDataObj As ModuleVersionDataObject
If theModuleVersionDataObjects Is Nothing Then Exit Function
For Each aHDataObj In theModuleVersionDataObjects
    If nameFromPath(aHDataObj.getWKBKPath) = wkbkName Then makeModuleDisplayByWKBK = aHDataObj.makeModuleDisplay: Exit Function
Next aHDataObj
End Function
    

'/===========================\
'#Print data to sheet
'\===========================/

Public Sub refrestAndReprintAll()
    Dim theHobj As ModuleVersionDataObject
    For Each theHobj In theModuleVersionDataObjects
       Call theHobj.refreshModuleDataFromTXT
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
