VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HeaderDataObjectsCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private theHeaderDataObjects As Collection
Private newestData() As HeaderData
Private totalPaths() As String


Public Sub setHeaderDataObjects(theHDataCol As Collection)
    Set theHeaderDataObjects = theHDataCol
End Sub
Public Sub setTotalPaths(tPaths() As String)
    totalPaths = tPaths
End Sub




Public Function getTotalPaths() As String()
    Dim tPaths() As String
    
    Dim myHD As HeaderDataObject
    For Each myHD In theHeaderDataObjects
        tPaths = AddToStringArray(tPaths, nameFromPath(myHD.getWKBKPath))
    Next myHD

    getTotalPaths = tPaths
End Function


Public Function getPathsWithData() As String()
    Dim PathsWData() As String
    
    Dim myHD As HeaderDataObject
    For Each myHD In theHeaderDataObjects
        PathsWData = AddToStringArray(PathsWData, myHD.getWKBKDataForDisplay)
    Next myHD

    getPathsWithData = PathsWData

End Function




Public Sub validateHeaderCollections()

Dim myHD As HeaderDataObject

Dim mainV() As String ' main versions

For Each myHD In theHeaderDataObjects
    newestData = ConcatenateHeaderData(newestData, myHD.getHeader)
    If nameFromPath(myHD.getWKBKPath) = "BestModules.xlsm" Then mainV = myHD.getIDS
Next myHD

Call getNewestVersions

For Each myHD In theHeaderDataObjects
    Call myHD.validateTheHeaders(newestData)
    Call myHD.FlagNewIDS(mainV)
Next myHD


End Sub


Public Sub displayHeaderObjectData(theSheet As Worksheet)
    Call clearWorkSpace(theSheet, 1, 6)
        With theSheet
        .Cells(10, 6).Value = "enter x to select"
        .Cells(10, 6).Interior.Color = getRFColor(E_Orange)
        End With
        
    Dim theHobj As HeaderDataObject
    Dim n As Integer: n = 10
    
    For Each theHobj In theHeaderDataObjects
       Call theHobj.printHeaderToColumn(theSheet, 1, n)
       n = n + theHobj.getSize + 2
    Next theHobj
    
End Sub



Public Sub updateToLatestVersions()

    Dim tPaths() As String
    Dim n As Integer: n = 1
    Dim aHDataObj As HeaderDataObject

  If Not theHeaderDataObjects Is Nothing Then
    
    For Each aHDataObj In theHeaderDataObjects
        aHDataObj.updateToNewest
        ReDim Preserve tPaths(1 To n) As String
        tPaths(n) = aHDataObj.getWKBKPath
        n = n + 1
    Next aHDataObj
       
  Else
    MsgBox "Not Initialised"
  End If
  
End Sub








Private Sub getNewestVersions()
    
    Dim ids() As String
    ids = getUniqueID(newestData)
    
    Dim x As Integer
    Dim bestVersions() As HeaderData
    Dim n As Integer: n = 1
    
    For x = LBound(ids) To UBound(ids)
        ReDim Preserve bestVersions(1 To n) As HeaderData
        bestVersions(n) = getNewestVersion(ids(x), newestData)
        n = n + 1
    Next x

    newestData = bestVersions

End Sub



    Private Function getNewestVersion(id As String, Arr() As HeaderData) As HeaderData
        Dim x As Integer
        
        Dim latestMajor As Integer: latestMajor = 0
        Dim latestMinor As Double: latestMinor = 0
        
        Dim newV As HeaderData
        
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
    
    

    Private Function newerVersion(tMajor As String, tMinor As String, bestMajor As Integer, bestMinor As Double) As Boolean
        
        On Error GoTo BADNewerVersion
        
        Dim testMajor As Integer
        Dim testMinor As Double
        
        testMajor = CInt(tMajor)
        testMinor = CDbl(tMinor)
    
        If testMajor > bestMajor Then newerVersion = True: Exit Function
        If testMajor < bestMajor Then newerVersion = False: Exit Function
    
        If testMinor > bestMinor Then newerVersion = True
    
    Exit Function
BADNewerVersion:
            newerVersion = False
    End Function


    Private Function getUniqueID(Arr() As HeaderData) As String()
        Dim x As Integer
        Dim n As Integer: n = 1
        
        Dim midArr() As String
        
        If Not HeaderDataHasStuff(Arr) Then GoTo getUniqueIDError:
        
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
