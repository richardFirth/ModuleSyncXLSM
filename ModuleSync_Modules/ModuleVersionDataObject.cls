VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModuleVersionDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Purpose:
' The header data object contains an array of ModuleVersionData that represents the header data of
' all the modules within a single workbook.


'/------ModuleVersionDataObject---------------------------------------------------------------------------------------\
' Factory Function: createModuleObjectFromWKBK(theWKBK As Workbook, theFolderName As String)                          |
'---------------------------------------------------------------------------------------------------------------------|
'  Function Name             | Return               |   Description                                                   |
'----------------------------|----------------------|-----------------------------------------------------------------|
' setModulesFolderPath       | void                 | sets the filepath for the modules folder                        |
' setWKBKPath                | void                 | sets the workbook path                                          |
' getWKBKPath                | String               | gets the workbook path                                          |
' setModVData                | void                 | sets myData ModuleVersionData                                   |
' getHeader                  | ModuleVersionData()  | gets myData ModuleVersionData                                   |
' getSize                    | Integer              | gets size of ModuleVersionData() array                          |
' getIDS                     | String()             | gets IDs from within ModuleVersionData() array                  |
' FlagNewIDS                 | void                 | flags ModuleVersionData that has new IDS                        |
' updateToNewest             | void                 | opens the workbook, and updates modules to latest versions      |
' resetChangeDetect          | void                 | resets change detection                                         |
' identifyOldModuleVersions  | void                 | checks if modules are the newest version                        |
' printHeaderToColumn        | void                 | Prints data to excel sheet                                      |
'-----Private----------------|----------------------|-----------------------------------------------------------------|
' getNumberOfOutdatedModules | Integer              | gets number of modules that may be updated                      |
' getNewModules              | String               | adds an '!' for each modules not found in myNewestData          |
' NewestModulePathGivenID    | String               | given an ID, finds a filepath in myNewestData for the ID        |
' testForNewest              | Boolean              | tests if a module is the newesr                                 |
'\--------------------------------------------------------------------------------------------------------------------/



Private WkbkPath As String           ' the path of the workbook that this object represents
                                     ' doubles as a key for the object

Private myData() As ModuleVersionData       ' the header data for this object
Private myNewestData() As ModuleVersionData ' the header data for the newest data


Private modulesFolderPath As String  ' the path of the folder containing the module
Private modulePaths() As String ' the paths of all the modules


Private oldVersionsFlagged As Boolean ' a flag

Private outPutRow As Integer
Private outPutSheet As Worksheet
Private outPutcolumn As Integer



' /===========================\
' |Setters and getters        |
' \===========================/


Public Sub setModulesFolderPath(thePath As String):    modulesFolderPath = thePath: End Sub
Public Sub setWKBKPath(thePath As String):    WkbkPath = thePath: End Sub
Public Sub setModulePaths(thePaths() As String):    modulePaths = thePaths: End Sub
Public Sub setModVData(theD() As ModuleVersionData):    myData = theD: End Sub

Public Function getHeader() As ModuleVersionData():    getHeader = myData: End Function
Public Function getWKBKPath() As String:    getWKBKPath = WkbkPath: End Function


Public Function getModuleDataByName(tName As String) As ModuleVersionData
    Dim x As Integer
    For x = LBound(myData) To UBound(myData)
        If myData(x).A_Name = tName Then
            getModuleDataByName = myData(x): Exit Function
        End If
    Next x
End Function



' /===========================\
' |New Module Identification  |
' \===========================/
' flips the oldVersion flag in each data object
Public Sub identifyOldModuleVersions(tNewData() As ModuleVersionData)
Dim x As Integer
If ModuleVersionDataHasStuff(myData) Then
    For x = LBound(myData) To UBound(myData)
        If myData(x).E_Vcontrol Then
            If Not testForNewest(myData(x), tNewData) Then
                myData(x).G_OldVersion = True
            End If
        End If
    Next x
End If
    myNewestData = tNewData
    oldVersionsFlagged = True
End Sub

    ' tests module data against the newest version to determine if it's the latest version or not
    Private Function testForNewest(theH As ModuleVersionData, newestData() As ModuleVersionData) As Boolean
        Dim x As Integer
        For x = LBound(newestData) To UBound(newestData)
            If newestData(x).H_ID = theH.H_ID Then
                If newestData(x).B_MajorVersion <> theH.B_MajorVersion Then testForNewest = False: Exit Function
                If newestData(x).C_MinorVersion <> theH.C_MinorVersion Then testForNewest = False: Exit Function
                testForNewest = True: Exit Function
            End If
        Next x
        
    End Function
    
    
' used to get the total list of ID's in this workbook
Public Function getIDS() As String()
    Dim tID() As String
    Dim x As Integer
    Dim n As Integer: n = 1
    If Not ModuleVersionDataHasStuff(myData) Then Exit Function
    For x = LBound(myData) To UBound(myData)
        If myData(x).E_Vcontrol Then
            ReDim Preserve tID(1 To n) As String
            tID(n) = myData(x).H_ID
            n = n + 1
        End If
    Next
    getIDS = tID
    
End Function

' used to identify IDs found in this workbook that are not found in the 'BestModules' workbook
Public Sub FlagNewIDS(currentIDList() As String)
    Dim tID() As String
    Dim x As Integer
    If Not ModuleVersionDataHasStuff(myData) Then Exit Sub
    
    For x = LBound(myData) To UBound(myData)
        If Not stringInArray(myData(x).H_ID, currentIDList) Then
            If myData(x).E_Vcontrol Then myData(x).I_newModule = True
        End If
    Next
End Sub

' /===========================\
' | Module Modification       |
' \===========================/

' opens the workbook and replaces all old modules with new ones
Public Sub updateToNewest()
    If Not oldVersionsFlagged Then Exit Sub
    If Not ModuleVersionDataHasStuff(myData) Then Exit Sub
    Dim hasChanged As Boolean
    Dim myWKBK As Workbook: Set myWKBK = Workbooks.Open(WkbkPath)
    Dim x As Integer
    For x = LBound(myData) To UBound(myData)
        If myData(x).G_OldVersion Then
           If RemoveModuleFromWKBKByName(myWKBK, myData(x).A_Name) Then
              Call ImportModuleToWKBK(myWKBK, NewestModulePathGivenID(myData(x).H_ID))
              hasChanged = True
              Call CopyFileRF(NewestModulePathGivenID(myData(x).H_ID), myData(x).F_ModulePath) ' copies the new module over the old.
           End If
        End If
    Next x
    Call myWKBK.Close(hasChanged)
End Sub

    ' gets the newest module path for a given ID
    Private Function NewestModulePathGivenID(theID As String) As String
        Dim x As Integer
        For x = LBound(myNewestData) To UBound(myNewestData)
            If myNewestData(x).H_ID = theID Then NewestModulePathGivenID = myNewestData(x).F_ModulePath: Exit Function
        Next x
    End Function


' /====================================================\
' | Module Modification - reset change detection       |
' \====================================================/
Public Sub acceptRejectSingleModuleChanges(tName As String, accept As Boolean)
    
    Dim tModule As ModuleVersionData
    tModule = getModuleDataByName(tName)
    
    
    If Not tModule.J_CodeChange Then Exit Sub
    
    Dim hasChanged As Boolean
    Dim myWKBK As Workbook
    Set myWKBK = Workbooks.Open(WkbkPath)
            If RemoveModuleFromWKBKByName(myWKBK, tModule.A_Name) Then
              If accept Then
                Call updateCountsInFile(tModule.F_ModulePath, A_INCREASE)
              Else
                Call updateCountsInFile(tModule.F_ModulePath, B_DECREASE)
              End If
              
              Call ImportModuleToWKBK(myWKBK, tModule.F_ModulePath)
              hasChanged = True
           End If
    Call myWKBK.Close(hasChanged)
    
End Sub

' reset!
Public Sub acceptRejectAllModuleChanges(accept As Boolean)
    
    Dim hasChanged As Boolean
    Dim myWKBK As Workbook
    Set myWKBK = Workbooks.Open(WkbkPath)

    Dim x As Integer
    For x = LBound(myData) To UBound(myData)
        If myData(x).J_CodeChange Then
           If RemoveModuleFromWKBKByName(myWKBK, myData(x).A_Name) Then
              If accept Then
                Call updateCountsInFile(myData(x).F_ModulePath, A_INCREASE)
              Else
                Call updateCountsInFile(myData(x).F_ModulePath, B_DECREASE)
              End If
              Call ImportModuleToWKBK(myWKBK, myData(x).F_ModulePath)
              hasChanged = True
           End If
        End If
    Next x
    
    Call myWKBK.Close(hasChanged)

End Sub


' /===========================\
' |refresh data               |
' \===========================/

Public Sub refreshModVerData()
    ' resets the myData to the latest version of the modules
    myData = extractModVerData(modulePaths)
End Sub



' /===========================\
' |UI/UX methods              |
' \===========================/

    Public Function makeModuleDisplay() As String()
        Dim x As Integer
        Dim n As Integer: n = 1
        Dim tResult() As String
        
        For x = LBound(myData) To UBound(myData)
            If myData(x).E_Vcontrol Then
                If myData(x).G_OldVersion Then
                 '   ReDim Preserve tResult(1 To n) As String
                 '   tResult(n) = myData(x).A_Name & "  -  Old"
                 '   n = n + 1
                ElseIf myData(x).J_CodeChange Then
                    ReDim Preserve tResult(1 To n) As String
                    tResult(n) = myData(x).A_Name
                    n = n + 1
                End If
            End If
        Next x
        
    makeModuleDisplay = tResult
    
    End Function

    Private Function getNumberOfOutdatedModules() As Integer
        Dim x As Integer: Dim n As Integer
        For x = LBound(myData) To UBound(myData)
            If myData(x).G_OldVersion Then
                n = n + 1
            End If
        Next x
    getNumberOfOutdatedModules = n
    End Function
    
    ' not yet used
    Private Function getNumberOfCodeChangeModules() As Integer
        Dim x As Integer: Dim n As Integer
        For x = LBound(myData) To UBound(myData)
            If myData(x).J_CodeChange Then
                n = n + 1
            End If
        Next x
    getNumberOfCodeChangeModules = n
    End Function
    
    
    Private Function getNewModules() As String
        Dim x As Integer
        Dim tResult As String
        For x = LBound(myData) To UBound(myData)
            If myData(x).I_newModule Then
                tResult = tResult & "!"
            End If
        Next x
    getNewModules = " " & tResult
    End Function


' /===========================\
' |Print data to sheet        |
' \===========================/

' helps parent object space out the data
Public Function getSize() As Integer
     If ModuleVersionDataHasStuff(myData) Then getSize = UBound(myData)
End Function


Public Sub printHeaderToColumn(theSheet As Worksheet, theCol As Integer, theRow As Integer)
     outPutRow = theRow
     outPutcolumn = theCol
     Set outPutSheet = theSheet
     Call reprintMyData
End Sub


Public Sub reprintMyData()
    Dim x As Long: Dim n As Long
    n = outPutRow + 1
    With outPutSheet
    .Cells(outPutRow, outPutcolumn).Value = nameFromPath(WkbkPath)
    .Cells(outPutRow, outPutcolumn + 1).Value = "Module Path"
    .Cells(outPutRow, outPutcolumn + 2).Value = "Major V"
    .Cells(outPutRow, outPutcolumn + 3).Value = "Minor V"
    .Cells(outPutRow, outPutcolumn + 4).Value = "ID"
    .Range(.Cells(outPutRow, outPutcolumn), .Cells(outPutRow, outPutcolumn + 4)).Interior.Color = getRFColor(T_Grey)
     
         If ModuleVersionDataHasStuff(myData) Then
            For x = LBound(myData) To UBound(myData)
               ' If theArr(x).E_Vcontrol Then
                 .Cells(n, outPutcolumn).Value = myData(x).A_Name
                 .Cells(n, outPutcolumn + 1).Value = myData(x).F_ModulePath
                 .Cells(n, outPutcolumn + 2).Value = myData(x).B_MajorVersion
                 .Cells(n, outPutcolumn + 3).Value = myData(x).C_MinorVersion
                 .Cells(n, outPutcolumn + 4).Value = myData(x).H_ID
                 If myData(x).E_Vcontrol Then .Range(.Cells(n, outPutcolumn), .Cells(n, outPutcolumn + 4)).Interior.Color = getRFColor(B_Green)
                 If myData(x).J_CodeChange Then .Range(.Cells(n, outPutcolumn), .Cells(n, outPutcolumn + 4)).Interior.Color = getRFColor(C_Yellow)
                 If myData(x).G_OldVersion Then .Range(.Cells(n, outPutcolumn), .Cells(n, outPutcolumn + 4)).Interior.Color = getRFColor(A_Red)
                 If myData(x).I_newModule Then .Range(.Cells(n, outPutcolumn), .Cells(n, outPutcolumn + 4)).Interior.Color = getRFColor(D_Blue)
                 
                 If myData(x).KK_TEMP Then .Range(.Cells(n, outPutcolumn), .Cells(n, outPutcolumn + 4)).Interior.Color = getRFColor(J_Pink)
                 
                 n = n + 1
                'End If
            Next x
        End If
    End With
End Sub

