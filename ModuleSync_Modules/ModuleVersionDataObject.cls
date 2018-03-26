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

'/T--ModuleVersionDataObject------------------------------------------------------------------------------------------------------------------------------------\
' Function Name                   | Return                      |  Description                                                                                  |
'---------------------------------|-----------------------------|-----------------------------------------------------------------------------------------------|
'----- Setters and getters--------------------------------------------------------------------------------------------------------------------------------------|
'setModulesFolderPath             | Void                        |  sets the modules folder path                                                                 |
'setWKBKPath                      | Void                        |  sets the workbook path                                                                       |
'setModulePaths                   | Void                        |  sets the module paths                                                                        |
'getMySingleModules               | Collection                  |  retrieves the collection of modules                                                          |
'refreshModuleDataFromTXT         | refreshModuleDataFromTXT()  |  refreshes the module data from the text files                                                |
'getLatestModuleData              | ModuleVersionData()         |  gets module version data for each module                                                     |
'getWKBKPath                      | String                      |  gets the workbook path                                                                       |
'getModuleDataByName              | X_SingleModuleObject_1      |  retrieve a moduleObject by it's name                                                         |
'setSingleModules                 | Void                        |  set the collection of single modules                                                         |
'-----New Module Identification---------------------------------------------------------------------------------------------------------------------------------|
'identifyOldModuleVersions        | Void                        |  flips the oldVersion flag in each data object                                                |
'~~testForNewest                  | Boolean                     |  tests module data against the newest version to determine if it's the latest version or not  |
'getIDS                           | String()                    |  used to get the total list of ID's in this workbook                                          |
'----- Module Modification--------------------------------------------------------------------------------------------------------------------------------------|
'updateToNewest                   | Void                        |  opens the workbook and replaces all old modules with new ones                                |
'~~NewestModulePathGivenID        | String                      |  gets the newest module path for a given ID                                                   |
'-----Module Modification - reset change detection--------------------------------------------------------------------------------------------------------------|
'commitChangesInAllModuleObjects  | Void                        |  updates the m                                                                                |
'acceptRejectSingleModuleChanges  | Void                        |  accept or reject changes within a single module                                              |
'acceptRejectAllModuleChanges     | Void                        |  accept or reject all module changes                                                          |
'updateAllTables                  | Void                        |  updates all the tables in all the modules                                                    |
'removeGapsFromfunctions          | Void                        |  removes gaps from functions in module                                                        |
'updateLogFunctions               | Boolean)                    |  updates all the Log Functions in all the modules                                             |
'----- UI/UX methods--------------------------------------------------------------------------------------------------------------------------------------------|
'makeModuleDisplay                | String()                    |  makes a display                                                                              |
'~~getNumberOfCodeChangeModules   | Integer                     |  gets count of modules for whichcode has changed                                              |
'----- Print data to sheet--------------------------------------------------------------------------------------------------------------------------------------|
'getSize                          | Integer                     |  helps parent object space out the data                                                       |
'printHeaderToColumn              | Void                        |  prints a header to a column                                                                  |
'reprintMyData                    | Void                        |  reprints the data to the main sheet                                                          |
'\--------------------------------------------------------------------------------------------------------------------------------------------------------------/

Private WKBKPath As String           ' the path of the workbook that this object represents
                                     ' doubles as a key for the object

Private myNewestData() As ModuleVersionData ' the header data for the newest data
Private mySingleModules As Collection
Private newestModuleData As Collection ' not used yet
Private modulesFolderPath As String  ' the path of the folder containing the module
Private modulePaths() As String ' the paths of all the modules
Private oldVersionsFlagged As Boolean ' a flag
Private outPutRow As Integer
Private outPutSheet As Worksheet
Private outPutcolumn As Integer

' /===========================\
'# Setters and getters
' \===========================/

Public Sub setModulesFolderPath(thePath As String)
' sets the modules folder path
modulesFolderPath = thePath
End Sub

Public Sub setWKBKPath(thePath As String)
' sets the workbook path
WKBKPath = thePath
End Sub

Public Sub setModulePaths(thePaths() As String)
' sets the module paths
modulePaths = thePaths
End Sub

Public Function getMySingleModules() As Collection
' retrieves the collection of modules
    Set getMySingleModules = mySingleModules
End Function

Public Function refreshModuleDataFromTXT()
' refreshes the module data from the text files
    Dim Smod As X_SingleModuleObject_1
    For Each Smod In mySingleModules
        Smod.refreshModVerDataFromTXT
    Next Smod
End Function

Public Function getLatestModuleData() As ModuleVersionData()
' gets module version data for each module
    Dim myData() As ModuleVersionData
    Dim Smod As X_SingleModuleObject_1
    For Each Smod In mySingleModules
        Dim sDat(1 To 1) As ModuleVersionData
        sDat(1) = Smod.getMyModuleVersionData()
        myData = ConcatenateModuleVersionData(myData, sDat)
    Next Smod
    getLatestModuleData = myData
End Function

Public Function getWKBKPath() As String
' gets the workbook path
getWKBKPath = WKBKPath
End Function

Public Function getModuleDataByName(tName As String) As X_SingleModuleObject_1
' retrieve a moduleObject by it's name
    Dim aMod As X_SingleModuleObject_1
    For Each aMod In mySingleModules
        ' compensate for modulename having extention on it
        If aMod.getModuleName = tName Then Set getModuleDataByName = aMod: Exit Function
    Next aMod
End Function

Public Sub setSingleModules(theMods As Collection)
' set the collection of single modules
    Set mySingleModules = theMods
End Sub

' /===========================\
'#New Module Identification
' \===========================/

Public Sub identifyOldModuleVersions(tNewData() As ModuleVersionData)
' flips the oldVersion flag in each data object
    Dim Smod As X_SingleModuleObject_1
    For Each Smod In mySingleModules
        Call Smod.SetIfOldVersion(tNewData)
    Next Smod
    myNewestData = tNewData
    oldVersionsFlagged = True
End Sub

    
Private Function testForNewest(theH As ModuleVersionData, newestData() As ModuleVersionData) As Boolean
' tests module data against the newest version to determine if it's the latest version or not
    Dim x As Integer
    For x = LBound(newestData) To UBound(newestData)
        If newestData(x).D_ID = theH.D_ID Then
            If newestData(x).E_MajorVersion <> theH.E_MajorVersion Then testForNewest = False: Exit Function
            If newestData(x).F_MinorVersion <> theH.F_MinorVersion Then testForNewest = False: Exit Function
            testForNewest = True: Exit Function
        End If
    Next x
End Function
    
    

Public Function getIDS() As String()
' used to get the total list of ID's in this workbook
    Dim tID() As String
    Dim x As Integer
    Dim n As Integer: n = 1
    Dim Smod As X_SingleModuleObject_1
    For Each Smod In mySingleModules
        If Smod.isVersionControl Then
            ReDim Preserve tID(1 To n) As String
            tID(n) = Smod.moduleID
            n = n + 1
        End If
    Next Smod
    getIDS = tID
End Function

'# Module Modification

Public Sub updateToNewest()
' opens the workbook and replaces all old modules with new ones
If Not oldVersionsFlagged Then Exit Sub
Debug.Print "A"
Dim hasChanged As Boolean
Application.EnableEvents = False ' stop workbook from triggering setup macros
Dim myWKBK As Workbook: Set myWKBK = Workbooks.Open(WKBKPath)
Dim importError As Boolean
Dim Smod As X_SingleModuleObject_1
For Each Smod In mySingleModules
    If Smod.isOldVersion Then
         If RemoveModuleFromWKBKByName(myWKBK, Smod.getModuleName) Then
            Dim newPath As String: newPath = NewestModulePathGivenID(Smod.moduleID)
            If Not ImportModuleToWKBK(myWKBK, newPath) Then importError = True
            hasChanged = True
            Call CopyFileRF(newPath, Smod.getModulePath) ' copies the new module over the old.
        End If
    End If
Next Smod
Debug.Print "HasChanged " & hasChanged
If importError Then
    MsgBox "Import Error!" & myWKBK.Name
    Call myWKBK.Close(False)
Else
    Call myWKBK.Close(hasChanged)
End If
Application.EnableEvents = True ' renable workbooks triggering setup macros
End Sub

    
Private Function NewestModulePathGivenID(theID As String) As String
' gets the newest module path for a given ID
Dim x As Integer
For x = LBound(myNewestData) To UBound(myNewestData)
    If myNewestData(x).D_ID = theID Then NewestModulePathGivenID = myNewestData(x).B_ModulePath: Exit Function
Next x
End Function


'#Module Modification - reset change detection

Public Sub commitChangesInAllModuleObjects()
' updates the modules in the workbook to reflect all edits made
    Dim WKBKhasChanged As Boolean
    Application.EnableEvents = False ' stop workbook from triggering setup macros
    Dim myWKBK As Workbook:    Set myWKBK = Workbooks.Open(WKBKPath)
    Dim problem As Boolean
    Dim tModule As X_SingleModuleObject_1
    For Each tModule In mySingleModules
        If RemoveModuleFromWKBKByName(myWKBK, tModule.getModuleName) Then
           If Not ImportModuleToWKBK(myWKBK, tModule.getModulePath) Then problem = True
           WKBKhasChanged = True
        End If
    Next tModule
    If problem Then
        Call myWKBK.Close(False)
        Debug.Print "Problem"
    Else
        Call myWKBK.Close(WKBKhasChanged)
    End If
        Application.EnableEvents = True ' renable workbooks triggering setup macros
End Sub

Public Sub acceptRejectSingleModuleChanges(tName As String, accept As Boolean)
' accept or reject changes within a single module
    Dim tModule As X_SingleModuleObject_1
    Set tModule = getModuleDataByName(tName)
    Dim hasChanged As Boolean
    Dim myWKBK As Workbook
    Application.EnableEvents = False ' stop workbook from triggering setup macros
    Set myWKBK = Workbooks.Open(WKBKPath)
    If RemoveModuleFromWKBKByName(myWKBK, tModule.getModuleName) Then
       If accept Then
         Call tModule.z_updateTable
         Call tModule.z_removeDoubleGaps
         Call tModule.acceptAsLatestVersion
         Call tModule.saveModule
       Else
         Call tModule.pushBackVersion
         Call tModule.saveModule
     End If
      Call ImportModuleToWKBK(myWKBK, tModule.getModulePath)
      hasChanged = True
   End If
Call myWKBK.Close(hasChanged)
Application.EnableEvents = True ' renable workbooks triggering setup macros

End Sub

Public Sub acceptRejectAllModuleChanges(accept As Boolean)
' accept or reject all module changes
Dim tModule As X_SingleModuleObject_1
For Each tModule In mySingleModules
    If tModule.isCodeChange Then
        If accept Then
           Call tModule.z_updateTable
           Call tModule.z_removeDoubleGaps
           Call tModule.acceptAsLatestVersion
        Else
           Call tModule.pushBackVersion
        End If
            Call tModule.saveModule
    End If
Next tModule
Me.commitChangesInAllModuleObjects
End Sub

Public Sub updateAllTables()
' updates all the tables in all the modules
Dim tModule As X_SingleModuleObject_1
For Each tModule In mySingleModules
    If Not tModule.isVersionControl Then
        Call tModule.z_updateTable
        Call tModule.z_removeDoubleGaps
        Call tModule.saveModule
    End If
Next tModule
Me.commitChangesInAllModuleObjects
End Sub

Public Sub removeGapsFromfunctions()
' removes gaps from functions in module
Dim tModule As X_SingleModuleObject_1
For Each tModule In mySingleModules
    If Not tModule.isVersionControl Then
        Call tModule.z_removeGapsInFunctions
        Call tModule.saveModule
    End If
Next tModule
Me.commitChangesInAllModuleObjects
End Sub

Public Sub updateLogFunctions(addLogs As Boolean)
' updates all the Log Functions in all the modules
Dim tModule As X_SingleModuleObject_1
For Each tModule In mySingleModules
    If addLogs Then
        If Not tModule.isVersionControl Then Call tModule.z_addLogFunction
    Else
        If Not tModule.isVersionControl Then Call tModule.z_removeLogFunction
    End If
    Call tModule.saveModule
Next tModule
Me.commitChangesInAllModuleObjects
End Sub

' /===========================\
'# UI/UX methods
' \===========================/

Public Function makeModuleDisplay() As String()
' makes a display
    Dim x As Integer
    Dim n As Integer: n = 1
    Dim tResult() As String
Dim Smod As X_SingleModuleObject_1
For Each Smod In mySingleModules
    If Smod.isVersionControl And Smod.isCodeChange Then
        ReDim Preserve tResult(1 To n) As String
        tResult(n) = Smod.getModuleName
        n = n + 1
    End If
Next Smod
makeModuleDisplay = tResult
End Function

    
Private Function getNumberOfCodeChangeModules() As Integer
' gets count of modules for whichcode has changed
Dim tCount As Integer
For Each Smod In mySingleModules
    If Smod.isCodeChange Then tCount = tCount + 1
Next Smod
getNumberOfCodeChangeModules = tCount
End Function
    
    

' /===========================\
'# Print data to sheet
' \===========================/

Public Function getSize() As Integer
' helps parent object space out the data
     If ModuleVersionDataHasStuff(getLatestModuleData) Then getSize = UBound(getLatestModuleData)
End Function

Public Sub printHeaderToColumn(theSheet As Worksheet, theCol As Integer, theRow As Integer)
' prints a header to a column
     outPutRow = theRow
     outPutcolumn = theCol
     Set outPutSheet = theSheet
     Call reprintMyData
End Sub

Public Sub reprintMyData()
' reprints the data to the main sheet
    Dim x As Long: Dim n As Long
    n = outPutRow + 1
    With outPutSheet
    .Cells(outPutRow, outPutcolumn).Value = nameFromPath(WKBKPath)
    .Cells(outPutRow, outPutcolumn + 1).Value = "Module Path"
    .Cells(outPutRow, outPutcolumn + 2).Value = "Major V"
    .Cells(outPutRow, outPutcolumn + 3).Value = "Minor V"
    .Cells(outPutRow, outPutcolumn + 4).Value = "ID"
    .Cells(outPutRow, outPutcolumn + 5).Value = "Date"
    .Range(.Cells(outPutRow, outPutcolumn), .Cells(outPutRow, outPutcolumn + 5)).Interior.Color = getRFColor(T_Grey)
        Dim Smod As X_SingleModuleObject_1
        For Each Smod In mySingleModules
             .Cells(n, outPutcolumn).Value = Smod.getModuleName
             .Cells(n, outPutcolumn + 1).Value = Smod.getModulePath
             .Cells(n, outPutcolumn + 2).Value = Smod.getMyModuleVersionData.E_MajorVersion
             .Cells(n, outPutcolumn + 3).Value = Smod.getMyModuleVersionData.F_MinorVersion
             .Cells(n, outPutcolumn + 4).Value = Smod.moduleID
             .Cells(n, outPutcolumn + 5).Value = Smod.getMyModuleVersionData.G_date
             If Smod.isVersionControl Then .Range(.Cells(n, outPutcolumn), .Cells(n, outPutcolumn + 5)).Interior.Color = getRFColor(B_Green)
             If Smod.isCodeChange Then .Range(.Cells(n, outPutcolumn), .Cells(n, outPutcolumn + 5)).Interior.Color = getRFColor(C_Yellow)
             If Smod.isOldVersion Then .Range(.Cells(n, outPutcolumn), .Cells(n, outPutcolumn + 5)).Interior.Color = getRFColor(A_Red)
             n = n + 1
        Next Smod
    End With
End Sub

