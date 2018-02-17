VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModuleUpdater 
   Caption         =   "Module Updater"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10728
   OleObjectBlob   =   "ModuleUpdater.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModuleUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private headerObjectsForComparison As HeaderDataObjectsCol
Private modSyncList() As String

Public Sub initialSetup()
    EraseExported.Value = True
    
    If Not FileThere(ThisWorkbook.Path & "\ModSyncList.txt") Then
        MsgBox "No Default list exists, creating example in this folder to get you started"
        Dim exampleData(1 To 1) As String: exampleData(1) = ThisWorkbook.Path & "\BestModules.xlsm"
        Call createTextFromStringArr(exampleData, ThisWorkbook.Path & "\ModSyncList.txt")
        Call complexRoutineEnd("")
        Exit Sub
    End If
    
     If Not FileThere(ThisWorkbook.Path & "\BestModules.xlsm") Then
        MsgBox "No BestModules.xlsm exists. This file normally holds a copy of the highest version level of each modules"
        Call complexRoutineEnd("")
        Exit Sub
    End If

End Sub



Private Sub useDefaultList_Click()
    
    If Not FileThere(ThisWorkbook.Path & "\ModSyncList.txt") Then
        MsgBox "No Default list exists, creating example in this folder to get you started"
        Dim exampleData(1 To 1) As String: exampleData(1) = ThisWorkbook.Path & "\BestModules.xlsm"
        Call createTextFromStringArr(exampleData, ThisWorkbook.Path & "\ModSyncList.txt")
        Exit Sub
        Call complexRoutineEnd("")
    End If
    
    modSyncList = getTxTDocumentAsString(ThisWorkbook.Path & "\ModSyncList.txt")
    
    Dim filePresent() As String
    
    Dim x As Integer
    For x = LBound(modSyncList) To UBound(modSyncList)
        ReDim Preserve filePresent(1 To x) As String
        If FileThere(modSyncList(x)) Then
            filePresent(x) = " "
        Else
            filePresent(x) = "Missing File"
        End If
        
    Next x
    
    Call PopulateListBoxWithStringArr(ListBox1, namesFromPaths(modSyncList))
    Call PopulateListBoxWithStringArr(ListBox2, filePresent)
    
    
    
End Sub

Private Sub BrowseTo_Click()

    Dim tPaths() As String
    tPaths = BrowseFilePaths(D_EXCEL_MACRO)
    modSyncList = tPaths
    Call PopulateListBoxWithStringArr(ListBox1, namesFromPaths(modSyncList))
    
End Sub


Private Sub EraseExportedFolder()
    'MsgBox "Erase ModuleSyncOutput"
    Dim tFolder As String: tFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\ModuleSyncOutput"
    If FolderThere(tFolder) Then Call DeleteFolderTreeRF(tFolder)
    
End Sub


Private Sub closeButton_Click()
    If EraseExported.Value Then Call EraseExportedFolder
    Me.Hide
End Sub


Private Sub CompareVersions_Click()
    If arrayHasStuff(modSyncList) Then
        Call compareWithVersons(modSyncList)
    Else
        Dim errMessage(1 To 1) As String
        errMessage(1) = "No Modules to Sync"
        Call PopulateListBoxWithStringArr(ListBox1, errMessage)
    End If
End Sub


Private Sub UpdateButton_Click()
    
Call complexRoutineStart("")
 If Not headerObjectsForComparison Is Nothing Then
        headerObjectsForComparison.updateToLatestVersions
        Call compareWithVersons(headerObjectsForComparison.getTotalPaths)
 Else
    MsgBox "Header Data Not Initialized"
 End If
Call complexRoutineEnd("")
 
End Sub


Private Sub compareWithVersons(thePaths() As String)

Call complexRoutineStart("")
    Dim totalPaths() As String
    totalPaths = thePaths
    
    If Not stringInArray("BestModules.xlsm", namesFromPaths(thePaths)) Then
        Dim newArr(1 To 1) As String
        newArr(1) = ThisWorkbook.Path & "\BestModules.xlsm"
        If Not FileThere(ThisWorkbook.Path & "\BestModules.xlsm") Then
            MsgBox "Can't Locate Best Modules"
            Call complexRoutineEnd("")
            Exit Sub
        End If
        
        totalPaths = removeDupesStringArray(ConcatenateArrays(newArr, totalPaths))
    End If
  
    Set headerObjectsForComparison = createHeaderObjectsCollection(totalPaths)
    Call headerObjectsForComparison.validateHeaderCollections
    Call headerObjectsForComparison.displayHeaderObjectData(ThisWorkbook.Sheets("VersionControl"))
    Call PopulateListBoxWithStringArr(ListBox1, headerObjectsForComparison.getTotalPaths)
    Call PopulateListBoxWithStringArr(ListBox2, headerObjectsForComparison.getPathsWithData)
    
    
    UpdateButton.Enabled = True
    
    
Call complexRoutineEnd("")
    
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If EraseExported.Value Then Call EraseExportedFolder
    If CloseMode = 0 Then End
End Sub
