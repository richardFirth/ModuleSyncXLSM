VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModuleUpdater 
   Caption         =   "Module Updater"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7812
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


Public Sub initialSetup()
    EraseExported.Value = True
End Sub



Private Sub CommandButton3_Click()
    
    If EraseExported.Value Then Call EraseExportedFolder
    
    Me.Hide
    
End Sub

Private Sub BrowseTo_Click()
Call complexRoutineStart("")
    Dim tPaths() As String
    tPaths = BrowseFilePaths(D_EXCEL_MACRO)
    Call compareWithVersons(tPaths)
Call complexRoutineEnd("")
End Sub


Private Sub EraseExportedFolder()
    'MsgBox "Erase ModuleSyncOutput"
    Dim tFolder As String: tFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\ModuleSyncOutput"
    If FolderThere(tFolder) Then Call DeleteFolderTreeRF(tFolder)
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub useDefaultList_Click()

Call complexRoutineStart("")
    Dim tPaths() As String
    tPaths = getTxTDocumentAsString(ThisWorkbook.Path & "\DefaultList.txt")
    Call compareWithVersons(tPaths)
Call complexRoutineEnd("")

End Sub

Private Sub UpdateButton_Click()
    
  If Not headerObjectsForComparison Is Nothing Then
  
        headerObjectsForComparison.updateToLatestVersions
        Call compareWithVersons(headerObjectsForComparison.getTotalPaths)
 Else
    MsgBox "Header Data Not Initialized"
 End If
 
End Sub


Private Sub compareWithVersons(thePaths() As String)

    Dim totalPaths() As String
    totalPaths = thePaths
    
    If Not stringInArray(ThisWorkbook.Path & "\BestModules.xlsm", thePaths) Then
        Dim newArr(1 To 1) As String
        newArr(1) = ThisWorkbook.Path & "\BestModules.xlsm"
        
        If Not FileThere(ThisWorkbook.Path & "\BestModules.xlsm") Then
            MsgBox "Can't Locate Best Modules"
            Call complexRoutineEnd("")
            Exit Sub
        End If
        
        totalPaths = ConcatenateArrays(newArr, totalPaths)
    End If
  
    Set headerObjectsForComparison = createHeaderObjectsCollection(totalPaths)
    
    Call headerObjectsForComparison.validateHeaderCollections
    Call headerObjectsForComparison.displayHeaderObjectData(ThisWorkbook.Sheets("VersionControl"))
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If EraseExported.Value Then Call EraseExportedFolder
    If CloseMode = 0 Then End

End Sub
