VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModuleUpdater 
   Caption         =   "Module Updater"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9612
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

Private myModuleCollection As ModuleVersionDataObjectsCol ' ModuleVersionDataObjectCollection object present
Private myDualListbox As UIUX_DualListBox_1

Private modSyncList() As String ' list of modules to sync

Public Sub initialSetup()
    If Not FileThere(ThisWorkbook.Path & "\ModSyncList.txt") Then
        MsgBox "No Default list exists, creating example in this folder to get you started"
        Dim exampleData(1 To 1) As String: exampleData(1) = ThisWorkbook.Path & "\" & BestModules
        Call createTextFromStringArr(exampleData, ThisWorkbook.Path & "\ModSyncList.txt")
        Call complexRoutineEnd("")
        Exit Sub
    End If
   '  If Not FileThere(ThisWorkbook.Path & "\" & BestModules) Then
   '     MsgBox "No " & BestModules & " exists. This file normally holds a copy of the highest version level of each modules"
   '     Call complexRoutineEnd("")
   '     Exit Sub
   ' End If
    updateTableCheckBox.Value = True
End Sub

Private Sub CompareVersions_Click()
    Call compareWithVersons(modSyncList)
End Sub

Private Sub compareWithVersons(thePaths() As String)
    If Not arrayHasStuff(thePaths) Then
        Dim errMessage(1 To 1) As String
        errMessage(1) = "No Modules to Sync"
        Call PopulateListBoxWithStringArr(ListBox1, errMessage)
        Exit Sub
    End If
Call complexRoutineStart("")
    Dim totalPaths() As String
    totalPaths = thePaths
    If stringInArray(ThisWorkbook.Name, namesFromPaths(thePaths)) Then
        MsgBox "Can't Run on self"
        Call complexRoutineEnd("")
        Exit Sub
    End If
    totalPaths = removeDupesStringArray(totalPaths)
    Set myModuleCollection = createModuleObjectsCollection(totalPaths)
    Call myModuleCollection.identifyAllOldModules
    Call myModuleCollection.displayHeaderObjectData(ThisWorkbook.Sheets("VersionControl"))
    Set myDualListbox = New UIUX_DualListBox_1
    Call myDualListbox.initializeDualList(ListBox1, ListBox2)
    Call refreshListBoxes
    UpdateButton.Enabled = True
Call complexRoutineEnd("")
End Sub

Private Sub ListBox1_Click()
If myDualListbox Is Nothing Then Exit Sub
myDualListbox.refreshSubmenu
End Sub

Private Sub updateTables_Click()
Call complexRoutineStart("")
    If myModuleCollection Is Nothing Then MsgBox "Only use once versions are compared!": Exit Sub
    Dim tHL1() As String
    tHL1 = getSelectedItemsFromListBox(ListBox1)
    If Not arrayHasStuff(tHL1) Then MsgBox "No Module Selected": Exit Sub
    Call myModuleCollection.UpdateTablesInWKBK(tHL1(1))
Call complexRoutineEnd("")
End Sub

Private Sub useDefaultList_Click()
    If Not FileThere(ThisWorkbook.Path & "\ModSyncList.txt") Then
        MsgBox "No Default list exists, creating example in this folder to get you started"
        Dim exampleData(1 To 1) As String: exampleData(1) = ThisWorkbook.Path & "\" & BestModules
        Call createTextFromStringArr(exampleData, ThisWorkbook.Path & "\ModSyncList.txt")
        Exit Sub
        Call complexRoutineEnd("")
    End If
    modSyncList = CleanArray(convertTXTDocumentToStringArr(ThisWorkbook.Path & "\ModSyncList.txt"))
    modSyncList = removeBlanksFromArray(modSyncList) ' in case there's an extra enter at end of txt file
    Dim filePresent() As String
    Dim x As Integer
    For x = LBound(modSyncList) To UBound(modSyncList)
        ReDim Preserve filePresent(1 To x) As String
        If Not FileThere(modSyncList(x)) Then
            filePresent(x) = "Missing File"
        ElseIf modSyncList(x) = ThisWorkbook.Path Then
            filePresent(x) = "This workbook!"
        Else
            filePresent(x) = " " ' need a blank space to show up in list box
        End If
    Next x
    Call PopulateListBoxWithStringArr(ListBox1, namesFromPaths(modSyncList))
    Call PopulateListBoxWithStringArr(ListBox2, filePresent)
End Sub

'Private Sub BrowseTo_Click()
'    Dim tPaths() As String
'    tPaths = BrowseFilePaths(D_EXCEL_MACRO)
'   modSyncList = tPaths
'    Call PopulateListBoxWithStringArr(ListBox1, namesFromPaths(modSyncList))
'End Sub

Private Sub UpdateButton_Click()
Call complexRoutineStart("")
 If Not myModuleCollection Is Nothing Then
        myModuleCollection.updateToLatestVersions
        myModuleCollection.refrestAndReprintAll
        Call refreshListBoxes
 Else
    MsgBox "Header Data Not Initialized"
 End If
Call complexRoutineEnd("")
End Sub

Private Sub refreshListBoxes()
    Dim totN() As String: totN = myModuleCollection.getTotalNames
    Dim y As Integer
    myDualListbox.ClearListBoxMenu
    For y = LBound(totN) To UBound(totN)
        Call myDualListbox.AddToListBoxMenu(myModuleCollection.makeModuleDisplayByWKBK(totN(y)), totN(y))
    Next
    myDualListbox.displayData
End Sub

' /================================\
' |accept reject buttons           |
' \================================/

Private Sub AcceptMod_Click()
    Call AcceptRejectModChanges(True)
End Sub

Private Sub AcceptWKBK_Click()
    Call AcceptRejectWKBKChanges(True)
End Sub

Private Sub RejectMod_Click()
' reject changes in a module
    Call AcceptRejectModChanges(False)
End Sub

Private Sub RejectWKBK_Click()
' reject changes in a workbook
    Call AcceptRejectWKBKChanges(False)
End Sub

Private Sub AcceptRejectModChanges(acceptCh As Boolean)
    If myModuleCollection Is Nothing Then MsgBox "Only use once versions are compared!": Exit Sub
    Dim tHL2() As String
    tHL2 = getSelectedItemsFromListBox(ListBox2)
    If Not arrayHasStuff(tHL2) Then MsgBox "No Module Selected": Exit Sub
    Dim tHL1() As String
    tHL1 = getSelectedItemsFromListBox(ListBox1)
    Call myModuleCollection.acceptRejectChangesInModule(tHL1(1), tHL2(1), acceptCh)
    Call refreshListBoxes
End Sub

Private Sub AcceptRejectWKBKChanges(acceptCh As Boolean)
    If myModuleCollection Is Nothing Then MsgBox "Only use once versions are compared!": Exit Sub
    Dim tHL1() As String
    tHL1 = getSelectedItemsFromListBox(ListBox1)
    If Not arrayHasStuff(tHL1) Then MsgBox "No Module Selected": Exit Sub
    Call myModuleCollection.acceptRejectChangesInWKBK(tHL1(1), acceptCh)
    Call refreshListBoxes
End Sub

' /===========================\
' |closing workbook           |
' \===========================/

Private Sub closeButton_Click()
    Call closeSequence
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call closeSequence
    If CloseMode = 0 Then End
End Sub

Private Sub closeSequence()
    Call clearWorkSpace(ThisWorkbook.Sheets(1), 1, 6)
    If Not StoreFilesOnDesktop Then Call EraseExportedFolder
    UI_OPEN = False
End Sub

Private Sub closeButton_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call closeSequence
End Sub

