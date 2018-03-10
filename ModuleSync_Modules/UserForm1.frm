VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7176
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim tWKBK As String

Private Sub addLogs_Click()
tWKBK = BrowseFilePath(D_EXCEL_MACRO)
Call updateLogsForWorkbook(tWKBK, True)
End Sub

Private Sub checkUnused_Click()
Call EntryCheckForUnused
MsgBox "Check Unused"
End Sub



Private Sub remLogs_Click()
tWKBK = BrowseFilePath(D_EXCEL_MACRO)
Call updateLogsForWorkbook(tWKBK, False)
End Sub



Private Sub tagUnused_Click()
Call tagUnusedfunction
MsgBox "Tag Unused"
Me.Hide
End Sub


Private Sub removeGaps_Click()
    tWKBK = BrowseFilePath(D_EXCEL_MACRO)
    Call removeGapsFromWorkbook(tWKBK)
End Sub

Private Sub updateTables_Click()
    tWKBK = BrowseFilePath(D_EXCEL_MACRO)
    Call updateTablesForWorkbook(tWKBK)
End Sub
