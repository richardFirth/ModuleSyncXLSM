Attribute VB_Name = "AA_MainComparison"
Option Explicit

Public UI_OPEN As Boolean



Public Sub ModuleupdaterButtonEntry(nused As String)
    
    If Not UI_OPEN Then
    
    UI_OPEN = True
    Dim myMU As New ModuleUpdater
    Call myMU.initialSetup
    myMU.Show
  
    End If

End Sub



