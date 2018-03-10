VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()

UserForm1.Show

End Sub

Private Sub RunTool_Click()

     'MsgBox "remember that getTXT method changed! - everything isn't working yet"
    ' MsgBox "also automate the date thing PLZ"
    
    If Not UI_OPEN Then
        UI_OPEN = True
        Dim myMU As New ModuleUpdater
        Call myMU.initialSetup
        myMU.Show
    End If
    
End Sub



