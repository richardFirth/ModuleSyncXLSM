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
    Call giveModulesToNewWKBK
End Sub

Private Sub CommandButton2_Click()
    Call RemoveModulesFromWKBK
End Sub


Private Sub RunTool_Click()
    Call ModuleupdaterButtonEntry("")
End Sub






Private Sub CommandButton5_Click()
    Call listModulesInThisWKBK
End Sub





