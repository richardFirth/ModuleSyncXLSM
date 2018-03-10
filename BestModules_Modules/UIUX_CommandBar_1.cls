VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIUX_CommandBar_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'$VERSIONCONTROL
'$*MINOR_VERSION*1.0
'$*DATE*3/5/2018*xxx
'$*ID*CommandBar
'$*CharCount*1711*xxxx
'$*RowCount*55*xxxxx

'/T--UIUX_CommandBar_1--------------------------------------------------\
' Function Name                 | Return|  Description                  |
'-------------------------------|-------|-------------------------------|
'InitializeCommandBarWithItems  | Void  |  initialize the commandBar    |
'removeCommandBar               | Void  |  remove the command bar       |
'AddSingleCommandBarItem        | Void  |  add single command bar item  |
'\----------------------------------------------------------------------/

Option Explicit

Private myCommandBar As CommandBar

Public Sub InitializeCommandBarWithItems()
' initialize the commandBar
    If myCommandBar Is Nothing Then Set myCommandBar = Application.CommandBars.Add("New Bar", msoBarTop, False, True)
    myCommandBar.Visible = True
End Sub

Public Sub removeCommandBar()
' remove the command bar
 If myCommandBar Is Nothing Then Exit Sub
 
    myCommandBar.Visible = False
    myCommandBar.Delete
    Set myCommandBar = Nothing
End Sub

Public Sub AddSingleCommandBarItem(tCaption As String, faceID As Integer, onAction As String)
' add single command bar item

        Dim ctlButton As CommandBarButton
        Set ctlButton = myCommandBar.Controls.Add(msoControlButton)
        ctlButton.Style = msoButtonIconAndCaption
            ctlButton.caption = tCaption
            ctlButton.faceID = faceID
            ctlButton.onAction = onAction
        
End Sub
