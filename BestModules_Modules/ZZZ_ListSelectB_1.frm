VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ZZZ_ListSelectB_1 
   Caption         =   "<OBJ.Caption>"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5436
   OleObjectBlob   =   "ZZZ_ListSelectB_1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ZZZ_ListSelectB_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'$VERSIONCONTROL
'$*MINOR_VERSION*1.7
'$*DATE*2/28/2018*xx
'$*ID*ListSelectB
'$*CharCount*1768*xxxx
'$*RowCount*62*xxxxx

'/T--ZZZ_ListSelectB_1-----------------------------------------------\
' Function Name       | Return  |  Description                       |
'---------------------|---------|------------------------------------|
'getSelectedOption    | String  | etSelectedOption = selectedOption  |
'setOptionsToSelect   | Void    | ptionsToSelect = theCat            |
'CButton1_Click       | Void    | electedOption = ComboBox1.Value    |
'UserForm_Initialize  | Void    | e.Left = 960 / 2                   |
'UserForm_QueryClose  | Void    | f CloseMode = 0 Then End           |
'\-------------------------------------------------------------------/

Option Explicit

Private OptionsToSelect() As String
Private selectedOption As String

Function getSelectedOption() As String
getSelectedOption = selectedOption
End Function

Public Sub setOptionsToSelect(theCat() As String)
OptionsToSelect = theCat
Call PopulateComboBoxWithStringArr(Me.ComboBox1, OptionsToSelect)

End Sub

Private Sub CButton1_Click()
selectedOption = ComboBox1.Value
Me.Hide
End Sub

Private Sub UserForm_Initialize()
Me.Left = 960 / 2
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then End
End Sub

