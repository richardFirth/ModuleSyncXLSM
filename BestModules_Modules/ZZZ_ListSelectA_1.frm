VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ZZZ_ListSelectA_1 
   Caption         =   "<OBJ.Caption>"
   ClientHeight    =   3765
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4980
   OleObjectBlob   =   "ZZZ_ListSelectA_1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ZZZ_ListSelectA_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'$VERSIONCONTROL
'$*MINOR_VERSION*1.7
'$*DATE*2/28/2018*xx
'$*ID*ListSelectA
'$*CharCount*2309*xxxx
'$*RowCount*67*xxxxx

'/T--ZZZ_ListSelectA_1----------------------------------------------------------------------------\
' Function Name       | Return    |  Description                                                  |
'---------------------|-----------|---------------------------------------------------------------|
'getSelectedOptions   | String()  |  retrieves selected options                                   |
'setOptionsToSelect   | Void      |  sets options to select                                       |
'CButton1_Click       | Void      | if the button is pushed, set the array to the selected items  |
'UserForm_Initialize  | Void      |  puts the userform in the middle                              |
'UserForm_QueryClose  | Void      |  closes the form                                              |
'\------------------------------------------------------------------------------------------------/

Option Explicit

Private OptionsToSelect() As String
Private selectedOptions() As String

Function getSelectedOptions() As String()
' retrieves selected options
getSelectedOptions = selectedOptions
End Function

Sub setOptionsToSelect(theOptions() As String, multiSelect As Boolean)
' sets options to select
OptionsToSelect = theOptions
If multiSelect Then Me.ListBox1.multiSelect = fmMultiSelectMulti

Call PopulateListBoxWithStringArr(Me.ListBox1, OptionsToSelect)
End Sub

Private Sub CButton1_Click()
'if the button is pushed, set the array to the selected items
selectedOptions = getSelectedItemsFromListBox(ListBox1)
Me.Hide
End Sub

Private Sub UserForm_Initialize()
' puts the userform in the middle
Me.Left = 960 / 2
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' closes the form
If CloseMode = 0 Then End
End Sub
