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
'$*MINOR_VERSION*1.4
'$*DATE*2Feb18
'$*ID*ListSelectA
'$*CharCount*1227*xxxx
'$*RowCount*57*xxxxx

Option Explicit

Private OptionsToSelect() As String
Private selectedOptions() As String

Function getSelectedOptions() As String()
getSelectedOptions = selectedOptions
End Function

Sub setOptionsToSelect(theOptions() As String, multiSelect As Boolean)
OptionsToSelect = theOptions
If multiSelect Then Me.ListBox1.multiSelect = fmMultiSelectMulti

Call PopulateListBoxWithStringArr(Me.ListBox1, OptionsToSelect)
End Sub

Private Sub CButton1_Click()
selectedOptions = getSelectedItemsFromListBox(ListBox1)
Me.Hide
End Sub

Private Sub UserForm_Initialize()
Me.Left = 960 / 2
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then End
End Sub
