Attribute VB_Name = "YYY_CommonUIObject_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.7
'$*DATE*2/28/2018*xx
'$*ID*CommonUIObject
'$*CharCount*4072*xxxx
'$*RowCount*113*xxxx

'/T--YYY_CommonUIObject_1----------------------------------------------------------------------------------\
' Function Name                  | Return    |  Description                                                |
'--------------------------------|-----------|-------------------------------------------------------------|
'PopulateListBoxWithStringArr    | Void      | populates an array of strings into a list bo                |
'getSelectedItemsFromListBox     | String()  | returns a string array of the selected items in a list box  |
'deleteSelectedItemsFromListBox  | Void      | deletes the selected items from a list box                  |
'deselectListBox                 | Void      | deselects the items in a list box                           |
'highlightSpecificItemsByArr     | Void      | highlights items in a list box by name                      |
'deleteSpecificItemsByArr        | Void      | deletes items from a list box that match an arr             |
'PopulateComboBoxWithStringArr   | Void      |  populates a combo box with a string array                  |
'\---------------------------------------------------------------------------------------------------------/

Option Explicit
' http://patorjk.com/software/taag/#p=display&f=Soft&t=Type%20Something%20

'  ,--.   ,--.        ,--.      ,-----.
'  |  |   `--' ,---.,-'  '-.    |  |) /_  ,---.,--.  ,--.
'  |  |   ,--.(  .-''-.  .-'    |  .-.  \| .-. |\  `'  /
'  |  '--.|  |.-'  `) |  |      |  '--' /' '-' '/  /.  \
'  `-----'`--'`----'  `--'      `------'  `---''--'  '--'

Public Sub PopulateListBoxWithStringArr(ByRef tListBox As MSForms.ListBox, theArr() As String)
'populates an array of strings into a list bo
With tListBox
.Clear
If Not arrayHasStuff(theArr) Then Exit Sub
Dim x As Integer
For x = LBound(theArr) To UBound(theArr)
If theArr(x) <> "" Then .AddItem theArr(x)
Next x
End With
End Sub

Public Function getSelectedItemsFromListBox(ByRef tListBox As MSForms.ListBox) As String()
'returns a string array of the selected items in a list box
Dim selectedOptions() As String
Dim n As Integer: n = 1

Dim zzz As Integer
For zzz = 0 To tListBox.ListCount - 1
If tListBox.Selected(zzz) Then
ReDim Preserve selectedOptions(1 To n) As String
selectedOptions(n) = tListBox.List(zzz)
n = n + 1
End If
Next zzz

getSelectedItemsFromListBox = selectedOptions

End Function

Public Sub deleteSelectedItemsFromListBox(ByRef tLBox As MSForms.ListBox)
'deletes the selected items from a list box
Call deleteSpecificItemsByArr(tLBox, getSelectedItemsFromListBox(tLBox))
End Sub

Public Sub deselectListBox(ByRef tLBox As MSForms.ListBox)
'deselects the items in a list box
Dim x As Integer
For x = tLBox.ListCount - 1 To 0 Step -1
tLBox.Selected(x) = False
Next x
End Sub

Public Sub highlightSpecificItemsByArr(ByRef tLBox As MSForms.ListBox, theArr() As String)
'highlights items in a list box by name
Dim x As Integer

For x = tLBox.ListCount - 1 To 0 Step -1
If stringInArray(tLBox.List(x), theArr) Then
tLBox.Selected(x) = True
Else
tLBox.Selected(x) = False
End If
Next x

End Sub

Public Sub deleteSpecificItemsByArr(ByRef tLBox As MSForms.ListBox, theArr() As String)
'deletes items from a list box that match an arr
Dim x As Integer
For x = tLBox.ListCount - 1 To 0 Step -1
If stringInArray(tLBox.List(x), theArr) Then tLBox.RemoveItem (x)
Next x

End Sub

'  ,-----.                 ,--.              ,-----.
' '  .--./ ,---. ,--,--,--.|  |-.  ,---.     |  |) /_  ,---.,--.  ,--.
' |  |    | .-. ||        || .-. '| .-. |    |  .-.  \| .-. |\  `'  /
' '  '--'\' '-' '|  |  |  || `-' |' '-' '    |  '--' /' '-' '/  /.  \
'  `-----' `---' `--`--`--' `---'  `---'     `------'  `---''--'  '--'

Public Sub PopulateComboBoxWithStringArr(ByRef tComboBox As MSForms.ComboBox, theArr() As String)
' populates a combo box with a string array
With tComboBox
If Not arrayHasStuff(theArr) Then Exit Sub
Dim x As Integer
For x = LBound(theArr) To UBound(theArr)
If theArr(x) <> "" Then .AddItem theArr(x)
Next x
End With
End Sub

