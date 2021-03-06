VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIUX_DualListBox_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'$VERSIONCONTROL
'$*MINOR_VERSION*1.6
'$*DATE*2/28/2018*xx
'$*ID*DualListBox
'$*CharCount*5227*xxxx
'$*RowCount*154*xxxx

' this class object is for having a list box where selecting each thing selects stuff in the other list box

'/T--UIUX_DualListBox_1----------------------------------------------------------------------------------------------------------------\
' Function Name          | Return              |  Description                                                                          |
'------------------------|---------------------|---------------------------------------------------------------------------------------|
'initializeDualList      | Void                |  intializes the object with two list boxes                                            |
'refreshSubmenu          | Void                |  refreshes the submenu                                                                |
'displayData             | Void                |  displays the data                                                                    |
'ClearListBoxMenu        | ClearListBoxMenu()  |  clears the list box menu                                                             |
'AddToListBoxMenu        | String)             |  adds to the list box menu                                                            |
'~~getTitles             | String()            |  gets the titles array                                                                |
'~~getSelectedSubMenu    | String()            |  gets the selected item from the submenu                                              |
'~~ConcatenateListBoxMenu| ListBoxMenu()       |  concatentates two listbox menu types                                                 |
'~~ListBoxMenuHasStuff   | Boolean             | https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro  |
'\-------------------------------------------------------------------------------------------------------------------------------------/

Option Explicit

Private Type ListBoxMenu
A_Title As String
B_SubMenu() As String
End Type

Private MainListBox As MSForms.ListBox
Private SecondaryListBox As MSForms.ListBox
Private MenuData() As ListBoxMenu

Public Sub initializeDualList(ByRef tMain As MSForms.ListBox, ByRef tSec As MSForms.ListBox)
' intializes the object with two list boxes
Set MainListBox = tMain
Set SecondaryListBox = tSec
End Sub

Public Sub refreshSubmenu()
' refreshes the submenu
Call PopulateListBoxWithStringArr(SecondaryListBox, getSelectedSubMenu)
End Sub

Public Sub displayData()
' displays the data
Call PopulateListBoxWithStringArr(MainListBox, getTitles)
MainListBox.Selected(0) = True
Call PopulateListBoxWithStringArr(SecondaryListBox, getSelectedSubMenu)

End Sub

' /----------------------------------------\
' |add to string array                     |
' \----------------------------------------/

Public Function ClearListBoxMenu()
' clears the list box menu
Dim newArr() As ListBoxMenu
MenuData = newArr

End Function

Public Function AddToListBoxMenu(theSubMenu() As String, theTitle As String)
' adds to the list box menu
Dim newArr(1 To 1) As ListBoxMenu
newArr(1).A_Title = theTitle
newArr(1).B_SubMenu = theSubMenu
MenuData = ConcatenateListBoxMenu(MenuData, newArr)

End Function

Private Function getTitles() As String()
' gets the titles array
If Not ListBoxMenuHasStuff(MenuData) Then Exit Function

Dim something() As String

Dim x As Integer
Dim n As Integer: n = 1
For x = LBound(MenuData) To UBound(MenuData)
ReDim Preserve something(1 To n) As String
something(n) = MenuData(x).A_Title
n = n + 1
Next x
getTitles = something

End Function

Private Function getSelectedSubMenu() As String()
' gets the selected item from the submenu
If Not ListBoxMenuHasStuff(MenuData) Then Exit Function

Dim tSTR() As String: tSTR() = getSelectedItemsFromListBox(MainListBox)
Dim aSTR As String
If Not arrayHasStuff(tSTR) Then Exit Function

aSTR = tSTR(1)

Dim x As Integer
For x = LBound(MenuData) To UBound(MenuData)
If MenuData(x).A_Title = aSTR Then getSelectedSubMenu = MenuData(x).B_SubMenu: Exit Function
Next x

MsgBox "reached"

End Function

Private Function ConcatenateListBoxMenu(theArray1() As ListBoxMenu, theArray2() As ListBoxMenu) As ListBoxMenu()
' concatentates two listbox menu types
Dim newArr() As ListBoxMenu

Dim n As Long: n = 1

Dim x As Long
If ListBoxMenuHasStuff(theArray1) Then
For x = LBound(theArray1) To UBound(theArray1)
ReDim Preserve newArr(1 To n) As ListBoxMenu
newArr(n) = theArray1(x)
n = n + 1
Next x
End If
If ListBoxMenuHasStuff(theArray2) Then

For x = LBound(theArray2) To UBound(theArray2)
ReDim Preserve newArr(1 To n) As ListBoxMenu
newArr(n) = theArray2(x)
n = n + 1
Next x

End If

ConcatenateListBoxMenu = newArr

End Function

' /-----------------------------------------\
' |returns false when array has nothing     |
' \-----------------------------------------/
Private Function ListBoxMenuHasStuff(theArr() As ListBoxMenu) As Boolean
'https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
If (Not Not theArr) <> 0 Then ListBoxMenuHasStuff = True
End Function
