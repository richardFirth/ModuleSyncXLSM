Attribute VB_Name = "YYY_CommonUIObject_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.4
'$*DATE*2Feb18
'$*ID*CommonUIObject



Option Explicit

' http://patorjk.com/software/taag/#p=display&f=Soft&t=Type%20Something%20

                                                       
'  ,--.   ,--.        ,--.      ,-----.
'  |  |   `--' ,---.,-'  '-.    |  |) /_  ,---.,--.  ,--.
'  |  |   ,--.(  .-''-.  .-'    |  .-.  \| .-. |\  `'  /
'  |  '--.|  |.-'  `) |  |      |  '--' /' '-' '/  /.  \
'  `-----'`--'`----'  `--'      `------'  `---''--'  '--'
                                                       
Public Sub PopulateListBoxWithStringArr(ByRef tListBox As MSForms.ListBox, theArr() As String)
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



Public Sub DeleteSelectedItemsFromListBox(ByRef tLBox As MSForms.ListBox)

    Dim x As Integer
    For x = tLBox.ListCount - 1 To 0 Step -1
            If tLBox.Selected(x) Then tLBox.RemoveItem (x)
    Next x

End Sub



Public Sub deselectListBox(ByRef tLBox As MSForms.ListBox)
    Dim x As Integer
    For x = tLBox.ListCount - 1 To 0 Step -1
            tLBox.Selected(x) = False
    Next x
End Sub



Public Sub highlightSpecificItemsByArr(ByRef tLBox As MSForms.ListBox, theArr() As String)

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
    With tComboBox
       If Not arrayHasStuff(theArr) Then Exit Sub
            Dim x As Integer
            For x = LBound(theArr) To UBound(theArr)
                If theArr(x) <> "" Then .AddItem theArr(x)
            Next x
    End With
End Sub




