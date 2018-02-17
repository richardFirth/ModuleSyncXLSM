Attribute VB_Name = "BB_Automatic_Main"
Option Explicit

Public Enum displayMODE
    A_OneList
    B_TwoLists
    C_ThisWorkbook
End Enum




Sub displayAMode(theM As displayMODE)
    If theM = A_OneList Then ThisWorkbook.Sheets("VersionControl").Cells(8, 8).Value = "ONELIST"
    If theM = B_TwoLists Then ThisWorkbook.Sheets("VersionControl").Cells(8, 8).Value = "TWOLIST"
    If theM = C_ThisWorkbook Then ThisWorkbook.Sheets("VersionControl").Cells(8, 8).Value = "THISWORK"
End Sub

Function getDisplayMode() As displayMODE
    If ThisWorkbook.Sheets("VersionControl").Cells(8, 8).Value = "ONELIST" Then getDisplayMode = A_OneList
    If ThisWorkbook.Sheets("VersionControl").Cells(8, 8).Value = "TWOLIST" Then getDisplayMode = B_TwoLists
    If ThisWorkbook.Sheets("VersionControl").Cells(8, 8).Value = "THISWORK" Then getDisplayMode = C_ThisWorkbook
End Function

Function getCurrentWKBKPath() As String
    If ThisWorkbook.Sheets("VersionControl").Cells(9, 8).Value = "" Then
        MsgBox "No Value in the cell": End
    End If
    
    getCurrentWKBKPath = ThisWorkbook.Sheets("VersionControl").Cells(9, 8).Value
End Function




Sub listModulesInOtherWKBK(nused As String)
    Dim otherPath As String: otherPath = BrowseFilePath(D_EXCEL_MACRO)
    
    If otherPath = ThisWorkbook.Path & "\" & ThisWorkbook.Name Then
        MsgBox "The workbook can't be this workbook."
        End
    End If
    
    Dim aWKBK As Workbook: Set aWKBK = Workbooks.Open(otherPath)
    
    Call ListModulesInWKBK(aWKBK)
    
    Call aWKBK.Close(False)
    Call displayAMode(A_OneList)
End Sub


'Sub listModulesInThisWKBK()
'    Call ListModulesInWKBK(ThisWorkbook)
'    Call displayAMode(C_ThisWorkbook)
'End Sub


Sub ListModulesInWKBK(aWKBK As Workbook)

ThisWorkbook.Sheets("VersionControl").Cells(9, 8).Value = aWKBK.Path & "\" & aWKBK.Name

Dim modulePathName As String
modulePathName = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & Left(aWKBK.Name, Len(aWKBK.Name) - 5) & "_Modules"

Dim theHData As HeaderDataObject
Set theHData = getModuleHeaderObjectFromWKBK(aWKBK, modulePathName)

Call clearWorkSpace(ThisWorkbook.Sheets(1), 1, 6)
Call theHData.printHeaderToColumn(ThisWorkbook.Sheets("VersionControl"), 1, 10)
    
End Sub


Sub giveModulesToNewWKBK()

MsgBox "Needs work"

If getDisplayMode = B_TwoLists Then
Else
    MsgBox "You need to have two lists to run this"
    End
End If

Dim theWKBK As Workbook
Set theWKBK = Workbooks.Open(getCurrentWKBKPath)

Dim x As Integer
With ThisWorkbook.Sheets("VersionControl")
    For x = 10 To .Cells.SpecialCells(xlCellTypeLastCell).Row
        If .Cells(x, 1).Value = "" Then Exit For
        If .Cells(x, 6).Value = "x" Then
            Call ImportModuleToWKBK(theWKBK, .Cells(x, 2).Value)
        End If
        
    Next x
End With

Call theWKBK.Close(True)

End Sub




Sub RemoveModulesFromWKBK()

MsgBox "Needs some working on"

If ThisWorkbook.Sheets("VersionControl").Cells(9, 8).Value = ThisWorkbook.Path & "\" & ThisWorkbook.Name Then
    MsgBox "Cannot run on self": Exit Sub
End If

If getDisplayMode = A_OneList Then
Else
    MsgBox "You can only run this on a single list from another workbook"
    End
End If

Dim theWKBK As Workbook
Set theWKBK = Workbooks.Open(getCurrentWKBKPath)

Dim x As Integer
With ThisWorkbook.Sheets("VersionControl")
    For x = 10 To .Cells.SpecialCells(xlCellTypeLastCell).Row
        If .Cells(x, 1).Value = "" Then Exit For
        If .Cells(x, 6).Value = "x" Then
            If RemoveModuleFromWKBKByName(theWKBK, .Cells(x, 1).Value) Then
                .Cells(x, 6).Interior.Color = getRFColor(B_Green)
            Else
                .Cells(x, 6).Interior.Color = getRFColor(A_Red)
            End If
            
        
        End If
    Next x
End With

Call theWKBK.Close(True)

End Sub

