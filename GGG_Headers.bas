Attribute VB_Name = "GGG_Headers"
Option Explicit


Function getOtherModuleHeaders() As HeaderData()
    Dim x As Integer: Dim xx As Integer
    MsgBox "getOtherModuleHeaders - 8Feb18"
    With ThisWorkbook.Sheets("VersionControl")
    
    For x = 11 To .Cells.SpecialCells(xlCellTypeLastCell).Row
        If .Cells(x, 1).Value = "Name" Then
            getOtherModuleHeaders = getHeaderFromNameRow(x)
            Exit For
        End If
    Next x
    

    End With

End Function


Function getHeaderFromNameRow(theNameRow As Integer) As HeaderData()
    Dim n As Integer: n = 1
    Dim x As Integer
    Dim locH() As HeaderData
    With ThisWorkbook.Sheets("VersionControl")
        For x = theNameRow + 1 To .Cells.SpecialCells(xlCellTypeLastCell).Row
            If .Cells(x, 1).Value = "" Then Exit For
            ReDim Preserve locH(1 To n) As HeaderData
            locH(n) = getHeaderFromRow(x)
            n = n + 1
        Next x
    End With
   
    getHeaderFromNameRow = locH

End Function


Function getHeaderFromRow(theRow As Integer) As HeaderData
    Dim locH As HeaderData
    
        With ThisWorkbook.Sheets("VersionControl")
        locH.A_Name = .Cells(theRow, 1).Value
        locH.B_MajorVersion = .Cells(theRow, 3).Value
        locH.C_MinorVersion = .Cells(theRow, 4).Value
        locH.F_ModulePath = .Cells(theRow, 2).Value
        
        locH.H_ID = .Cells(theRow, 5).Value
        If locH.H_ID <> "" Then locH.E_Vcontrol = True
        End With
         
getHeaderFromRow = locH
        
End Function






Public Function ConcatenateHeaderData(theArray1() As HeaderData, theArray2() As HeaderData) As HeaderData()

Dim newArr() As HeaderData

Dim n As Long: n = 1

Dim x As Long
If HeaderDataHasStuff(theArray1) Then
    For x = LBound(theArray1) To UBound(theArray1)
        ReDim Preserve newArr(1 To n) As HeaderData
        newArr(n) = theArray1(x)
        n = n + 1
    Next x
End If
If HeaderDataHasStuff(theArray2) Then

    For x = LBound(theArray2) To UBound(theArray2)
        ReDim Preserve newArr(1 To n) As HeaderData
        newArr(n) = theArray2(x)
        n = n + 1
    Next x

End If

ConcatenateHeaderData = newArr

End Function

Public Function HeaderDataHasStuff(theArr() As HeaderData) As Boolean
'https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro
    If (Not Not theArr) <> 0 Then HeaderDataHasStuff = True
End Function
