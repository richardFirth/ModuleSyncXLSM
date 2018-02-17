VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZZZ_CSVLookupTable_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'$VERSIONCONTROL
'$*MINOR_VERSION*1.2
'$*DATE*13Feb18
'$*ID*CSVTable


Option Explicit


 Private Type LookupWrapper
    A_Title As String
    B_Data() As String
 End Type


Private lookupTable() As LookupWrapper

Private theValues() As String
Private thePath As String


Public Sub initialSetupFromFile(fPath As String)
    thePath = fPath
    If FileThere(thePath) Then Call getCSVFileToLookupTable(thePath)
End Sub



Public Function getStringArrByName(tName As String) As String()

    Dim x As Integer
    
    For x = LBound(lookupTable) To UBound(lookupTable)
        If lookupTable(x).A_Title = tName Then getStringArrByName = lookupTable(x).B_Data: Exit Function
    Next x

End Function



    Public Sub showEntireLookupTable()
        Dim output As Workbook
        Set output = Workbooks.Add
        
        Dim x As Integer
        
        For x = LBound(lookupTable) To UBound(lookupTable)
            Call printStringArrToColumn(lookupTable(x).B_Data, output.Sheets(1), x, lookupTable(x).A_Title)
        Next x

    End Sub


    Private Sub getCSVFileToLookupTable(filePath As String)
    
    Dim newDat As Workbook
    Set newDat = Workbooks.Add(filePath)
    
    Dim localLookupTable() As LookupWrapper
    
    With newDat.Sheets(1)
        Dim y As Integer
        For y = 1 To Cells.SpecialCells(xlCellTypeLastCell).Column
            ReDim Preserve localLookupTable(1 To y) As LookupWrapper
            localLookupTable(y) = getLookupWrapperFromColumn(newDat.Sheets(1), y)
        Next y
    End With
    
    lookupTable = localLookupTable
    
    Call newDat.Close(False)
    End Sub


    Private Function getLookupWrapperFromColumn(theSheet As Worksheet, theColumn As Integer) As LookupWrapper
    
    Dim locWrap As LookupWrapper
    Dim locSTR() As String
    
    With theSheet
    
        
        Dim x As Long
        For x = 1 To .Cells.SpecialCells(xlCellTypeLastCell).Row ' last cell will match overall blanks
                ReDim Preserve locSTR(1 To x) As String
                locSTR(x) = .Cells(x + 1, theColumn).Value ' offset by 1 for title
        Next x
        
        locWrap.B_Data = locSTR
        locWrap.A_Title = .Cells(1, theColumn).Value
    End With
    
        getLookupWrapperFromColumn = locWrap
    End Function
    
    
