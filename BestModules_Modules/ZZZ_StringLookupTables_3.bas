Attribute VB_Name = "ZZZ_StringLookupTables_3"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.1
'$*DATE*18Jan18
'$*ID*StringLookupTables
'$*CharCount*2037*xxxx
'$*RowCount*79*xxxxx


'/---ZZZ_StringLookupTables_2----------------------------------------------------\
'  Function Name                   | Return          |   Description                         |
'-----------------------------     |-----------------|---------------------------------------|
'" getLookupValue(theLookupTable | String
' |    |"
' LookupForcolumn(theColumn | void |    |
'\---------------------------------------------------------------------------------/



Option Explicit


Public Type StringLookupTable
A_inputKEY() As String
B_OutputKEY() As String
End Type


Public Function getLookupTable(theSheet As Worksheet, theColumn As Integer) As StringLookupTable

Dim locLookup As StringLookupTable
Dim locInputKey() As String
Dim locOutputKey() As String


With theSheet
Dim x As Long
For x = 1 To .Cells.SpecialCells(xlCellTypeLastCell).Row
If .Cells(x, theColumn).Value = "" Then Exit For
ReDim Preserve locInputKey(1 To x) As String
ReDim Preserve locOutputKey(1 To x) As String
locInputKey(x) = .Cells(x, theColumn).Value
locOutputKey(x) = .Cells(x, theColumn + 1).Value
Next x
End With
locLookup.A_inputKEY = locInputKey
locLookup.B_OutputKEY = locOutputKey

getLookupTable = locLookup

End Function


Public Function getLookupValue(theLookupTable As StringLookupTable, theInputKey As String) As String

Dim x As Integer
For x = LBound(theLookupTable.A_inputKEY) To UBound(theLookupTable.A_inputKEY)
If theLookupTable.A_inputKEY(x) = theInputKey Then getLookupValue = theLookupTable.B_OutputKEY(x): Exit Function
Next x
getLookupValue = "Not Found"

End Function


Public Sub LookupForcolumn(theColumn As Integer, theWorksheet As Worksheet, thelookup As StringLookupTable)

Dim x As Integer
With theWorksheet
For x = 1 To theWorksheet.Cells.SpecialCells(xlCellTypeLastCell).Row
If .Cells(x, theColumn).Value = "" Then Exit For
.Cells(x, theColumn + 1).Value = getLookupValue(thelookup, .Cells(x, theColumn).Value)

Next x
End With

End Sub


