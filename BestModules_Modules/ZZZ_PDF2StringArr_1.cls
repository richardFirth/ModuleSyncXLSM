VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZZZ_PDF2StringArr_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'$VERSIONCONTROL
'$*MINOR_VERSION*1.2
'$*DATE*15Feb18
'$*ID*pdf2string

' requires Microsoft Word 16.0 Object Library


Option Explicit


Private theObjWord As New Word.Application
Private theData() As String



Public Sub setDataFromPDFAsString(theFormPath As String)

    'Dim theObjWord As Word.Application
    'Set theObjWord = New Word.Application
    On Error GoTo getDataError
    
    theObjWord.Documents.Open theFormPath
    
    Dim certPDFDoc As Word.Document
    Set certPDFDoc = theObjWord.ActiveDocument
    
    
    Dim Arr() As String
    Arr = Split(certPDFDoc.Content.Text, Chr(13))
    
    theData = Arr
    
    Call certPDFDoc.Close(False)
    
    Exit Sub
    
getDataError:

End Sub



Public Function retrieveFullSTRFromPartial(theT As String) As String
    Dim x As Integer
    
    For x = LBound(theData) To UBound(theData)
        If InStr(1, theData(x), theT, vbBinaryCompare) > 0 Then retrieveFullSTRFromPartial = theData(x): Exit Function

    Next x

End Function


Public Function retrieveNextFullSTRFromPartial(theT As String) As String
    Dim x As Integer
    
    For x = LBound(theData) To UBound(theData)
        If InStr(1, theData(x), theT, vbBinaryCompare) > 0 Then retrieveNextFullSTRFromPartial = theData(x) & theData(x + 1) & theData(x + 2) & theData(x + 3): Exit Function

    Next x

End Function
