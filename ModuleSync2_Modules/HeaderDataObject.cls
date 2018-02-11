VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HeaderDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Private myData() As HeaderData
Private myNewestData() As HeaderData

Private WkbkPath As String
Private modulesFolderPath As String

Private oldVersionsFlagged As Boolean



Public Sub setModulesFolderPath(thePath As String)
    modulesFolderPath = thePath
End Sub

Public Sub setWKBKPath(thePath As String)
    WkbkPath = thePath
End Sub

Public Sub setHeader(theD() As HeaderData)
    myData = theD
End Sub




Public Function getWKBKPath() As String
    getWKBKPath = WkbkPath
End Function

Public Function getHeader() As HeaderData()
    getHeader = myData
End Function

Public Function getSize() As Integer
     If HeaderDataHasStuff(myData) Then getSize = UBound(myData)
End Function


Public Function getIDS() As String()
    Dim tID() As String
    Dim x As Integer
    Dim n As Integer: n = 1
    
    If Not HeaderDataHasStuff(myData) Then Exit Function
    
    For x = LBound(myData) To UBound(myData)
        If myData(x).E_Vcontrol Then
            ReDim Preserve tID(1 To n) As String
            tID(n) = myData(x).H_ID
            n = n + 1
        End If
    Next
    getIDS = tID
    
End Function

Public Function FlagNewIDS(currentIDList() As String)
    Dim tID() As String
    
    Dim x As Integer
    
    If Not HeaderDataHasStuff(myData) Then Exit Function
    
    For x = LBound(myData) To UBound(myData)
        If Not stringInArray(myData(x).H_ID, currentIDList) Then
            If myData(x).E_Vcontrol Then myData(x).I_newModule = True
        End If
    Next

End Function




Public Sub updateToNewest()
    If Not oldVersionsFlagged Then Exit Sub
    If Not HeaderDataHasStuff(myData) Then Exit Sub
    
    Dim hasChanged As Boolean
    
    Dim myWKBK As Workbook
    Set myWKBK = Workbooks.Open(WkbkPath)
    
    Dim x As Integer
    
    For x = LBound(myData) To UBound(myData)
        If myData(x).G_OldVersion Then
           If RemoveModuleFromWKBKByName(myWKBK, myData(x).A_Name) Then
              Call ImportModuleToWKBK(myWKBK, modulePathGivenID(myData(x).H_ID))
              hasChanged = True
           End If
        End If
    Next x
    
    If hasChanged Then
        Call myWKBK.Close(True)
    Else
        Call myWKBK.Close(False)
    End If
End Sub


Private Function modulePathGivenID(theID As String) As String

    Dim x As Integer
    
    For x = LBound(myNewestData) To UBound(myNewestData)
        If myNewestData(x).H_ID = theID Then modulePathGivenID = myNewestData(x).F_ModulePath: Exit Function
    Next x

End Function










Public Sub validateTheHeaders(tNewData() As HeaderData)

Dim x As Integer

If HeaderDataHasStuff(myData) Then

    For x = LBound(myData) To UBound(myData)
        If myData(x).E_Vcontrol Then
            If Not testForNewest(myData(x), tNewData) Then
                myData(x).G_OldVersion = True
            End If
        End If
    Next x
    
End If
    myNewestData = tNewData
    oldVersionsFlagged = True
End Sub

Private Function testForNewest(theH As HeaderData, newestData() As HeaderData) As Boolean

    Dim x As Integer
    For x = LBound(newestData) To UBound(newestData)
        If newestData(x).H_ID = theH.H_ID Then
            If newestData(x).B_MajorVersion <> theH.B_MajorVersion Then testForNewest = False: Exit Function
            If newestData(x).C_MinorVersion <> theH.C_MinorVersion Then testForNewest = False: Exit Function
            testForNewest = True: Exit Function
        End If
    Next x
    
End Function




Public Sub printHeaderToColumn(theSheet As Worksheet, theCol As Integer, theRow As Integer)
    
    Dim x As Long: Dim n As Long
    n = theRow + 1
    With theSheet
    
    .Cells(theRow, theCol).Value = nameFromPath(WkbkPath)
    .Cells(theRow, theCol + 1).Value = "Module Path"
    .Cells(theRow, theCol + 2).Value = "Major V"
    .Cells(theRow, theCol + 3).Value = "Minor V"
    .Cells(theRow, theCol + 4).Value = "ID"

     .Range(.Cells(theRow, theCol), .Cells(theRow, theCol + 4)).Interior.Color = getRFColor(T_Grey)
     
     If HeaderDataHasStuff(myData) Then

        For x = LBound(myData) To UBound(myData)
           ' If theArr(x).E_Vcontrol Then
             .Cells(n, theCol).Value = myData(x).A_Name
             .Cells(n, theCol + 1).Value = myData(x).F_ModulePath
             .Cells(n, theCol + 2).Value = myData(x).B_MajorVersion
             .Cells(n, theCol + 3).Value = myData(x).C_MinorVersion
             .Cells(n, theCol + 4).Value = myData(x).H_ID
             If myData(x).E_Vcontrol Then .Range(.Cells(n, theCol), .Cells(n, theCol + 4)).Interior.Color = getRFColor(B_Green)
             If myData(x).G_OldVersion Then .Range(.Cells(n, theCol), .Cells(n, theCol + 4)).Interior.Color = getRFColor(A_Red)
             If myData(x).I_newModule Then .Range(.Cells(n, theCol), .Cells(n, theCol + 4)).Interior.Color = getRFColor(D_Blue)
             n = n + 1
            'End If
        Next x
    
    End If
    
    End With
    
    
End Sub
