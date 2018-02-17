VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZZ_FolderItemHistory_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'$VERSIONCONTROL
'$*MINOR_VERSION*1.4
'$*DATE*7Feb18
'$*ID*folderItemsCLS



'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@@ Depends on ZZZ_StringArrayFunctions_5 @@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Option Explicit


Private mainFolderPath As String
Private dataListFileName As String
Private setupHappened As Boolean

Private fileExt As String

' SETUP Valid filetypes



Public Sub initializeData(tPath As String, tDatList As String, validExt As String)
    mainFolderPath = tPath
    dataListFileName = tDatList
    setupHappened = True
    fileExt = validExt
    
    If Not FileThere(mainFolderPath & "\" & dataListFileName) Then
        ' initial setup of data?
    End If
    
End Sub


Public Sub commitPresentData()
    Call createTextFromStringArr(ZgetPresentData, mainFolderPath & "\" & dataListFileName)
End Sub


Public Function ZgetPresentData() As String()
    Dim filesSTR() As String: filesSTR = getFilePathsInFolder2Array(mainFolderPath)
    
    Dim validFiles() As String
    
    If Not arrayHasStuff(filesSTR) Then Exit Function
    
    Dim x As Integer
    Dim n As Integer: n = 1
    For x = LBound(filesSTR) To UBound(filesSTR)
        If isAValidFile(filesSTR(x)) Then
            ReDim Preserve validFiles(1 To n) As String
            validFiles(n) = filesSTR(x)
            n = n + 1
        End If
    
    Next x

    ZgetPresentData = validFiles

End Function

Private Function isAValidFile(theFilePath As String) As Boolean
     If UCase(Right(theFilePath, 4)) = UCase(fileExt) Then isAValidFile = True
End Function


Public Function ZgetDataList() As String()

If FileThere(mainFolderPath & "\" & dataListFileName) Then
    ZgetDataList = getTxTDocumentAsString(mainFolderPath & "\" & dataListFileName)
End If

End Function

Public Function getFilesAdded() As String()
      getFilesAdded = DifferenceBetweenSets(ZgetPresentData, ZgetDataList)
End Function

Public Function getFilesRemoved() As String()
   getFilesRemoved = DifferenceBetweenSets(ZgetDataList, ZgetPresentData)
End Function
