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
'$*MINOR_VERSION*1.7
'$*DATE*2/28/2018*xx
'$*ID*folderItemsCLS
'$*CharCount*3318*xxxx
'$*RowCount*100*xxxx

'/T--ZZ_FolderItemHistory_2--------------------------------------------------------------------------\
' Function Name     | Return    |  Description                                                       |
'-------------------|-----------|--------------------------------------------------------------------|
'initializeData     | Void      | initialize the class with the folder,datalist, and extention type  |
'commitPresentData  | Void      |  write the datafile to contain the present data                    |
'ZgetPresentData    | String()  |  gets the data presently in the folder                             |
'isAValidFile       | Boolean   |  checks if a file is valid                                         |
'ZgetDataList       | String()  |  gets the data from the data list                                  |
'getFilesAdded      | String()  |  gets the files that have been added                               |
'getFilesRemoved    | String()  |  gets the files that have been removed                             |
'\---------------------------------------------------------------------------------------------------/

Option Explicit

Private mainFolderPath As String
Private dataListFileName As String
Private setupHappened As Boolean

Private fileExt As String

Public Sub initializeData(tPath As String, tDatList As String, validExt As String)
'initialize the class with the folder,datalist, and extention type
mainFolderPath = tPath
dataListFileName = tDatList
setupHappened = True
fileExt = validExt

If Not FileThere(mainFolderPath & "\" & dataListFileName) Then
' initial setup of data?
End If

End Sub

Public Sub commitPresentData()
' write the datafile to contain the present data
Call createTextFromStringArr(ZgetPresentData, mainFolderPath & "\" & dataListFileName)
End Sub

Public Function ZgetPresentData() As String()
' gets the data presently in the folder
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
' checks if a file is valid
If UCase(Right(theFilePath, 4)) = UCase(fileExt) Then isAValidFile = True
End Function

Public Function ZgetDataList() As String()
' gets the data from the data list
If FileThere(mainFolderPath & "\" & dataListFileName) Then
ZgetDataList = getTxTDocumentAsString(mainFolderPath & "\" & dataListFileName)
End If

End Function

Public Function getFilesAdded() As String()
' gets the files that have been added
getFilesAdded = DifferenceBetweenSets(ZgetPresentData, ZgetDataList)
End Function

Public Function getFilesRemoved() As String()
' gets the files that have been removed
getFilesRemoved = DifferenceBetweenSets(ZgetDataList, ZgetPresentData)
End Function
