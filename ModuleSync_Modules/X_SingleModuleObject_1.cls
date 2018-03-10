VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "X_SingleModuleObject_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'/T--X_SingleModuleObject_1-------------------------------------------------------------------------------------------------------------------------\
' Function Name            | Return                 |  Description                                                                                  |
'--------------------------|------------------------|-----------------------------------------------------------------------------------------------|
'----- Getters & Setters----------------------------------------------------------------------------------------------------------------------------|
'getMyModuleVersionData    | ModuleVersionData      |  getter for moduleversiondata                                                                 |
'getModuleName             | String                 |  getter for module name                                                                       |
'getModulePath             | String                 |  getter for module path                                                                       |
'isVersionControl          | Boolean                |  getter for isOldVersion                                                                      |
'moduleID                  | String                 |  getter for module id                                                                         |
'isOldVersion              | Boolean                |  getter for old version                                                                       |
'isCodeChange              | Boolean                |  getter for codeChange                                                                        |
'getModuleContents         | String()               |  getter for module contents                                                                   |
'initializeModule          | Void                   | initialize the module data                                                                    |
'SetIfOldVersion           | ModuleVersionData)     |  tests module data against the newest version to determine if it's the latest version or not  |
'saveModule                | Void                   |  save the module code                                                                         |
'hasCodeChanged            | Boolean                |  return if the code has changed                                                               |
'acceptAsLatestVersion     | Void                   |  accept the module as the latest version                                                      |
'pushBackVersion           | Void                   | push module back a version                                                                    |
'~~updateCountsInFile      | Boolean                | update the version control header to current                                                  |
'~~makeDateSTR             | String                 |  make a string to use as the date                                                             |
'~~makeCharCountSTR        | String                 |  make a string to use as the char count                                                       |
'~~makeRowCountSTR         | String                 |  make a string to use as the Row count                                                        |
'~~makeMinorVersion        | String                 |  make a string to use as the minor version                                                    |
'~~getTrueRowCount         | Integer                |  get the true row count                                                                       |
'~~getWrittenRowCount      | Integer                |  get the written character count                                                              |
'~~getTrueCharCount        | Long                   |  get the true character count                                                                 |
'~~getWrittenCharCount     | Long                   |  get the character count recorded                                                             |
'----- Initial Data Extraction----------------------------------------------------------------------------------------------------------------------|
'refreshModVerDataFromTXT  | Void                   |  entry point for extracting the module version data from a single module                      |
'~~underVersionControl     | Boolean                |  checks if a module is under version control                                                  |
'~~extractVersionName      | String                 |  extracts the name                                                                            |
'~~extractMajorVersion     | String                 |  extracts the major version                                                                   |
'~~extractMinorVer         | String                 |  extracts the minor version                                                                   |
'~~extractDateVer          | String                 |  extracts the date version                                                                    |
'~~extractModuleID         | String                 |  extracts the module ID                                                                       |
'----- Auto Formatting------------------------------------------------------------------------------------------------------------------------------|
'z_updateTable             | Void                   |  updates the function data table                                                              |
'~~z_DeleteTable           | Integer                | Deletes the table                                                                             |
'z_removeDoubleGaps        | z_removeDoubleGaps()   | removes spaces greater than 2                                                                 |
'z_addCommentToFunction    | String)                |  adds a comment to a function                                                                 |
'z_getFunctionLocation     | Integer                |  retrives the location of a function in the module data                                       |
'z_addLogFunction          | z_addLogFunction()     |  adds a logging function to the end of each function                                          |
'z_removeLogFunction       | z_removeLogFunction()  |  removes the function logging thing                                                           |
'\--------------------------------------------------------------------------------------------------------------------------------------------------/

Option Explicit

Private Enum AlterVersion
    A_INCREASE
    B_DECREASE
    C_LEAVEALONE
End Enum

Private myModuleVersionData As ModuleVersionData
'    A_Name As String
'    E_MajorVersion As String
'    F_MinorVersion As String
'    G_date As String
'    C_Vcontrol As Boolean
'    B_ModulePath As String
'    H_OldVersion As Boolean
'    D_ID As String
'    J_CodeChange As Boolean

Private moduleContents() As String

'/--------------------\
'# Getters & Setters
'\--------------------/

Public Function getMyModuleVersionData() As ModuleVersionData
' getter for moduleversiondata
    getMyModuleVersionData = myModuleVersionData
End Function

Public Function getModuleName() As String
' getter for module name
    getModuleName = myModuleVersionData.A_Name
End Function

Public Function getModulePath() As String
' getter for module path
    getModulePath = myModuleVersionData.B_ModulePath
End Function

Public Function isVersionControl() As Boolean
' getter for isOldVersion
    isVersionControl = myModuleVersionData.C_Vcontrol
End Function

Public Function moduleID() As String
' getter for module id
    moduleID = myModuleVersionData.D_ID
End Function

Public Function isOldVersion() As Boolean
' getter for old version
    isOldVersion = myModuleVersionData.H_OldVersion
End Function

Public Function isCodeChange() As Boolean
' getter for codeChange
    isCodeChange = myModuleVersionData.J_CodeChange
End Function

Public Function getModuleContents() As String()
' getter for module contents
    getModuleContents = moduleContents
End Function

Public Sub initializeModule(modPath As String)
'initialize the module data
    myModuleVersionData.B_ModulePath = modPath
    Call refreshModVerDataFromTXT
End Sub

Public Sub SetIfOldVersion(newestData() As ModuleVersionData)
' tests module data against the newest version to determine if it's the latest version or not
If Not myModuleVersionData.C_Vcontrol Then Exit Sub

On Error GoTo setIfOldVErr
    Dim x As Integer
    For x = LBound(newestData) To UBound(newestData)
        If newestData(x).D_ID = myModuleVersionData.D_ID Then
            If CDate(newestData(x).G_date) > CDate(myModuleVersionData.G_date) Then myModuleVersionData.H_OldVersion = True: Exit Sub
            If newestData(x).E_MajorVersion <> myModuleVersionData.E_MajorVersion Then myModuleVersionData.H_OldVersion = True: Exit Sub
            If newestData(x).F_MinorVersion <> myModuleVersionData.F_MinorVersion Then myModuleVersionData.H_OldVersion = True: Exit Sub
            myModuleVersionData.H_OldVersion = False: Exit Sub
        End If
    Next x
Exit Sub
setIfOldVErr:
    Debug.Print "setIfOldVErr"
    myModuleVersionData.H_OldVersion = True
End Sub

Public Sub saveModule(Optional otherPath As String = "")
' save the module code
    If otherPath = "" Then otherPath = myModuleVersionData.B_ModulePath
    Dim nCode() As String
    ' there's an extra newline character at end of createTextFromStringArr,
    ' and we want to make sure the tested values equal the written values
    If moduleContents(UBound(moduleContents)) = "" Then
        Dim y As Integer
        Dim m As Integer: m = 1
        For y = LBound(moduleContents) To UBound(moduleContents) - 1
            ReDim Preserve nCode(1 To m) As String
            nCode(m) = moduleContents(y)
            m = m + 1
        Next y
    Else
        nCode = moduleContents
    End If
    Call createTextFromStringArr(nCode, otherPath)
End Sub

Public Function hasCodeChanged() As Boolean
' return if the code has changed

'If getModuleName = "ZZZ_GenericUseful_2" Then
'    Dim nSTR(1 To 5) As String
'    nSTR(1) = getModuleName
'    nSTR(2) = getTrueRowCount
'    nSTR(3) = getWrittenRowCount
'    nSTR(4) = getTrueCharCount
'    nSTR(5) = getWrittenCharCount
'    Call appendAsRowToCSV(ThisWorkbook.Path & "\codeChangeLog.csv", nSTR)
'End If

If getTrueCharCount <> getWrittenCharCount Then hasCodeChanged = True: Exit Function

If UCase(Right(myModuleVersionData.B_ModulePath, 3)) = "FRM" Then
    If getTrueRowCount - 2 = getWrittenRowCount Then Exit Function
    If getTrueRowCount - 1 = getWrittenRowCount Then Exit Function
    If getTrueRowCount = getWrittenRowCount Then Exit Function
    hasCodeChanged = True
Else
    If getTrueRowCount <> getWrittenRowCount Then hasCodeChanged = True
End If

End Function

Public Sub acceptAsLatestVersion()
' accept the module as the latest version
    Call updateCountsInFile(A_INCREASE)
End Sub

Public Sub pushBackVersion()
'push module back a version
    Call updateCountsInFile(B_DECREASE)
End Sub

Private Function updateCountsInFile(altV As AlterVersion) As Boolean
    'update the version control header to current
    Dim TrueCharCount As Integer: TrueCharCount = getTrueCharCount
    Dim TrueRowCount As Integer: TrueRowCount = getTrueRowCount

    Dim x As Integer
    For x = LBound(moduleContents) To UBound(moduleContents)
    
        If Left(moduleContents(x), 6) = "'$*MIN" Then
            If altV = A_INCREASE Then moduleContents(x) = makeMinorVersion(moduleContents(x), True)
            If altV = B_DECREASE Then moduleContents(x) = makeMinorVersion(moduleContents(x), False)
        End If
        If altV = A_INCREASE Then
            If Left(moduleContents(x), 7) = "'$*DATE" Then
                moduleContents(x) = makeDateSTR()
            End If
        End If
    
        If Left(moduleContents(x), 6) = "'$*Cha" Then
            moduleContents(x) = makeCharCountSTR(TrueCharCount)
        End If
        
        If Left(moduleContents(x), 6) = "'$*Row" Then
            moduleContents(x) = makeRowCountSTR(TrueRowCount)
            Exit For
        End If
    Next x
    updateCountsInFile = True
End Function

    Private Function makeDateSTR() As String
    ' make a string to use as the date
    Dim locSTR As String
        locSTR = "'$*DATE*" & Date & "*"
        
        Dim x As Integer: Dim y As Integer: y = 20 - Len(locSTR)
        For x = 1 To y
            locSTR = locSTR & "x"
        Next x
        makeDateSTR = locSTR
    End Function

    Private Function makeCharCountSTR(aCount As Integer) As String
    ' make a string to use as the char count
        Dim locSTR As String
        locSTR = "'$*CharCount*" & aCount & "*"
        
        Dim x As Integer: Dim y As Integer: y = 22 - Len(locSTR)
        For x = 1 To y
            locSTR = locSTR & "x"
        Next x
        makeCharCountSTR = locSTR
    End Function
    
    Private Function makeRowCountSTR(aCount As Integer) As String
    ' make a string to use as the Row count
        Dim locSTR As String
        locSTR = "'$*RowCount*" & aCount & "*"
        
        Dim x As Integer: Dim y As Integer: y = 20 - Len(locSTR)
        For x = 1 To y
            locSTR = locSTR & "x"
        Next x
        makeRowCountSTR = locSTR
    End Function
    
    Private Function makeMinorVersion(tVer As String, addToMinor As Boolean) As String
    ' make a string to use as the minor version
        Dim loc() As String
        loc = Split(tVer, "*")
        Dim mkmr As String
        Dim newV As Double
        If addToMinor Then newV = loc(2) + 0.1 Else newV = loc(2) - 0.1
        
        mkmr = "'$*MINOR_VERSION*" & newV

        If newV = 1 Then makeMinorVersion = mkmr & ".0": Exit Function
        If newV = 2 Then makeMinorVersion = mkmr & ".0": Exit Function
        If newV = 3 Then makeMinorVersion = mkmr & ".0": Exit Function
        If newV = 4 Then makeMinorVersion = mkmr & ".0": Exit Function
        
        makeMinorVersion = mkmr
    
    End Function

    Private Function getTrueRowCount() As Integer
    ' get the true row count
        getTrueRowCount = UBound(moduleContents)
    End Function
    
    Private Function getWrittenRowCount() As Integer
    ' get the written character count
        Dim x As Integer
        For x = LBound(moduleContents) To UBound(moduleContents)
            If Left(moduleContents(x), 6) = "'$*Row" Then
                Dim getR() As String
                getR = Split(moduleContents(x), "*")
                getWrittenRowCount = getR(2): Exit Function
            End If
        Next x
    End Function
    
    
    Private Function getTrueCharCount() As Long
    ' get the true character count
        Dim x As Integer: Dim tCount As Integer
        For x = LBound(moduleContents) To UBound(moduleContents)
            tCount = tCount + Len(moduleContents(x))
        Next x
        
        getTrueCharCount = tCount
    End Function
    
    Private Function getWrittenCharCount() As Long
    ' get the character count recorded
        Dim x As Integer
        For x = LBound(moduleContents) To UBound(moduleContents)
            If Left(moduleContents(x), 6) = "'$*Cha" Then
                Dim getR() As String
                getR = Split(moduleContents(x), "*")
                getWrittenCharCount = getR(2): Exit Function
            End If
        Next x
    End Function
    
'/--------------------------\
'# Initial Data Extraction
'\--------------------------/
    
Public Sub refreshModVerDataFromTXT()
' entry point for extracting the module version data from a single module
moduleContents = CleanArray(convertTXTDocumentToStringArr(myModuleVersionData.B_ModulePath))

If Not arrayHasStuff(moduleContents) Then
    Dim tERR(1 To 2) As String: tERR(1) = "No Data!"
    tERR(2) = myModuleVersionData.B_ModulePath
    Call reportError("refreshModVerDataFromTXT", tERR)
    Exit Sub
End If

myModuleVersionData.A_Name = extractVersionName(myModuleVersionData.B_ModulePath)
myModuleVersionData.C_Vcontrol = underVersionControl(moduleContents)

    If myModuleVersionData.C_Vcontrol Then
        myModuleVersionData.D_ID = extractModuleID(moduleContents)
        myModuleVersionData.E_MajorVersion = extractMajorVersion(myModuleVersionData.A_Name)
        myModuleVersionData.F_MinorVersion = extractMinorVer(moduleContents)
        myModuleVersionData.G_date = extractDateVer(moduleContents)
        myModuleVersionData.H_OldVersion = False
        myModuleVersionData.J_CodeChange = hasCodeChanged()
    End If
End Sub
    
    
Private Function underVersionControl(theCode() As String) As Boolean
' checks if a module is under version control
  Dim x As Integer
  For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 16) = "'$VERSIONCONTROL" Then underVersionControl = True: Exit Function
        If x = 30 Then Exit Function ' versioncontrol won't be this far down
  Next x
End Function

Private Function extractVersionName(fullFilePath As String) As String
' extracts the name
    Dim zz() As String
    zz = Split(fullFilePath, "\")
    extractVersionName = Left(zz(UBound(zz)), Len(zz(UBound(zz))) - 4)
End Function

Private Function extractMajorVersion(aName As String) As String
' extracts the major version
    Dim zz() As String
    zz = Split(aName, "_")
    extractMajorVersion = zz(UBound(zz))

End Function

Private Function extractMinorVer(theCode() As String) As String
' extracts the minor version
  Dim x As Integer
  For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 8) = "'$*MINOR" Then
            Dim loc() As String
            loc = Split(theCode(x), "*")
            If loc(1) = "MINOR_VERSION" Then extractMinorVer = loc(2): Exit Function
        End If
        If x = 30 Then Exit For ' extractMinorVer won't be this far down
  Next x
   
  extractMinorVer = "NA"
    
End Function

Private Function extractDateVer(theCode() As String) As String
' extracts the date version
  Dim x As Integer
  For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 7) = "'$*DATE" Then
            Dim loc() As String
            loc = Split(theCode(x), "*")
            If loc(1) = "DATE" Then extractDateVer = loc(2): Exit Function
        End If
        If x = 30 Then Exit For ' extractMinorVer won't be this far down
  Next x
  
  extractDateVer = "NA"
    
End Function

Private Function extractModuleID(theCode() As String) As String
' extracts the module ID
  Dim x As Integer
  For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 5) = "'$*ID" Then
            Dim loc() As String
            loc = Split(theCode(x), "*")
            If loc(1) = "ID" Then extractModuleID = loc(2): Exit Function
        End If
        If x = 30 Then Exit For ' extractMinorVer won't be this far down
  Next x
  
  extractModuleID = "NA"
    
End Function

                                                                                                           
'   ,---.            ,--.         ,------.                                 ,--.    ,--.  ,--.
'  /  O  \ ,--.,--.,-'  '-. ,---. |  .---',---. ,--.--.,--,--,--. ,--,--.,-'  '-.,-'  '-.`--',--,--,  ,---.
' |  .-.  ||  ||  |'-.  .-'| .-. ||  `--,| .-. ||  .--'|        |' ,-.  |'-.  .-''-.  .-',--.|      \| .-. |
' |  | |  |'  ''  '  |  |  ' '-' '|  |`  ' '-' '|  |   |  |  |  |\ '-'  |  |  |    |  |  |  ||  ||  |' '-' '
' `--' `--' `----'   `--'   `---' `--'    `---' `--'   `--`--`--' `--`--'  `--'    `--'  `--'`--''--'.`-  /
'                                                                                                    `---'
'/---------------------------\
'# Auto Formatting
'\---------------------------/
    
Public Sub z_updateTable()
' updates the function data table
    Dim x As Long
    x = z_DeleteTable
    If x = -1 Then Exit Sub
    Dim tTable() As String
    tTable = createTableFromModuleData(moduleContents, getModuleName)
    moduleContents = insertArray(moduleContents, tTable, x)
    Debug.Print "Table updated " & myModuleVersionData.A_Name
End Sub
    
Private Function z_DeleteTable() As Integer
'Deletes the table
    Dim x As Integer
    Dim insertPoint As Integer
    Dim deleteStuff As Boolean
    For x = LBound(moduleContents) To UBound(moduleContents)
        If Left(moduleContents(x), 5) = "'/T--" Then
            deleteStuff = True
            insertPoint = x
        End If
        If deleteStuff And Left(moduleContents(x), 1) = "'" Then
            moduleContents(x) = ""
        Else
            If deleteStuff Then
            z_DeleteTable = insertPoint
            Exit Function
            End If
        End If
    Next x
z_DeleteTable = -1
End Function

Public Function z_removeDoubleGaps()
'removes spaces greater than 2
Dim res() As String

 Dim x As Integer
 Dim n As Integer: n = 1
 Dim blankDetect As Boolean
 
    For x = LBound(moduleContents) To UBound(moduleContents) - 1
    If blankDetect And moduleContents(x) = "" Then
    Else
        ReDim Preserve res(1 To n) As String
        res(n) = moduleContents(x)
        n = n + 1
        blankDetect = False
    End If
       If moduleContents(x) = "" Then
          blankDetect = True
       Else
          blankDetect = False
       End If
    Next x
 
 ReDim Preserve res(1 To n) As String
 res(n) = moduleContents(UBound(moduleContents))
 
    moduleContents = res
End Function

Public Sub z_addCommentToFunction(tFunctionName As String, theComment As String)
' adds a comment to a function
Dim tLoc As Integer
tLoc = z_getFunctionLocation(tFunctionName)

If tLoc > 0 Then moduleContents = insertString(moduleContents, theComment, tLoc - 1)

End Sub

Public Function z_getFunctionLocation(tFunctionName As String) As Integer
' retrives the location of a function in the module data
Dim x As Integer
Dim checker() As String: checker = TrimArray(moduleContents)
For x = LBound(checker) To UBound(checker)
    If checkForSubOrFunction(checker(x)) Then
        If getFunctionName(checker(x)) = tFunctionName Then z_getFunctionLocation = x: Exit Function
    End If
Next x
z_getFunctionLocation = -1
End Function

Public Sub z_addLogFunction()
' adds a logging function to the end of each function

  Dim something() As String
  Dim myFunctionName As String
  
  Dim x As Integer
  Dim n As Integer: n = 1
  For x = LBound(moduleContents) To UBound(moduleContents)
    
    If checkForSubOrFunction(moduleContents(x)) Then
        myFunctionName = getFunctionName(moduleContents(x))
    End If
    If InStr(1, moduleContents(x), "End Sub") > 0 Or InStr(1, moduleContents(x), "End Function") Then
        ReDim Preserve something(1 To n) As String
         something(n) = "Call addToDebugLog(""" & myFunctionName & """,""" & myModuleVersionData.A_Name & """)"
         n = n + 1
    End If
    ReDim Preserve something(1 To n) As String
    something(n) = moduleContents(x)
    n = n + 1
  Next x

  moduleContents = something

End Sub

Public Function z_removeLogFunction()
' removes the function logging thing
  Dim newContent() As String
  Dim x As Integer
  Dim n As Integer: n = 1
  For x = LBound(moduleContents) To UBound(moduleContents)
    If Left(moduleContents(x), 18) = "Call addToDebugLog" Then
    Else
        ReDim Preserve newContent(1 To n) As String
        newContent(n) = moduleContents(x)
        n = n + 1
    End If
  Next x
  moduleContents = newContent

End Function



Public Function z_removeGapsInFunctions()
' removes alll empty rows in subs and functions

  Dim x As Integer: Dim newContent() As String
  Dim n As Integer: n = 1

 Dim withinFunction As Boolean
 
For x = LBound(moduleContents) To UBound(moduleContents)
    If checkForSubOrFunction(moduleContents(x)) Then withinFunction = True
    If checkForEndSubOrFunction(moduleContents(x)) Then withinFunction = False

    If withinFunction Then
        If Not Trim(moduleContents(x)) = "" Then
            ReDim Preserve newContent(1 To n) As String
            newContent(n) = moduleContents(x)
            n = n + 1
        Else
            Debug.Print (x)
        End If
    Else
        ReDim Preserve newContent(1 To n) As String
        newContent(n) = moduleContents(x)
        n = n + 1
    End If
Next x

moduleContents = newContent
    
End Function






