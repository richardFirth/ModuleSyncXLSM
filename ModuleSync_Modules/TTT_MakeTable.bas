Attribute VB_Name = "TTT_MakeTable"

'/T--TTT_MakeTable-----------------------------------------------------------------------------------------------------------------\
' Function Name                 | Return               |  Description                                                              |
'-------------------------------|----------------------|---------------------------------------------------------------------------|
'testUpdateTables               | Void                 |  test updating the tables                                                 |
'updateTablesForWorkbook        | Void                 |  Updates the tables for all functions in all modules in a workbook        |
'updateTablesForFolder          | Void                 |  updates all tables                                                       |
'updateTablesForPath            | Void                 |  updates the tables                                                       |
'----- Add and Remove Debug Logs---------------------------------------------------------------------------------------------------|
'updateLogsForWorkbook          | Void                 |  update the dubug log for a single workbook                               |
'createTableFromModuleData      | String()             |  main function. given module data, returns the table at the top           |
'ZgetSubsAndFunctions           | String()             |  gets all subs and functions as a string array                            |
'----- Function Declarations-------------------------------------------------------------------------------------------------------|
'~~ZgetFunctionDeclations       | FunctionDeclation()  |  gets all subs and functions as functionDeclaration                       |
'~~PrintFunctionDeclations      | FunctionDeclation)   |  prints the function declarations to a sheet                              |
'~~getHeadingName               | String               |  gets the name of a Heading                                               |
'~~longestString                | Integer              |  finds the longest string in an array                                     |
'getPubOrPrivName               | pubOrPriv            |  gets if pubOrPriv                                                        |
'~~makeCommentTableEntries      | String()             |  makes the main entries in the table using the functionDeclaration array  |
'~~getStringArrayFromFunctionDec| String()             |  gets the string array from the function declaration                      |
'~~stringToLength               | String               |  adds spaces to a string until it is the desired length                   |
'~~makeSingleTableRow           | String               |  makes a single table row                                                 |
'~~makeHeading                  | String               |  makes a heading                                                          |
'~~makeBreakRow                 | String               |  makes a breakRow                                                         |
'~~makeTableHeaderFooter        | String               |  makes a table header or footer                                           |
'~~bunchOfDash                  | String               |  generates tn number of dashes                                            |
'~~addTabs                      | String               |  adds tabs to a row                                                       |
'\---------------------------------------------------------------------------------------------------------------------------------/

Option Explicit
Public Enum pubOrPriv
    A_PUBLIC
    B_PRIVATE
End Enum

Public Enum WhichType
    A_WantName
    B_WantReturn
    C_WantDesc
End Enum

Public Type FunctionDeclation
    A_SourceSTR As String
    B_Name As String
    C_Return As String
    D_Description As String
    E_Scope As pubOrPriv
    F_isHeader As Boolean
End Type

Public Sub testUpdateTables()
    ' test updating the tables
    Call updateTablesForPath("C:\Users\rfirth1\Desktop\CD_TariffCode.bas")
End Sub








Public Sub updateTablesForWorkbook(tkbk As String)
' Updates the tables for all functions in all modules in a workbook
Call complexRoutineStart("")

    Dim theWKBK As Workbook
    Set theWKBK = Workbooks.Open(tkbk)
    Dim aFPath As String: aFPath = theWKBK.Path & "\Mods"
    Dim aModVDOB As ModuleVersionDataObject
    Set aModVDOB = createModuleHeaderObjectFromWKBK(theWKBK, aFPath)
    theWKBK.Close
    Call aModVDOB.updateAllTables

Call complexRoutineEnd("")

End Sub

Public Sub updateTablesForFolder(tFolder As String)
' updates all tables
Dim tFiles() As PathAndName
tFiles = DetailFilesInFolder2Array(tFolder)
  Dim x As Integer
  For x = LBound(tFiles) To UBound(tFiles)
    Call updateTablesForPath(tFiles(x).A_Path)
  Next x

End Sub

Public Sub updateTablesForPath(tpath As String)
' updates the tables
    Dim tWKBK As New X_SingleModuleObject_1
        tWKBK.initializeModule (tpath)
        tWKBK.z_updateTable
        tWKBK.z_removeDoubleGaps
        tWKBK.saveModule
End Sub

'# Add and Remove Debug Logs
Public Sub updateLogsForWorkbook(tkbk As String, addLog As Boolean)
' update the dubug log for a single workbook

Call complexRoutineStart("")
    Dim theWKBK As Workbook
    Set theWKBK = Workbooks.Open(tkbk)
    Dim aFPath As String: aFPath = theWKBK.Path & "\Mods"
    Dim aModVDOB As ModuleVersionDataObject
    Set aModVDOB = createModuleHeaderObjectFromWKBK(theWKBK, aFPath)
    theWKBK.Close
    Call aModVDOB.updateLogFunctions(addLog)
Call complexRoutineEnd("")

End Sub

Public Function createTableFromModuleData(mData() As String, mName As String) As String()
' main function. given module data, returns the table at the top
    Dim tFunc() As FunctionDeclation
    tFunc = ZgetFunctionDeclations(mData)
    createTableFromModuleData = makeCommentTableEntries(tFunc, mName)
End Function

Public Function ZgetSubsAndFunctions(moduleContents() As String) As String()
' gets all subs and functions as a string array
        Dim tContents() As String
        tContents = TrimAndCleanArray(moduleContents)
        
        Dim n As Integer: n = 1
        Dim gSubsandFunc() As String
        
        Dim x As Integer
        For x = LBound(tContents) To UBound(tContents)
            If checkForSubOrFunction(tContents(x)) Then
                ReDim Preserve gSubsandFunc(1 To n) As String
                gSubsandFunc(n) = tContents(x)
                n = n + 1
            End If
        Next x
        ZgetSubsAndFunctions = gSubsandFunc
End Function

'# Function Declarations

Private Function ZgetFunctionDeclations(moduleContents() As String) As FunctionDeclation()
' gets all subs and functions as functionDeclaration
Dim tContents() As String
tContents = TrimAndCleanArray(moduleContents)

Dim n As Integer: n = 1
Dim gSubsandFunc() As FunctionDeclation
        
Dim x As Integer
For x = LBound(tContents) To UBound(tContents)
    If checkForSubOrFunction(tContents(x)) Then
        ReDim Preserve gSubsandFunc(1 To n) As FunctionDeclation
        gSubsandFunc(n).A_SourceSTR = tContents(x)
        gSubsandFunc(n).B_Name = getFunctionName(tContents(x))
        gSubsandFunc(n).C_Return = getReturnType(tContents(x))
        gSubsandFunc(n).E_Scope = getPubOrPrivName(tContents(x))
        If tContents(x + 1) <> "" Then
        gSubsandFunc(n).D_Description = Right(tContents(x + 1), Len(tContents(x + 1)) - 1)
        End If
        n = n + 1
    End If
    
    If Left(tContents(x), 2) = "'#" Then
        ReDim Preserve gSubsandFunc(1 To n) As FunctionDeclation
        gSubsandFunc(n).B_Name = getHeadingName(tContents(x))
        gSubsandFunc(n).F_isHeader = True
        n = n + 1
    End If
    
Next x
ZgetFunctionDeclations = gSubsandFunc
End Function

Private Sub PrintFunctionDeclations(pDec() As FunctionDeclation)
' prints the function declarations to a sheet
Dim x As Integer
Dim n As Integer: n = 1

Dim tOutput As Workbook
Set tOutput = Workbooks.Add

withtOutput.Sheets (1)
For x = LBound(pDec) To UBound(pDec)
    .Cells(n, 1).Value = pDec(x).A_SourceSTR
    .Cells(n, 2).Value = pDec(x).B_Name
    .Cells(n, 3).Value = pDec(x).C_Return
    .Cells(n, 4).Value = pDec(x).D_Description
    .Cells(n, 5).Value = pDec(x).E_Scope
    n = n + 1
Next x
End With

End Sub

Private Function getHeadingName(theading As String) As String
' gets the name of a Heading
    Dim locs() As String
    locs = Split(theading, "#")
    getHeadingName = locs(1)

End Function

Private Function longestString(tArr() As String) As Integer
' finds the longest string in an array
  Dim biggest As Integer
  Dim x As Integer
  For x = LBound(tArr) To UBound(tArr)
    If biggest < Len(tArr(x)) Then biggest = Len(tArr(x))
  Next x
  
  longestString = biggest

End Function

Public Function getPubOrPrivName(tFDec As String) As pubOrPriv
' gets if pubOrPriv
    Dim SPL1() As String
    SPL1 = Split(tFDec, " ")
    If SPL1(0) = "Private" Then getPubOrPrivName = B_PRIVATE: Exit Function
        
    getPubOrPrivName = A_PUBLIC
   
End Function

Private Function makeCommentTableEntries(tDec() As FunctionDeclation, tblName As String) As String()
' makes the main entries in the table using the functionDeclaration array

Dim FNames() As String: FNames = getStringArrayFromFunctionDec(tDec, A_WantName)
Dim fReturns() As String: fReturns = getStringArrayFromFunctionDec(tDec, B_WantReturn)
Dim fDesc() As String: fDesc = getStringArrayFromFunctionDec(tDec, C_WantDesc)

Dim nameLength As Integer: nameLength = longestString(FNames) + 2
Dim returnLength As Integer: returnLength = longestString(fReturns) + 2
Dim decLength As Integer: decLength = longestString(fDesc) + 2

 Dim x As Integer
 Dim n As Integer: n = 1
 
 Dim tblData() As String
 For x = LBound(tDec) To UBound(tDec)
    ReDim Preserve tblData(1 To n) As String
        If tDec(x).F_isHeader Then
            tblData(n) = makeHeading(tDec(x), nameLength, returnLength, decLength)
        Else
            tblData(n) = makeSingleTableRow(tDec(x), nameLength, returnLength, decLength)
        End If
    n = n + 1
 Next x
 
 Dim headerD As FunctionDeclation
 headerD.B_Name = " Function Name"
 headerD.C_Return = "Return"
 headerD.D_Description = " Description"
 
Dim HDR(1 To 3) As String
HDR(1) = makeTableHeaderFooter(nameLength, returnLength, decLength, True, tblName)
HDR(2) = makeSingleTableRow(headerD, nameLength, returnLength, decLength)
HDR(3) = makeBreakRow(nameLength, returnLength, decLength)

Dim FTR(1 To 1) As String
FTR(1) = makeTableHeaderFooter(nameLength, returnLength, decLength, False, tblName)

tblData = ConcatenateArrays(HDR, tblData)
tblData = ConcatenateArrays(tblData, FTR)

makeCommentTableEntries = tblData

End Function

Private Function getStringArrayFromFunctionDec(tDec() As FunctionDeclation, key As WhichType) As String()
' gets the string array from the function declaration
 Dim x As Integer
 Dim n As Integer: n = 1
 
 Dim wSTR() As String
 
 For x = LBound(tDec) To UBound(tDec)
     If tDec(x).F_isHeader = False Then
        ReDim Preserve wSTR(1 To n) As String
        If key = A_WantName Then wSTR(n) = tDec(x).B_Name
        If key = B_WantReturn Then wSTR(n) = tDec(x).C_Return
        If key = C_WantDesc Then wSTR(n) = tDec(x).D_Description
        n = n + 1
     End If
 Next x

getStringArrayFromFunctionDec = wSTR

End Function

Private Function stringToLength(tSTR As String, desireLen As Integer) As String
' adds spaces to a string until it is the desired length
    Dim locSTR As String: locSTR = tSTR
    Dim x As Integer
    For x = 1 To desireLen
        If Len(locSTR) < desireLen Then
            locSTR = locSTR & " "
        Else
            Exit For
        End If
    Next x
    stringToLength = locSTR
End Function

Private Function makeSingleTableRow(tDec As FunctionDeclation, nameL As Integer, retL As Integer, decL As Integer) As String
' makes a single table row
If tDec.E_Scope = A_PUBLIC Then
    makeSingleTableRow = "'" & stringToLength(tDec.B_Name, nameL) & "| " & stringToLength(tDec.C_Return, retL) & "| " & stringToLength(tDec.D_Description, decL) & "|"
Else
    makeSingleTableRow = "'" & stringToLength("~~" & tDec.B_Name, nameL) & "| " & stringToLength(tDec.C_Return, retL) & "| " & stringToLength(tDec.D_Description, decL) & "|"
End If
End Function

Private Function makeHeading(tDec As FunctionDeclation, nameL As Integer, retL As Integer, decL As Integer) As String
' makes a heading
Dim totalWidth As Integer: totalWidth = nameL + retL + decL

Dim initalSTR As String: initalSTR = "'-----" & tDec.B_Name

    makeHeading = initalSTR & bunchOfDash(5 + totalWidth - Len(initalSTR)) & "|"
End Function

Private Function makeBreakRow(nameL As Integer, retL As Integer, decL As Integer) As String
' makes a breakRow
    makeBreakRow = "'" & bunchOfDash(nameL) & "|-" & bunchOfDash(retL) & "|-" & bunchOfDash(decL) & "|"
End Function

Private Function makeTableHeaderFooter(nameL As Integer, retL As Integer, decL As Integer, header As Boolean, tblName As String) As String
' makes a table header or footer
    Dim tSTR As String
    If header Then
        tSTR = "'/T--" & tblName & bunchOfDash(nameL + retL + decL - Len(tblName)) & "\"
    Else
        tSTR = "'\" & bunchOfDash(nameL + retL + decL + 3) & "/"
    End If
    makeTableHeaderFooter = tSTR
End Function

Private Function bunchOfDash(tn As Integer) As String
' generates tn number of dashes
    Dim lSTR As String: Dim x As Integer
    For x = 1 To tn
        lSTR = lSTR & "-"
    Next x
    bunchOfDash = lSTR
End Function
'4444
Private Function addTabs(n As Integer) As String
' adds tabs to a row
    Dim x As Integer
    Dim lSTR As String
    For x = 1 To 4 * n
        lSTR = lSTR & " "
    Next x
    addTabs = lSTR
End Function

