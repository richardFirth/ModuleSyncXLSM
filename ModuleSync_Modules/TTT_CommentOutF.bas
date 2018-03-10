Attribute VB_Name = "TTT_CommentOutF"

'/T--TTT_CommentOutF------------------------------------------\
' Function Name              | Return    |  Description       |
'----------------------------|-----------|--------------------|
'TESTcommentoutfunction      | Void      |                    |
'entryPointcommentOut        | Void      |                    |
'commentOutUnused            | String)   |  doesn't work yet  |
'commentOutFunctionInModule  | String()  |                    |
'\------------------------------------------------------------/

Option Explicit

Sub TESTcommentoutfunction()

Dim tpath As String: tpath = "C:\Users\rfirth1\Desktop\TTT_MakeTable.bas"

Dim aMod As New X_SingleModuleObject_1
aMod.initializeModule (tpath)
    Dim tFunctions(1 To 2) As String
    tFunctions(1) = "testUpdateTables"
    tFunctions(2) = "longestString"

Dim newCode() As String
newCode = commentOutFunctionInModule(aMod, tFunctions)

Call createTextFromStringArr(newCode, tpath)

End Sub

Sub entryPointcommentOut()

End Sub

Sub commentOutUnused(tkbk As String, allModules() As String, allFunctions() As String)
' doesn't work yet
    Dim theWKBK As Workbook
    Set theWKBK = Workbooks.Open(tkbk)
    Dim aFPath As String: aFPath = theWKBK.Path & "\Mods"
    Dim aModVDOB As ModuleVersionDataObject
    Set aModVDOB = createModuleHeaderObjectFromWKBK(theWKBK, aFPath)
    
    Dim nWKBK As Workbook
    Set nWKBK = Workbooks.Add
    
    Dim theModuleToUse() As String
    theModuleToUse = removeDupesStringArray(allModules)

    Dim x As Integer
    For x = LBound(theModuleToUse) To UBound(theModuleToUse)
         Dim unusedFinMod() As String
         unusedFinMod = getUnusedFunctionsInModule(aModVDOB.getModuleDataByName(theModuleToUse(x)), allModules, allFunctions)
         Call printStringArrToColumn(unusedFinMod, nWKBK.Sheets(1), x + 1, theModuleToUse(x))
    Next x
    
    nWKBK.Sheets(1).Cells(1, 1).Value = theWKBK.Path & "\" & theWKBK.Name
    
    
    Call nWKBK.SaveAs(theWKBK.Path & "\Unused.xlsx")
    Call nWKBK.Close(False)
    theWKBK.Close
End Sub

Function commentOutFunctionInModule(tMod As X_SingleModuleObject_1, tFunctions() As String) As String()

Dim startSTR() As String: startSTR = tMod.getModuleContents

Dim x As Integer
Dim something() As String

Dim isTheF As Boolean
 Dim n As Integer: n = 1
 For x = LBound(startSTR) To UBound(startSTR)
        If checkForSubOrFunction(startSTR(x)) Then
            Dim y As Integer
            For y = LBound(tFunctions) To UBound(tFunctions)
                If InStr(1, startSTR(x), tFunctions(y)) Then
                    ReDim Preserve something(1 To n) As String
                    something(n) = "' function removed"
                    n = n + 1
                    isTheF = True
                End If
            Next y
        End If
        
        ReDim Preserve something(1 To n) As String
        If isTheF Then
            something(n) = "'" & startSTR(x)
        Else
            something(n) = startSTR(x)
        End If
     
     If checkForEndSubOrFunction(startSTR(x)) Then isTheF = False
     n = n + 1
 Next x
 
 commentOutFunctionInModule = something
 
End Function

