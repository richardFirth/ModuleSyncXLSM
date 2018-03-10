Attribute VB_Name = "TTT_CheckForUnused"

'/T--TTT_CheckForUnused------------------------------------------------------------------\
' Function Name              | Return     |  Description                                 |
'----------------------------|------------|----------------------------------------------|
'tagUnusedfunction           | Void       |  tag the unused function                     |
'TagUnusedFunctionsInWKBK    | Workbook)  |  update the dubug log for a single workbook  |
'AddTheTags                  | Void       |  adds the tags                               |
'EntryCheckForUnused         | Void       |  check if functions are unused               |
'checkUnusedFunctionsInWKBK  | String)    |  update the dubug log for a single workbook  |
'getUnusedFunctionsInModule  | String()   |  get the unused function in a module         |
'getFunctionsInModule        | String()   |  get function names in a module              |
'\---------------------------------------------------------------------------------------/

Option Explicit

Sub tagUnusedfunction()
' tag the unused function
    Dim thePath As String: thePath = BrowseFilePath(B_EXCEL)
    Dim dataWKBK As Workbook: Set dataWKBK = Workbooks.Open(thePath)

    Dim theWKBKPath As String: theWKBKPath = dataWKBK.Sheets(1).Cells(1, 1).Value

    Call TagUnusedFunctionsInWKBK(theWKBKPath, dataWKBK)

End Sub

Sub TagUnusedFunctionsInWKBK(tkbk As String, dataWKBK As Workbook)
' update the dubug log for a single workbook

    Dim theWKBK As Workbook
    Set theWKBK = Workbooks.Open(tkbk)
    
    Dim aFPath As String: aFPath = theWKBK.Path & "\Mods"
    Dim aModVDOB As ModuleVersionDataObject
    Set aModVDOB = createModuleHeaderObjectFromWKBK(theWKBK, aFPath)
    Call theWKBK.Close(False)
    Dim y As Integer
    For y = 2 To 40
        If dataWKBK.Sheets(1).Cells(1, y).Value = "" Then Exit For
        Dim FNames() As String
        FNames = getArrayFromColumn(dataWKBK.Sheets(1), y)
        Call AddTheTags(aModVDOB.getModuleDataByName(FNames(1)), FNames)
    Next y
    aModVDOB.commitChangesInAllModuleObjects
    dataWKBK.Close (False)
    
End Sub

Sub AddTheTags(theMOD As X_SingleModuleObject_1, theFNames() As String)
' adds the tags
Dim x As Integer
For x = LBound(theFNames) To UBound(theFNames)
    Call theMOD.z_addCommentToFunction(theFNames(x), "' Function Not Used")
    Call theMOD.saveModule
Next x

End Sub

Sub EntryCheckForUnused()
' check if functions are unused
    Dim thePath As String: thePath = BrowseFilePath(A_CSV)
    Dim FandModData As New ZZZ_CSVLookupTable_1
    Call FandModData.initialSetupFromFile(thePath)
    Dim WKBKPath As String
    WKBKPath = pathFromName(thePath) & FandModData.getAccessKeyForColumn(3)
    Dim allFunc() As String: allFunc = FandModData.getStringArrByName("Function")
    Dim allMod() As String: allMod = FandModData.getStringArrByName("Module")
    Call checkUnusedFunctionsInWKBK(WKBKPath, allMod, allFunc)
End Sub

Sub checkUnusedFunctionsInWKBK(tkbk As String, allModules() As String, allFunctions() As String)
' update the dubug log for a single workbook

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
        Debug.Print "X = " & theModuleToUse(x) & "Y"
         Dim unusedFinMod() As String
         unusedFinMod = getUnusedFunctionsInModule(aModVDOB.getModuleDataByName(theModuleToUse(x)), allModules, allFunctions)
         Call printStringArrToColumn(unusedFinMod, nWKBK.Sheets(1), x + 1, theModuleToUse(x))
    Next x
    
    nWKBK.Sheets(1).Cells(1, 1).Value = theWKBK.Path & "\" & theWKBK.Name
    Call nWKBK.SaveAs(theWKBK.Path & "\Unused.xlsx")
    Call nWKBK.Close(False)
    theWKBK.Close
End Sub

Function getUnusedFunctionsInModule(tMod As X_SingleModuleObject_1, aModules() As String, tFunc() As String) As String()
' get the unused function in a module
    Dim theF() As String: theF = getFunctionsInModule(tMod.getModuleName, aModules, tFunc) ' all that have been called in the module
    
    Dim aSTR() As String: Dim x As Integer
    aSTR = ZgetSubsAndFunctions(tMod.getModuleContents)
    For x = LBound(aSTR) To UBound(aSTR)
         aSTR(x) = getFunctionName(aSTR(x)) ' total functions in module
    Next x

getUnusedFunctionsInModule = DifferenceBetweenSets(aSTR, theF)

End Function

Function getFunctionsInModule(theMOD As String, allModules() As String, allFunctions() As String) As String()
' get function names in a module
 Dim theFunctions() As String
 
 Dim x As Long
 Dim n As Long: n = 1
 
 For x = LBound(allModules) To UBound(allModules)
     If allModules(x) = theMOD Then
        ReDim Preserve theFunctions(1 To n) As String
        theFunctions(n) = allFunctions(x)
        n = n + 1
     End If
 Next x
getFunctionsInModule = theFunctions
End Function

