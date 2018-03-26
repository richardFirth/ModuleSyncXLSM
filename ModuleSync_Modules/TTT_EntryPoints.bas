Attribute VB_Name = "TTT_EntryPoints"
Option Explicit

'/T--TTT_EntryPoints------------------------------------------------------------------------------------\
' Function Name           | Return|  Description                                                        |
'-------------------------|-------|---------------------------------------------------------------------|
'updateTablesForWorkbook  | Void  |  Updates the tables for all functions in all modules in a workbook  |
'updateLogsForWorkbook    | Void  |  update the dubug log for a single workbook                         |
'removeGapsFromWorkbook   | Void  |  removes gaps from within all functions in workbook                 |
'EntryCheckForUnused      | Void  |  check if functions are unused                                      |
'\------------------------------------------------------------------------------------------------------/

Public Sub updateTablesForWorkbook(tkbk As String)
' Updates the tables for all functions in all modules in a workbook
Call complexRoutineStart("")
    Dim theWKBK As Workbook
    Application.EnableEvents = True ' stop macros triggering when the workbook opens
    Set theWKBK = Workbooks.Open(tkbk)
    Dim aFPath As String: aFPath = theWKBK.Path & "\Mods"
    Dim aModVDOB As ModuleVersionDataObject
    Set aModVDOB = createModuleHeaderObjectFromWKBK(theWKBK, aFPath)
    theWKBK.Close
    Application.EnableEvents = False ' stop macros triggering when the workbook opens
    Call aModVDOB.updateAllTables
Call complexRoutineEnd("")
End Sub

Public Sub updateLogsForWorkbook(tkbk As String, addLog As Boolean)
' update the dubug log for a single workbook
Call complexRoutineStart("")
    Dim theWKBK As Workbook
    Application.EnableEvents = False ' stop macros triggering when the workbook opens
    Set theWKBK = Workbooks.Open(tkbk)
    Dim aFPath As String: aFPath = theWKBK.Path & "\Mods"
    Dim aModVDOB As ModuleVersionDataObject
    Set aModVDOB = createModuleHeaderObjectFromWKBK(theWKBK, aFPath)
    theWKBK.Close
    Application.EnableEvents = True ' stop macros triggering when the workbook opens
    Call aModVDOB.updateLogFunctions(addLog)
Call complexRoutineEnd("")
End Sub

Public Sub removeGapsFromWorkbook(tkbk As String)
' removes gaps from within all functions in workbook
Call complexRoutineStart("")
    Dim theWKBK As Workbook
    Application.EnableEvents = False ' stop macros triggering when the workbook opens
    
    Set theWKBK = Workbooks.Open(tkbk)
    Dim aFPath As String: aFPath = theWKBK.Path & "\Mods"
    Dim aModVDOB As ModuleVersionDataObject
    Set aModVDOB = createModuleHeaderObjectFromWKBK(theWKBK, aFPath)
    theWKBK.Close
    
    Application.EnableEvents = True ' reenable macros triggering when the workbook opens
    
    Call aModVDOB.removeGapsFromfunctions
Call complexRoutineEnd("")
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
