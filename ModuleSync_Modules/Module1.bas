Attribute VB_Name = "Module1"
Option Explicit



Public Sub removeGapsFromWorkbook(tkbk As String)
' Updates the tables for all functions in all modules in a workbook
Call complexRoutineStart("")

    Dim theWKBK As Workbook
    Set theWKBK = Workbooks.Open(tkbk)
    Dim aFPath As String: aFPath = theWKBK.Path & "\Mods"
    Dim aModVDOB As ModuleVersionDataObject
    Set aModVDOB = createModuleHeaderObjectFromWKBK(theWKBK, aFPath)
    theWKBK.Close
    Call aModVDOB.removeGapsFromfunctions

Call complexRoutineEnd("")

End Sub







Sub testtesttest()

    Dim sumMOD As New X_SingleModuleObject_1
    Call sumMOD.initializeModule("C:\Users\rfirth1\Desktop\Mods\TTT_GenericFormatting.bas")
    Call sumMOD.z_removeGapsInFunctions
    sumMOD.saveModule

End Sub
