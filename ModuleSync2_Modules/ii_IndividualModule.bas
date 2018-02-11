Attribute VB_Name = "ii_IndividualModule"
Option Explicit




Public Function selectModuleNameFromWKBK(theWKBK As Workbook) As VBIDE.VBComponent

    Dim theSelector As XXX_UISelect
    Set theSelector = New XXX_UISelect

    Dim modCol As Collection
    Set modCol = getAllModules(theWKBK)

        Call theSelector.setOptionsToSelect(ModuleCollection2String(modCol))
    
    theSelector.Show
    
    Set selectModuleNameFromWKBK = getModuleByName(modCol, theSelector.getSelectedOption())
    
    Unload theSelector

End Function




Public Sub viewAllFunctionsInModuleComponent(theModule As VBIDE.VBComponent)

    Dim modulePath As String
    modulePath = ExportVBAModulePath(selectedModule)
    If modulePath = "BAD" Then
        MsgBox "BAD Module Path": Exit Sub
    End If
    
    Dim VBACode() As String: VBACode = getTxTDocumentAsString(modulePath)
    
    Kill modulePath
    
    Dim listofF() As String
    listofF = identifySubsAndFunctions(VBACode)
    
    
    Dim newWKBK As Workbook
    Set newWKBK = Workbooks.Add
    
    Call printStringArrToColumn(listofF, newWKBK.Sheets(1), 1, "The Functions")
    
End Sub






