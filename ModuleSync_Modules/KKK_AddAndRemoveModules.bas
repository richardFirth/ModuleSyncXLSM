Attribute VB_Name = "KKK_AddAndRemoveModules"
Option Explicit


'/---KKK_AddAndRemoveModules-------------------------------------------------------------------------------------------------------------------\
'  Function Name             | Return             |   Description                                                                              |
'----------------------------|--------------------|--------------------------------------------------------------------------------------------|
' ImportModulesToWKBK        | Boolean            | imports many modules to a workbook, given paths                                            |
' ImportModuleToWKBK         | Boolean            | imports one module to a workbook                                                           |
' ExportVBAModulesToPaths    | String()           | exports all modules in a collection to a given file path. returns paths of module files    |
' ExportVBAModuleToPath      | String             | exports one module to a given file path. returns path of module file                       |
' RemoveModuleFromWKBKByName | Boolean            | removes a module from a workbook given the name of the module                              |
' RemoveModuleFromWKBK       | Boolean            | removes a module object from a workbook given the object                                   |
'\---------------------------------------------------------------------------------------------------------------------------------------------/


Public Function ImportModulesToWKBK(theWKBK As Workbook, theModulePaths() As String) As Boolean
    Dim x As Integer
    For x = LBound(theModulePaths) To UBound(theModulePaths)
        Call ImportModuleToWKBK(theWKBK, theModulePaths(x))
    Next x
    ImportModulesToWKBK = True
End Function


Public Function ImportModuleToWKBK(theWKBK As Workbook, theModulePath As String) As Boolean
    On Error GoTo addmoduleprob
    
    Call theWKBK.VBProject.VBComponents.Import(theModulePath)
   
    ImportModuleToWKBK = True
Exit Function
addmoduleprob:
    ImportModuleToWKBK = False
End Function





Public Function ExportVBAModulesToPaths(theModules As Collection, thePath As String) As String()
        Dim locExportPaths() As String
        Dim aModule As VBIDE.VBComponent
        Dim n As Integer: n = 1
        
    On Error GoTo exportModulesProblem
    
    For Each aModule In theModules
        ReDim Preserve locExportPaths(1 To n) As String
        locExportPaths(n) = thePath & "\" & aModule.Name & getModuleExtention(aModule)
        aModule.Export locExportPaths(n)
        n = n + 1
    Next aModule

ExportVBAModulesToPaths = locExportPaths

Exit Function
        
exportModulesProblem:
    Dim errDat(1 To 4) As String
    errDat(1) = thePath
    errDat(2) = n
    errDat(3) = aModule.Name & getModuleExtention(aModule)
    errDat(4) = locExportPaths(n)
    Call reportError("ExportVBAModulesToPaths", errDat)
    Resume Next
End Function


Public Function ExportVBAModuleToPath(theModule As VBIDE.VBComponent, thePath As String) As String
        Dim locExportPath As String
    On Error GoTo exportProblem
        locExportPath = thePath & "\" & theModule.Name & getModuleExtention(theModule)
        theModule.Export locExportPath
        ExportVBAModuleToPath = locExportPath

Exit Function
        
exportProblem:
    ExportVBAModuleToPath = "BAD"
   
End Function





Public Function RemoveModuleFromWKBKByName(theWKBK As Workbook, theModule As String) As Boolean

    Dim theMod As VBIDE.VBComponent
    Set theMod = getModuleByNameFromWKBK(theWKBK, theModule)
        If theMod Is Nothing Then
            RemoveModuleFromWKBKByName = False: Exit Function
        End If
    RemoveModuleFromWKBKByName = RemoveModuleFromWKBK(theWKBK, theMod)

End Function




Public Function RemoveModuleFromWKBK(theWKBK As Workbook, theModule As VBIDE.VBComponent) As Boolean
 
 On Error GoTo RemoveModuleFromWKBKERR
   Call theWKBK.VBProject.VBComponents.Remove(theModule)
   
RemoveModuleFromWKBK = True

Exit Function

RemoveModuleFromWKBKERR:
    RemoveModuleFromWKBK = False
End Function



