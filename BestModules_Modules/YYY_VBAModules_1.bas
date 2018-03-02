Attribute VB_Name = "YYY_VBAModules_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.8
'$*DATE*2/28/2018*xx
'$*ID*VBAModules
'$*CharCount*5955*xxxx
'$*RowCount*153*xxxx

'/T--YYY_VBAModules_1-------------------------------------------------------------------------------------------------------------------\
' Function Name                  | Return    |  Description                                                                             |
'--------------------------------|-----------|------------------------------------------------------------------------------------------|
'saveAllModules                  | Void      |  saves all the modules to a folder called Thisworkbook.name_Modules                      |
'saveAllModulesFromWKBKToFolder  | Void      |  saves all modules in a wkbk to a folder.                                                |
'getModuleExtention              | String    |  gets the extention for a module                                                         |
'ImportModulesToWKBK             | Boolean   | imports many modules to a workbook, given paths                                          |
'ImportModuleToWKBK              | Boolean   |  imports one module to a workbook                                                        |
'ExportVBAModulesToPaths         | String()  | exports all modules in a collection to a given file path. returns paths of module files  |
'ExportVBAModuleToPath           | String    | exports one module to a given file path. returns path of module file                     |
'RemoveModuleFromWKBKByName      | Boolean   | removes a module from a workbook given the name of the module                            |
'RemoveModuleFromWKBK            | Boolean   | removes a module object from a workbook given the object                                 |
'\--------------------------------------------------------------------------------------------------------------------------------------/

Option Explicit

' you need to add tools/references - microsoft visual basic for applications extensibility 5.3
' use this function to dump modules into existing folder. from here you can commit to github
' and see the changes in the modules code from there.

Private Sub saveAllModules()
' saves all the modules to a folder called Thisworkbook.name_Modules
Dim modulePath As String
modulePath = ThisWorkbook.Path & "\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "_Modules"
Call saveAllModulesFromWKBKToFolder(ThisWorkbook, modulePath)
End Sub

Private Sub saveAllModulesFromWKBKToFolder(aWKBK As Workbook, fPath As String)
' saves all modules in a wkbk to a folder.
Dim VBProj As VBIDE.VBProject: Set VBProj = aWKBK.VBProject
Dim VBComp As VBIDE.VBComponent
For Each VBComp In VBProj.VBComponents
    Call createDirectoryRF(fPath)
    VBComp.Export fPath & "\" & VBComp.Name & getModuleExtention(VBComp)
Next VBComp
End Sub

Public Function getModuleExtention(VBComp As VBIDE.VBComponent) As String
' gets the extention for a module
Select Case VBComp.Type
Case vbext_ct_ClassModule
getModuleExtention = ".cls"
Case vbext_ct_MSForm
getModuleExtention = ".frm"
Case vbext_ct_StdModule
getModuleExtention = ".bas"
Case vbext_ct_Document
getModuleExtention = ".cls"
End Select
End Function

Public Function ImportModulesToWKBK(theWKBK As Workbook, theModulePaths() As String) As Boolean
'imports many modules to a workbook, given paths
Dim x As Integer
For x = LBound(theModulePaths) To UBound(theModulePaths)
    Call ImportModuleToWKBK(theWKBK, theModulePaths(x))
Next x
ImportModulesToWKBK = True
End Function

Public Function ImportModuleToWKBK(theWKBK As Workbook, theModulePath As String) As Boolean
' imports one module to a workbook
On Error GoTo addmoduleprob

Call theWKBK.VBProject.VBComponents.Import(theModulePath)

ImportModuleToWKBK = True
Exit Function
addmoduleprob:
ImportModuleToWKBK = False
End Function

Public Function ExportVBAModulesToPaths(theModules As Collection, thePath As String) As String()
'exports all modules in a collection to a given file path. returns paths of module files
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
'exports one module to a given file path. returns path of module file
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
'removes a module from a workbook given the name of the module
Dim theMod As VBIDE.VBComponent
Set theMod = getModuleByNameFromWKBK(theWKBK, theModule)
If theMod Is Nothing Then
    Dim info(1 To 2) As String
    info(1) = theWKBK.Name
    info(2) = theModule
    Call reportError("RemoveModuleFromWKBKByName", info)
    RemoveModuleFromWKBKByName = False: Exit Function
End If
RemoveModuleFromWKBKByName = RemoveModuleFromWKBK(theWKBK, theMod)

End Function

Public Function RemoveModuleFromWKBK(theWKBK As Workbook, theModule As VBIDE.VBComponent) As Boolean
'removes a module object from a workbook given the object
On Error GoTo RemoveModuleFromWKBKERR
Call theWKBK.VBProject.VBComponents.Remove(theModule)

RemoveModuleFromWKBK = True

Exit Function

RemoveModuleFromWKBKERR:
RemoveModuleFromWKBK = False
End Function

