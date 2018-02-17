Attribute VB_Name = "YYY_VBAModules_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.3
'$*DATE*16Feb18
'$*ID*VBAModules

Option Explicit

' you need to add tools/references - microsoft visual basic for applications extensibility 5.3
' use this function to dump modules into existing folder. from here you can commit to github
' and see the changes in the modules code from there.

Sub saveAllModules()
    Dim modulePath As String
    modulePath = ThisWorkbook.Path & "\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "_Modules"
    Call saveAllModulesFromWKBKToFolder(ThisWorkbook, modulePath)
End Sub


Sub saveAllModulesFromWKBKToFolder(aWKBK As Workbook, fPath As String)
        Dim VBProj As VBIDE.VBProject: Set VBProj = aWKBK.VBProject
        Dim VBComp As VBIDE.VBComponent
                    
        For Each VBComp In VBProj.VBComponents
            Call createDirectoryRF(fPath)
            VBComp.Export fPath & "\" & VBComp.Name & getModuleExtention(VBComp)
        Next VBComp
End Sub

Public Function getModuleExtention(VBComp As VBIDE.VBComponent) As String
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
