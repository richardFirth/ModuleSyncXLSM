Attribute VB_Name = "YYY_VBAModules_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.0
'$*DATE*9Feb18
'$*ID*VBAModules






Option Explicit





' use this function to dump modules into existing folder. from here you can commit to github
' and see the changes in the modules code from there.


Sub saveAllModules()
    Call saveAllModulesFromWKBKToFolder(ThisWorkbook, ThisWorkbook.Path)
End Sub


Sub saveAllModulesFromWKBKToFolder(aWKBK As Workbook, fPath As String)


        Dim VBProj As VBIDE.VBProject: Set VBProj = aWKBK.VBProject
        Dim VBComp As VBIDE.VBComponent
                    

        For Each VBComp In VBProj.VBComponents
            
            VBComp.Export fPath & "\" & VBComp.name & getModuleExtention(VBComp)
             
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
