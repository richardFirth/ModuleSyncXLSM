Attribute VB_Name = "JJJ_ModuleCollections"





Option Explicit


Function getAllModules(theWKBK As Workbook) As Collection
    Dim loc As Collection
    Set loc = getModulesByType(theWKBK, A_NORMAL)
    Set loc = combineCollections(loc, getModulesByType(theWKBK, B_CLASS))
    Set loc = combineCollections(loc, getModulesByType(theWKBK, C_FORM))

    Set getAllModules = loc

End Function



Function getModulesByType(theWKBK As Workbook, ms As ModuleSpecies) As Collection

        Dim VBProj As VBIDE.VBProject: Set VBProj = theWKBK.VBProject
        Dim VBComp As VBIDE.VBComponent
                      
        Dim n As Integer: n = 1
        Dim locCollect As New Collection

        For Each VBComp In VBProj.VBComponents
            
            If ms = A_NORMAL And VBComp.Type = vbext_ct_StdModule Then locCollect.Add VBComp
            If ms = B_CLASS And VBComp.Type = vbext_ct_ClassModule Then locCollect.Add VBComp
            If ms = C_FORM And VBComp.Type = vbext_ct_MSForm Then locCollect.Add VBComp
             
        Next VBComp
        
        Set getModulesByType = locCollect
        
End Function



Function getModuleByName(theCol As Collection, moduleName As String) As VBIDE.VBComponent

    Dim locMod As VBIDE.VBComponent

    If moduleName = "" Then Exit Function

    For Each locMod In theCol
       If moduleName = locMod.Name Then Set getModuleByName = locMod
    Next locMod
    
End Function



Function getModuleByNameFromWKBK(theWKBK As Workbook, moduleName As String) As VBIDE.VBComponent
    Dim modCol As Collection
    Set modCol = getAllModules(theWKBK)
    
    Set getModuleByNameFromWKBK = getModuleByName(modCol, moduleName)
    
End Function






Function combineCollections(col1 As Collection, col2 As Collection) As Collection

    Dim i As Integer
    Dim superCol As New Collection

    For i = 1 To col1.Count
        superCol.Add col1.Item(i)
    Next i
    
    For i = 1 To col2.Count
        superCol.Add col2.Item(i)
    Next i
    Set combineCollections = superCol
    
End Function




Function ModuleCollection2String(theCollection As Collection) As String()
    
    Dim VBComp As VBIDE.VBComponent
    Dim locSTR() As String
    Dim n As String: n = 1
    
    For Each VBComp In theCollection
       ReDim Preserve locSTR(1 To n) As String
       locSTR(n) = VBComp.Name
       n = n + 1
    Next VBComp
    ModuleCollection2String = locSTR
End Function


