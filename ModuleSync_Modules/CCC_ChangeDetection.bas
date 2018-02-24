Attribute VB_Name = "CCC_ChangeDetection"
Option Explicit


Public Enum AlterVersion
    A_INCREASE
    B_DECREASE
    C_LEAVEALONE
End Enum


Public Function detectCodeChange(VBACODE() As String, tPath As String) As Boolean

    If getTrueCharCount(VBACODE) <> getWrittenCharCount(VBACODE) Then detectCodeChange = True: Exit Function


    If UCase(Right(tPath, 3)) = "FRM" Then
        If getTrueRowCount(VBACODE) - 1 = getWrittenRowCount(VBACODE) Then Exit Function
        If getTrueRowCount(VBACODE) = getWrittenRowCount(VBACODE) Then Exit Function
        
        detectCodeChange = True
        
        Dim tDat(1 To 4) As String
        tDat(1) = tPath
        tDat(2) = getTrueRowCount(VBACODE)
        tDat(3) = getWrittenRowCount(VBACODE)
        tDat(4) = detectCodeChange
        Call reportError("detectCodeChange", tDat)
    Else
        If getTrueRowCount(VBACODE) <> getWrittenRowCount(VBACODE) Then detectCodeChange = True
    End If
    
    
End Function



Function updateCountsInFile(tPath As String, altV As AlterVersion) As Boolean
    
    Dim VBACODE() As String: VBACODE = getTxTDocumentAsString(tPath)
    
    Dim TrueCharCount As Integer: TrueCharCount = getTrueCharCount(VBACODE)
    Dim TrueRowCount As Integer: TrueRowCount = getTrueRowCount(VBACODE)

    Dim x As Integer
    For x = LBound(VBACODE) To UBound(VBACODE)
    
        If Left(VBACODE(x), 6) = "'$*MIN" Then
            If altV = A_INCREASE Then VBACODE(x) = makeMinorVersion(VBACODE(x), True)
            If altV = B_DECREASE Then VBACODE(x) = makeMinorVersion(VBACODE(x), False)
        End If
    
        If Left(VBACODE(x), 6) = "'$*Cha" Then
            VBACODE(x) = makeCharCountSTR(TrueCharCount)
        End If
        
        If Left(VBACODE(x), 6) = "'$*Row" Then
            VBACODE(x) = makeRowCountSTR(TrueRowCount)
            Exit For ' this is the last thing touched
        End If
    Next x
    Call saveModuleToPath(VBACODE, tPath)
    updateCountsInFile = True
End Function



Sub saveModuleToPath(theMod() As String, thePath As String)
    Dim nCode() As String
    ' there's an extra newline character at end of createTextFromStringArr, and we want to make sure the tested values equal the written values
    If theMod(UBound(theMod)) = "" Then
            Dim y As Integer
            Dim m As Integer: m = 1
            For y = LBound(theMod) To UBound(theMod) - 1
                ReDim Preserve nCode(1 To m) As String
                nCode(m) = theMod(y)
                m = m + 1
            Next y
    Else
        nCode = theMod
    End If

    Call createTextFromStringArr(nCode, thePath)
End Sub




Function makeCharCountSTR(aCount As Integer) As String
    Dim locSTR As String
    locSTR = "'$*CharCount*" & aCount & "*"
    
    Dim x As Integer: Dim y As Integer: y = 22 - Len(locSTR)
    For x = 1 To y
        locSTR = locSTR & "x"
    Next x
    makeCharCountSTR = locSTR
End Function



Function makeRowCountSTR(aCount As Integer) As String
    Dim locSTR As String
    locSTR = "'$*RowCount*" & aCount & "*"
    
    Dim x As Integer: Dim y As Integer: y = 20 - Len(locSTR)
    For x = 1 To y
        locSTR = locSTR & "x"
    Next x
    makeRowCountSTR = locSTR
End Function



Function makeMinorVersion(tVer As String, addToMinor As Boolean) As String
    Dim loc() As String
    loc = Split(tVer, "*")
    
    If addToMinor Then
        makeMinorVersion = "'$*MINOR_VERSION*" & loc(2) + 0.1
    Else
        makeMinorVersion = "'$*MINOR_VERSION*" & loc(2) - 0.1
    End If
End Function



Function getTrueRowCount(theCode() As String) As Integer
    
'    Dim x As Integer
'    For x = LBound(theCode) To UBound(theCode)
'        If theCode(x) <> "" Then Exit For
'    Next x

    getTrueRowCount = UBound(theCode) ' - x
    
End Function

Function getWrittenRowCount(theCode() As String) As Integer
    Dim x As Integer
    
    For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 6) = "'$*Row" Then
            Dim getR() As String
            getR = Split(theCode(x), "*")
            getWrittenRowCount = getR(2): Exit Function
        End If
    Next x
    
    
End Function


Function getTrueCharCount(theCode() As String) As Integer
    Dim tCount As Integer
    Dim x As Integer
    For x = LBound(theCode) To UBound(theCode)
        tCount = tCount + Len(theCode(x))
    Next x
    
    getTrueCharCount = tCount
End Function

Function getWrittenCharCount(theCode() As String) As Integer
    Dim x As Integer
    
    For x = LBound(theCode) To UBound(theCode)
        If Left(theCode(x), 6) = "'$*Cha" Then
            Dim getR() As String
            getR = Split(theCode(x), "*")
            getWrittenCharCount = getR(2): Exit Function
        End If
    Next x
End Function

