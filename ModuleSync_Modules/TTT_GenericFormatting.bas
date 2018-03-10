Attribute VB_Name = "TTT_GenericFormatting"

'/T--TTT_GenericFormatting-----------------------------------------------------------------------\
' Function Name            | Return   |  Description                                             |
'--------------------------|----------|----------------------------------------------------------|
'getReturnType             | String   |  gets the return type for a function                     |
'getFunctionName           | String   |  gets the name of a function or sub ' used many places!  |
'checkForSubOrFunction     | Boolean  | checks if a line item is a sub or function               |
'checkForEndSubOrFunction  | Boolean  | checks for the end of a sub or function                  |
'\-----------------------------------------------------------------------------------------------/

Option Explicit

Public Function getReturnType(tFDec As String) As String
' gets the return type for a function
    If InStr(1, tFDec, "Function", vbBinaryCompare) > 0 Then
        Dim strG() As String
        strG = Split(tFDec, " ")
        getReturnType = strG(UBound(strG))
    Else
        getReturnType = "Void"
    End If
End Function

Public Function getFunctionName(tFDec As String) As String
' gets the name of a function or sub ' used many places!
    Dim SPL1() As String
    SPL1 = Split(tFDec, "(")
        
    Dim strG() As String
    strG = Split(SPL1(0), " ")
    getFunctionName = strG(UBound(strG))
   
End Function

Public Function checkForSubOrFunction(tSTR As String) As Boolean
'checks if a line item is a sub or function
    If Left(Trim(tSTR), 3) = "Sub" Then checkForSubOrFunction = True: Exit Function
    If Left(Trim(tSTR), 11) = "Public Sub " Then checkForSubOrFunction = True: Exit Function
    If Left(Trim(tSTR), 12) = "Private Sub " Then checkForSubOrFunction = True: Exit Function
    If Left(Trim(tSTR), 8) = "Function" Then checkForSubOrFunction = True: Exit Function
    If Left(Trim(tSTR), 16) = "Public Function " Then checkForSubOrFunction = True: Exit Function
    If Left(Trim(tSTR), 17) = "Private Function " Then checkForSubOrFunction = True: Exit Function
    
End Function

Public Function checkForEndSubOrFunction(tSTR As String) As Boolean
'checks for the end of a sub or function
    If Left(Trim(tSTR), 7) = "End Sub" Then checkForEndSubOrFunction = True: Exit Function
    If Left(Trim(tSTR), 12) = "End Function" Then checkForEndSubOrFunction = True: Exit Function
End Function
