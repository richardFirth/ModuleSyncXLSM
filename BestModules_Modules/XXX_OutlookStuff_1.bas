Attribute VB_Name = "XXX_OutlookStuff_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.1
'$*DATE*30Jan18
'$*ID*OutlookStuff



Option Explicit

' use microsoft outlook 16.0 object library



Sub getAttachmentsFromFiles(theFilePaths() As String, toSavePath As String)

    Dim objOL As Outlook.Application
    'Dim Msg As Outlook.MailItem
    Dim msg As Object
    
    Dim att As Outlook.Attachment

    Set objOL = CreateObject("Outlook.Application")
    
    Dim x As Integer
    
    For x = LBound(theFilePaths) To UBound(theFilePaths)
        If Right(theFilePaths(x), 4) <> ".msg" Then GoTo notMSG
        On Error GoTo problemOpening
            Set msg = objOL.Session.OpenSharedItem(theFilePaths(x))
        On Error GoTo 0
        msg.Display
        For Each att In msg.Attachments
            att.SaveAsFile toSavePath & "\" & att.fileName
        Next att
        
          msg.Close (olDiscard)
        Set msg = Nothing
        
problemOpening:
notMSG:
    Next x

    Set objOL = Nothing
    
    
End Sub


Function TestOutlookIsOpen() As Boolean
    Dim testOutlook As Object

    On Error Resume Next
    Set testOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If testOutlook Is Nothing Then
       ' MsgBox "Outlook is not open, open Outlook and try again"
       TestOutlookIsOpen = False
    Else
       ' MsgBox
       TestOutlookIsOpen = True
    End If
    
    Set testOutlook = Nothing
    
End Function
