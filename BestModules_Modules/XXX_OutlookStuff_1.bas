Attribute VB_Name = "XXX_OutlookStuff_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.3
'$*DATE*2/28/2018*xx
'$*ID*OutlookStuff
'$*CharCount*1761*xxxx
'$*RowCount*72*xxxxx

'/T--XXX_OutlookStuff_1-----------------------------------------------------------\
' Function Name           | Return   |  Description                               |
'-------------------------|----------|--------------------------------------------|
'getAttachmentsFromFiles  | Void     |  gets attachements from downloaded emails  |
'TestOutlookIsOpen        | Boolean  |  checks if outlook is open                 |
'\--------------------------------------------------------------------------------/

Option Explicit

' use microsoft outlook 16.0 object library

Sub getAttachmentsFromFiles(theFilePaths() As String, toSavePath As String)
' gets attachements from downloaded emails
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
' checks if outlook is open
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
