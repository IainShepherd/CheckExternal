Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'Developer: Iain Shepherd
'Version: 1.1.1
'Last Updated: 10/06/2019

'Set your interal email address domain here:
Dim HomeDomain As String
HomeDomain = "@nationalgrideso.com"

Cancel = False
Dim xMailItem As Outlook.MailItem
Dim xRecipitents As Outlook.Recipient
Dim i As Long
Dim xRecipientAddress As String
Dim ExternalDomain As Boolean
ExternalDomain = False
Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
Dim DomainList As String
If Item.Class <> olMail Then Exit Sub

Set xMailItem = Item
'Set xRecipitents = xMailItem.Recipients

'GetSMTPAddressForRecipients xMailItem

For Each recip In xMailItem.Recipients
    Set pa = recip.PropertyAccessor
    'check the domain
    If InStr(1, LCase(pa.GetProperty(PR_SMTP_ADDRESS)), LCase(HomeDomain)) = 0 Then
        'external domain
        ExternalDomain = True
        
        'List the domains
        'DomainList = DomainList & Chr(10) & Mid(pa.GetProperty(PR_SMTP_ADDRESS), InStr(1, pa.GetProperty(PR_SMTP_ADDRESS), "@"))
        'List the emails
        DomainList = DomainList & Chr(10) & pa.GetProperty(PR_SMTP_ADDRESS)
        
    End If
    'Debug.Print pa.GetProperty(PR_SMTP_ADDRESS)
Next

If ExternalDomain = True Then
    sendcheck = MsgBox("WARNING: One or more external emails" & Chr(10) & "Do you wish to send this email externally?" & Chr(10) & Chr(10) _
                        & "The external emails are: " & Chr(10) & DomainList, vbYesNoCancel + vbExclamation)
    If sendcheck = vbNo Or sendcheck = vbCancel Then
        Cancel = True
    End If
End If

End Sub
