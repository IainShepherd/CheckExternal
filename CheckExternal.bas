Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'Developer: Iain Shepherd
'Version: 1.0.1
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
        Exit For
    End If
    'Debug.Print pa.GetProperty(PR_SMTP_ADDRESS)
Next

If ExternalDomain = True Then
    sendcheck = MsgBox("WARNING: One or more external emails" & Chr(10) & "Do you wish to send this email externally?", vbYesNoCancel + vbExclamation)
    If sendcheck = vbNo Or sendcheck = vbCancel Then
        Cancel = True
    End If
End If

End Sub
