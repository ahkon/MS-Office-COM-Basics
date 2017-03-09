; This script resolves an exchange user's display name to a SMTP address.
; An email is open in an Inspector window for this example.

olApp := ComObjActive("Outlook.Application")
Window := olApp.ActiveWindow 
if (Window.Class = 35)  ; 35 = An Inspector object. (The email is open in its own window, as opposed to the main window.)
{
    SmtpAddress := ""
    MailItem := Window.CurrentItem
    FromName := MailItem.SenderName
    Recip := olApp.Session.CreateRecipient(FromName)
    Recip.Resolve
    if (Recip.Resolved)
    {
        UserType := Recip.AddressEntry.AddressEntryUserType
        if (UserType = 0 || UserType = 10)  ; olExchangeUserAddressEntry || olOutlookContactAddressEntry
            SmtpAddress := Recip.AddressEntry.GetExchangeUser.PrimarySmtpAddress
        else if (UserType = 1)  ; olExchangeDistributionListAddressEntry
            SmtpAddress := Recip.AddressEntry.GetExchangeDistributionList.PrimarySmtpAddress
    }
    if (SmtpAddress = "")
        SmtpAddress := MailItem.SenderEmailAddress
}
MsgBox, % "SmtpAddress: " SmtpAddress
return

; References
;   - http://answers.microsoft.com/en-us/office/forum/office_2007-customize/need-vba-to-obtain-smtp-address-of-exchange-user/97833c7c-18e3-4b8d-923a-606b81c9ecd1?auth=1
;   - https://autohotkey.com/board/topic/71335-mickers-outlook-com-msdn-for-ahk-l/page-3#entry731838
