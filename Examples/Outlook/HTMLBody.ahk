; This script creates an email, adds "test.png" as an attachment, then includes the image in the body of the email.

Image := A_ScriptDir "\test.png"  ; The path of the image to include.
SplitPath, Image, ImageName  ; Get the file name from the path.

; Constants
olMailItem := 0
olByValue := 1

olApp := ComObjCreate("Outlook.Application")  ; Create an application object.
olMail := olApp.CreateItem(olMailItem)  ; Create a new email.
olMail.To := "abc@example.com"
olMail.CC := "xyz@example.com"
olMail.Subject := "foo"
olMail.Body := "abc123`n`n"
olMail.Attachments.Add(Image, olByValue, 0)  ; Add an attachment.
olMail.HTMLBody := olMail.HTMLBody "<br><B>Embedded Image:</B><br>"  ; Include the image in the body of the email.
                . "<img src='cid:" ImageName "' width='500' height='400'><br>"
                . "<br>Best Regards, <br>Someone</font></span>"
olMail.Display
;~ olMail.Send
