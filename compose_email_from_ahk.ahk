::;minutes:: 
	;variable are set a separate file
	#Include my_variables.ahk

	Clipboard := "" ;Empty the clipboard
	Send ^a
	Send ^c
	ClipWait, 1
	Send {Esc}
	meeting_name := Clipboard
	FormatTime, CurrentDateTime,, yyyy-MM-dd
	
	try
		outlookApp := ComObjActive("Outlook.Application")
	catch
		outlookApp := ComObjCreate("Outlook.Application")
	MailItem := outlookApp.CreateItem(0)
	MailItem.BodyFormat := 2 ; olFormatHTML
	MailItem.Display
	MailItem.Subject := meeting_name "minutes of " CurrentDateTime
	;MailItem.To := ""
	MailItem.HTMLBody := "Hello all, <br> <br> Attached you can find our last, " meeting_name " minutes. Please, do not hesitate to reach me out in case of any questions. <br> <br> Best Regards,<br> Karolis" 
    Run, %meeting_minutes_folder_path%
	return


;^r:: Reload ;reload for testing new function. Comment out when testing is done.







