;START SCRIPT OPTIMIZATIONS 
    #NoEnv
    #KeyHistory 0
    #SingleInstance force		; Cannot have multiple instances of program
    #MaxHotkeysPerInterval 200	; Won't crash if button held down
    #Persistent
   ; Menu, tray, icon, icons\spg.png
    ListLines Off
    Process, Priority, , A
    SetBatchLines, -1
    SetKeyDelay, -1, -1
    SetMouseDelay, -1
    SetDefaultMouseSpeed, 0
    SetWinDelay, -1
    SetControlDelay, -1
    SendMode Input
;END SCRIPT OPTIMIZATIONS 

;VARIABLES
    notion_version := "2021-05-13"
    bearer_token := ""
    database_id := ""
;START TEXT EXPANDING :*:here_goes_hotword::here_goes_desired_text_for_expansion
    :*:;example::example of text expansion worked
    :*:;today::
FormatTime, CurrentDateTime,, yyyy-MM-dd
SendInput %CurrentDateTime%
return
;END TEXT EXPANDING

;START TASK TO NOTION
    ::;task::
        Send ^a
        Send ^c
        Send {Esc}
        Sleep 300
        ;IF TASK IN OUTLOOK IS SELECTED
        if WinActive("ahk_class rctrl_renwnd32 ahk_exe OUTLOOK.EXE")
        {	
                olApp := ComObjActive("Outlook.Application")
            try
                olItem := olApp.ActiveWindow.CurrentItem
            catch
                olItem := olApp.ActiveExplorer.Selection.Item(1)
            if (olItem.Class = 43)
                olMailItem := olItem
            else
            {
                MsgBox Mail Item Not Current or Selected
                return
            }
            Sleep 300

            Clipboard := Clipboard ">>>" olMailItem.SenderName " --> " olMailItem.To " - " olMailItem.ReceivedTime " - " olMailItem.Subject
            whr := ComObjCreate("WinHttp.WinHttpRequest.5.1") 
            whr.Open("POST", "https://api.notion.com/v1/pages/") 
            whr.SetRequestHeader("Authorization", bearer_token) 
            whr.SetRequestHeader("Content-Type", "application/json")
            whr.SetRequestHeader("Notion-Version", notion_version)
            body = {"parent":{"database_id":"%database_id%"},"properties":{"Name":{"title":[{"text":{"content":"%Clipboard%"}}]}}}
            whr.Send(body) 
            MsgBox, % Clipboard 
            ;Msgbox, % whr.ResponseText 
            return 
            } 
            ;EVERYWHERE ELSE  
            else { Sleep 300 whr := ComObjCreate("WinHttp.WinHttpRequest.5.1") 
            whr.Open("POST", "https://api.notion.com/v1/pages/") 
            whr.SetRequestHeader("Authorization", bearer_token) 
            whr.SetRequestHeader("Content-Type", "application/json")
            whr.SetRequestHeader("Notion-Version", notion_version)
            body = {"parent":{"database_id":"%database_id%"},"properties":{"Name":{"title":[{"text":{"content":"%Clipboard%"}}]}}}
            whr.Send(body)
            MsgBox, % Clipboard
            ;Msgbox, % whr.ResponseText
            return
        }
;END TASK TO NOTION
DllCall("Sleep",UInt,17) ;I just used the precise sleep function to wait exactly 17 milliseconds