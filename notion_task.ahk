;script optimizations
    #NoEnv
    #KeyHistory 0
    #SingleInstance force		; Cannot have multiple instances of program
    #MaxHotkeysPerInterval 200	; Won't crash if button held down
    #Persistent
    ListLines Off
    Process, Priority, , A
    SetBatchLines, -1
    SetKeyDelay, -1, -1
    SetMouseDelay, -1
    SetDefaultMouseSpeed, 0
    SetWinDelay, -1
    SetControlDelay, -1
    SendMode Input

;variable are set a separate file
#Include my_variables.ahk

;define script icon
    Menu, tray, icon, %A_ScriptDir%/icons/notion.png  ;Setting tray icon

;open notion if not already running
	Process, Exist, Notion.exe ; check to see if notion is running
		{
		If (ErrorLevel = 0) ;notion is not running
			Run %notion_exe_dir%
		else
            MsgBox, "Notion is already running"
		}
    return

;focus task
    ::;find::
        send {Esc}
        Send ^c
        ClipWait, 1
        If WinExist("ahk_class rctrl_renwnd32 ahk_exe OUTLOOK.EXE")
            WinActivate, ahk_class rctrl_renwnd32 ahk_exe OUTLOOK.EXE
        Send ^e
        Send ^a
        Send ^v
        Send {Enter}
        return
        


;reminder task home
    ::;remindh:: 
        Clipboard := "" ;Empty the clipboard
        Send ^a
        Send ^c
        ClipWait, 1
        Send {Esc}
        task_name := Clipboard
        notion_task_title := task_name
        Sleep 100
        SetTitleMatchMode, 2
        ; if WinActive("kirna.karolis@gmail.com")  ;if chrome page with title including kirna.karolis@gmail.com is actvive (targeting gmail) 
        if WinActive("@gmail.com - Gmail - Google Chrome")  ;if chrome page with title including kirna.karolis@gmail.com is actvive (targeting gmail) 
        {
            Clipboard := ""
            Send, {F6}
            Send ^c
            ClipWait, 1
            gmail_email_url := Clipboard
            notion_task_title := task_name ">>>" gmail_email_url
        }
        whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        whr.Open("POST", "https://api.notion.com/v1/pages/")
        whr.SetRequestHeader("Authorization", bearer_token)
        whr.SetRequestHeader("Content-Type", "application/json")
        whr.SetRequestHeader("Notion-Version", notion_api_version)
        body = {"parent":{"database_id":"%work_task_database_id%"},"properties":{"Name":{"title":[{"text":{"content":"%Clipboard%"}}]},"Tags":{"relation":[{"id":"%tags_relation_id%"}]},"Wco~":{"date":{"start":"%A_YYYY%-%A_MM%-%A_DD%","end":null}},"<aw{":{"multi_select":[{"name":"Urgent"},{"name":"Important"}]},"OlWY":{"multi_select":[{"name":"Need To Remind"}]}}}
        whr.Send(body)
        MsgBox, % notion_task_title
        ; Msgbox, % whr.ResponseText
        return  

;task home   
    ::;taskh:: 
        Clipboard := "" ;Empty the clipboard
        Send ^a
        Send ^c
        ClipWait, 1
        Send {Esc}
        task_name := Clipboard
        notion_task_title := task_name
        Sleep 100
        SetTitleMatchMode, 2
        if WinActive("@gmail.com - Gmail - Google Chrome")  ;if chrome page with title including kirna.karolis@gmail.com is actvive (targeting gmail) 
        {
            Clipboard := ""
            Send, {F6}
            Send ^c
            ClipWait, 1
            gmail_email_url := Clipboard
            notion_task_title := task_name ">>>" gmail_email_url
        }
        whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        whr.Open("POST", "https://api.notion.com/v1/pages/")
        whr.SetRequestHeader("Authorization", bearer_token)
        whr.SetRequestHeader("Content-Type", "application/json")
        whr.SetRequestHeader("Notion-Version", notion_api_version)
        body = {"parent":{"database_id":"%work_task_database_id%"},"properties":{"Responsible":{"relation":[{"id":"%responsible_relation_id%"}]},"Tags":{"relation":[{"id":"%tags_relation_id%"}]},"Name":{"title":[{"text":{"content":"%notion_task_title%"}}]}}}
        whr.Send(body)
        MsgBox, % notion_task_title
        ; Msgbox, % whr.ResponseText
        return

;task work
    ::;taskw:: 
    ; StartTime := A_TickCount
        Clipboard := "" ;Empty the clipboard
        Send ^a
        Send ^c
        ClipWait, 1
        Send {Esc}
        task_name := Clipboard
        notion_task_title := task_name
        Sleep 100
        if WinActive("ahk_class rctrl_renwnd32 ahk_exe OUTLOOK.EXE") ;if e-mail in outlook task is selected
        {	
            olApp := ComObjActive("Outlook.Application")
            try
            {
                olItem := olApp.ActiveWindow.CurrentItem
            }
            catch
            {
                olItem := olApp.ActiveExplorer.Selection.Item(1)
            }
            if (olItem.Class = 43)
            {
                olMailItem := olItem
            }
            else
            {
                MsgBox Mail Item Not Current or Selected
                return
            }
            notion_task_title := task_name ">>>" olMailItem.SenderName " --> " olMailItem.To " - " olMailItem.ReceivedTime " - " olMailItem.Subject 
        } 
        whr := ComObjCreate("WinHttp.WinHttpRequest.5.1") 
        whr.Open("POST", "https://api.notion.com/v1/pages/") 
        whr.SetRequestHeader("Authorization", bearer_token) 
        whr.SetRequestHeader("Content-Type", "application/json")
        whr.SetRequestHeader("Notion-Version", notion_api_version)
        body = {"parent":{"database_id":"%work_task_database_id%"},"properties":{"Responsible":{"relation":[{"id":"%responsible_relation_id%"}]},"Name":{"title":[{"text":{"content":"%notion_task_title%"}}]}}}
        whr.Send(body)
        MsgBox, % notion_task_title
        ;Msgbox, % whr.ResponseText
    ; ElapsedTime := A_TickCount - StartTime
    ; MsgBox,  %ElapsedTime% milliseconds have elapsed.
        return

;reminder work
    ::;remindw:: 
        Clipboard := "" ;Empty the clipboard
        Send ^a
        Send ^c
        ClipWait, 1
        Send {Esc}
        task_name := Clipboard
        notion_task_title := task_name
        Sleep 100
        if WinActive("ahk_class rctrl_renwnd32 ahk_exe OUTLOOK.EXE") ;if e-mail in outlook task is selected
        {	
            olApp := ComObjActive("Outlook.Application")
            try
            {
                olItem := olApp.ActiveWindow.CurrentItem
            }
            catch
            {
                olItem := olApp.ActiveExplorer.Selection.Item(1)
            }
            if (olItem.Class = 43)
            {
                olMailItem := olItem
            }
            else
            {
                MsgBox Mail Item Not Current or Selected
                return
            }
            notion_task_title := task_name ">>>" olMailItem.SenderName " --> " olMailItem.To " - " olMailItem.ReceivedTime " - " olMailItem.Subject 
        } 
        whr := ComObjCreate("WinHttp.WinHttpRequest.5.1") 
        whr.Open("POST", "https://api.notion.com/v1/pages/") 
        whr.SetRequestHeader("Authorization", bearer_token) 
        whr.SetRequestHeader("Content-Type", "application/json")
        whr.SetRequestHeader("Notion-Version", notion_api_version)
        body = {"parent":{"database_id":"%work_task_database_id%"},"properties":{"Name":{"title":[{"text":{"content":"%notion_task_title%"}}]},"Wco~":{"date":{"start":"%A_YYYY%-%A_MM%-%A_DD%","end":null}},"<aw{":{"multi_select":[{"name":"Urgent"},{"name":"Important"}]},"OlWY":{"multi_select":[{"name":"Need To Remind"}]}}}
        whr.Send(body)
        MsgBox, % notion_task_title
        ; Msgbox, % whr.ResponseText
        return

;copy email data
    ::;copy:: ;COPY OUTLOOK DATA
        Send {Esc}
        Clipboard := "" ; Empty the clipboard
        sleep 300
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
            Clipboard := Clipboard ">>>" olMailItem.SenderName " --> " olMailItem.To " - " olMailItem.ReceivedTime " - " olMailItem.Subject 
            
            ClipWait, 1
            if ErrorLevel
            {
                MsgBox, The attempt to copy text onto the clipboard failed.
                return
            }
            MsgBox, % Clipboard 
            ;Msgbox, % whr.ResponseText 
            return 
            } 
            ;COPY EVERYWHERE ELSE 
            else { 
            MsgBox, "Outlook item is not selected"
            ;Msgbox, % whr.ResponseText
            return
        }

DllCall("Sleep",UInt,17) ;?????????I just used the precise sleep function to wait exactly 17 milliseconds

;edit the script
::;edit::
    Send {Esc}
    Run, %code_exe_path% "%A_ScriptDir%"  
    return

;restart the script
::;restart::
    Send {Esc}
    Reload
    Sleep 1000 ; If successful, the reload will close this instance during the Sleep, so the line below will never be reached.
    MsgBox, 4, The script could not be reloaded. Would you like to open it for editing?
    IfMsgBox, Yes, Edit
    return