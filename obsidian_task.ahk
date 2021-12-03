
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

;variable are set a separate file
#Include my_variables.ahk

    ::;obst:: 
        Clipboard := "" ;Empty the clipboard
        Send ^a
        Send ^c
        ClipWait, 1
        Send {Esc}
        task_name := Clipboard
        obsidian_task_title := task_name
        Sleep 100
        ;SetTitleMatchMode, 2
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
        
            obsidian_task_title := task_name ">>>" olMailItem.SenderName " --> " olMailItem.To " - " olMailItem.ReceivedTime " - " olMailItem.Subject 
        }
        file := FileOpen(obsidian_daily_note_file_name,"a")    
        file.write("`n- [ ] #inbox " obsidian_task_title)
        file.close()
        ;MsgBox %obsidian_daily_note_file_name% 
        return