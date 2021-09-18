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

;caps-keys
    ; Terminate window, shut down, restart script, delete key
    CapsLock & q:: !F4
    ; Vimlike bindings
    CapsLock & h:: Left
    CapsLock & j:: Down
    CapsLock & k:: Up
    CapsLock & l:: Right
    CapsLock & a:: Send {end}
    CapsLock & x:: Send {delete}
    CapsLock & i:: Send {home}
    CapsLock & u:: Send {PGUP}
    CapsLock & d:: Send {PGDN}

    Capslock & w::
    Send {Ctrl Down}{Right}{Ctrl Up}
    If GetKeyState("Shift")
    Send {Ctrl Down}{Shift Down}{Right}{Ctrl Up}{Shift Up}
    Return

    Capslock & b:: Send {Ctrl Down}{Left}{Ctrl Up}


DllCall("Sleep",UInt,17) ;?????????I just used the precise sleep function to wait exactly 17 milliseconds