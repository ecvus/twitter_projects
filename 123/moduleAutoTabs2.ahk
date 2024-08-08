#MaxThreadsPerHotkey 0
#MaxHotkeysPerInterval 1000000000
#Persistent 

global switchButton := False
global imageFoundBySecondMethod := False
global Scroll := False
global NextPageGo := False

SetKeyDelay, 2, 2
SetMouseDelay, 2, 2

turndeOFFON(){
    switchButton := !switchButton
    if(switchButton = true)
    {
        TabWork()
    }
    Return
}

TabWork(){
    WinGetPos, X, Y, Width, Height, ahk_exe anty.exe
    sleep, 300
    Scroll := False
    imageFoundBySecondMethod := False
    ControlClick, x1449 y515, ahk_exe anty.exe ,,Left, 1
    NextPageGo := False
    startWork := A_TickCount
    tabsEmpty := False
    ControlSend,, j , ahk_exe anty.exe

    ; Переменные для хранения координат найденных изображений
    tabsFound1 := ""
    tabsFound2 := ""
    tabs1Found1 := ""
    tabs1Found2 := ""

    while(tabsEmpty = False && switchButton = true) {
        ; Поиск первого изображения
        WinActivate, ahk_exe anty.exe
        ImageSearch, tabsFound1, tabsFound2, 0, 0, Width, Height, *50 %A_ScriptDir%\img\Screenshot_1.png ; Not reposted 
        if (ErrorLevel = 0) {
            imageFoundBySecondMethod := False
        } else {
            tabsFound1 := ""
            tabsFound2 := ""
        }

        ; Поиск второго изображения
        WinActivate, ahk_exe anty.exe
        ImageSearch, tabs1Found1, tabs1Found2, 0, 0, Width, Height, *50 %A_ScriptDir%\img\Screenshot_2.png ; already reposted
        if (ErrorLevel = 0) {
            imageFoundBySecondMethod := True
        } else {
            tabs1Found1 := ""
            tabs1Found2 := ""
        }

        ; Проверка координат и выполнение соответствующей функции
        if(tabsFound1 <> "" && tabs1Found1 <> "") {
            if(tabsFound2 < tabs1Found2) {
                tabsEmpty := true
                Scroll := True
                FirstCaseRepost(tabsFound1, tabsFound2)
                NextPageGo := True
                Break
            } else {
                tabsEmpty := true
                Scroll := True
                SecondCaseRepost(tabs1Found1, tabs1Found2)
                NextPageGo := True
                Break
            }
        } else if(tabsFound1 <> "") {
            tabsEmpty := true
            Scroll := True
            FirstCaseRepost(tabsFound1, tabsFound2)
            NextPageGo := True
            Break
        } else if(tabs1Found1 <> "") {
            tabsEmpty := true
            Scroll := True
            SecondCaseRepost(tabs1Found1, tabs1Found2)
            NextPageGo := True
            Break
        }

        if((A_TickCount - startWork) > 3500) {
            MsgBox, Result Null. Work stopped.
            switchButton := False
            Break
        } 
    } 
    Return
} 

SetTimer, NextPage, 001
NextPage() {
    if(NextPageGo = True) {
        TabWork()
    }
}

FirstCaseRepost(cord1, cord2) { ; repost one time
    ControlClick, x%cord1% y%cord2% , ahk_exe anty.exe ,,Left, 1
    sleep, 180
    ControlClick, x%cord1% y%cord2% , ahk_exe anty.exe ,,Left, 1
    sleep 180 
    ControlSend, ,{CtrlDown}w{CtrlUp}, ahk_exe anty.exe
}

SecondCaseRepost(cord1, cord2) { ; repost if already reposted
    ControlClick, x%cord1% y%cord2% , ahk_exe anty.exe ,,Left, 1
    sleep, 180
    ControlClick, x%cord1% y%cord2% , ahk_exe anty.exe ,,Left, 1
    sleep 100 
    ControlClick, x%cord1% y%cord2% , ahk_exe anty.exe ,,Left, 1
    sleep, 180
    ControlClick, x%cord1% y%cord2% , ahk_exe anty.exe ,,Left, 1
    sleep, 180
    ControlSend, ,{CtrlDown}w{CtrlUp}, ahk_exe anty.exe
}

