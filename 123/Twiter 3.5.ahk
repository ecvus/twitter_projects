#MaxThreadsPerHotkey 0
#MaxHotkeysPerInterval 1000000000
#Persistent 

SetBatchLines, -1
SetTitleMatchMode, 2

;;***********Exel Create*************
global XL := ComObjCreate("Excel.Application") ; Создаем новый COM объект Excel
XL.Visible := True
global WB := XL.Workbooks.open(A_ScriptDir . "\\Settings.xlsx")
global WS := WB.Worksheets(1) ; Выбираем первый лист

;;*********Exel Settings*************

ShowExcelTable() {
    global XL, WS
    XL.Visible := !XL.Visible
    Return
}

;******Exel Range*******************

global accountCount := Round(WS.Range("G2").Value)

Global ValueOfA2, ValueOfA3, ValueOfA4, ValueOfA5, ValueOfA6, ValueOfA7, ValueOfA8, ValueOfA9, ValueOfA10, ValueOfA11, ValueOfA12, ValueOfA13, ValueOfA14, ValueOfA15, ValueOfA16, ValueOfA17, ValueOfA18, ValueOfA19, ValueOfA20, ValueOfA21, ValueOfA22

UpdateData()

;;********Variables*****************

global Score := 00
global count := 0
global accounts := [ValueOfA2, ValueOfA3, ValueOfA4, ValueOfA5, ValueOfA6, ValueOfA7, ValueOfA8, ValueOfA9, ValueOfA10, ValueOfA11, ValueOfA12, ValueOfA13, ValueOfA14, VlaueOfA15, VlaueOfA16, ValueOfA17, ValueOfA18, ValueOfA19, ValueOfA20, ValueOfA21, ValueOfA22]
global Circles := 0
global Rounds := 1
global Value := accounts.Length()
global TimeRN := Round(A_Hour)

;;**********Timers*************

SetTimer, CheckButtonClick, 001
SetTimer, CheckScore, 100
SetTimer, alert, 100
SetTimer, CheckCount, 100
SetTimer, tipForMouse, 001

;;********Create gui***********

Gui, Font, s12 bold, Arial 
Gui, color, 0E5669
Gui, Show, x1400 y200 w200 h200
Gui, -LastFound +AlwaysOnTop +Caption +Border 

;;********Inside gui***********    
Gui, add, text, x75 y40, GROUP
Gui, add, text, x97 y60 VScore1, %Score%
Gui, add, Button, x95 y80 h20 w20 gRestart, R
Gui, add, Button, x75 y80 h20 w20 gMinus, -
Gui, add, Button, x115 y80 h20 w20 gPlus, +
Gui, add, Button, x71 y140 h20 w70 gcodeExit, Exit
Gui, add, Button, x49 y140 h20 w20 gShowExcelTable, #
Gui, add, Button, x49 y115 h20 w20 gAccountPluse, +
Gui, add, Button, x79 y115 h20 w20 gAccountMinus, -
Gui, add, text, x110 y115 h20 w300 VCount1, %presentAccount%
Gui, add, Button, x2 y2 h20 w20 gUpdateData, ↻

;;********Function************************************************************************************************

UpdateData(){
    ValueOfA2 := WS.Range("A2").Value
    ValueOfA3 := WS.Range("A3").Value
    ValueOfA4 := WS.Range("A4").Value
    ValueOfA5 := WS.Range("A5").Value
    ValueOfA6 := WS.Range("A6").Value
    ValueOfA7 := WS.Range("A7").Value
    ValueOfA8 := WS.Range("A8").Value
    ValueOfA9 := WS.Range("A9").Value
    ValueOfA9 := WS.Range("A9").Value
    ValueOfA10 := WS.Range("A10").Value
    ValueOfA11 := WS.Range("A11").Value
    ValueOfA12 := WS.Range("A12").Value
    ValueOfA13 := WS.Range("A13").Value
    ValueOfA14 := WS.Range("A14").Value
    ValueOfA15 := WS.Range("A15").Value
    ValueOfA16 := WS.Range("A16").Value
    ValueOfA17 := WS.Range("A17").Value
    ValueOfA18 := WS.Range("A18").Value
    ValueOfA19 := WS.Range("A19").Value
    ValueOfA20 := WS.Range("A20").Value
    ValueOfA21 := WS.Range("A21").Value
    ValueOfA22 := WS.Range("A22").Value
    accounts := [ValueOfA2, ValueOfA3, ValueOfA4, ValueOfA5, ValueOfA6, ValueOfA7, ValueOfA8, ValueOfA9, ValueOfA10, ValueOfA11, ValueOfA12, ValueOfA13, ValueOfA14, VlaueOfA15, VlaueOfA16, ValueOfA17, ValueOfA18, ValueOfA19, ValueOfA20, ValueOfA21, ValueOfA22]
    Return
}

CheckCount(){
    presentAccount := accounts[count + 1]
    GuiControl,text, Count1, %presentAccount%
    Return
}

changeExelVision(){
    visibleExel := !visibleExel
    Return
}

timeCheck:
    If(TimeRN > 7){
        TimeRN := "8:00"
    }
    Else{
        TimeRN := "6:00"
    }
Return

codeExit(){ ;; Exit code
    MsgBox, 3, EXIT, Save result?
    IfMsgBox, Yes
    {
        StartRow := 2
        EndRow := 32
        Column := "I"

        Loop, %EndRow%
        {
            Row := A_Index + StartRow - 1
            Cell := WS.Range(Column . Row).Value
            If(Cell = "" )
            {
                WS.Range(Column . Row).Value := WS.Range("E1").Value
                WB.Save
                Break
            }
        }

        Gosub, timeCheck
        textForCopy := "Off " TimeRN "`n"WS.Range("A2").Value "r"Round(WS.Range("B2").Value) "`n"WS.Range("A3").Value " " "r"Round(WS.Range("B3").Value) "`n"WS.Range("A4").Value " " "r"Round(WS.Range("B4").Value) 
        textForCopy .= "`n"WS.Range("A5").Value " " "r"Round(WS.Range("B5").Value) "`n"WS.Range("A6").Value " " "r"Round(WS.Range("B6").Value) "`n"WS.Range("A7").Value " " "r"Round(WS.Range("B7").Value) 
        textForCopy .= "`n"WS.Range("A8").Value " " "r"Round(WS.Range("B8").Value) "`n"WS.Range("A9").Value " " "r"Round(WS.Range("B9").Value) "`n"WS.Range("A10").Value " " "r"Round(WS.Range("B10").Value) 
        textForCopy .= "`n"WS.Range("A11").Value " " "r"Round(WS.Range("B11").Value) "`n"WS.Range("A12").Value " " "r"Round(WS.Range("B12").Value) "`n"WS.Range("A13").Value " " "r"Round(WS.Range("B13").Value) 
        textForCopy .= "`n"WS.Range("A14").Value " " "r"Round(WS.Range("B14").Value) "`n"WS.Range("A15").Value " " "r"Round(WS.Range("B15").Value) "`n"WS.Range("A16").Value " " "r"Round(WS.Range("B16").Value) 
        textForCopy .= "`n"WS.Range("A17").Value " " "r"Round(WS.Range("B17").Value) "`n"WS.Range("A18").Value " " "r"Round(WS.Range("B18").Value) "`n"WS.Range("A19").Value " " "r"Round(WS.Range("B19").Value) 
        textForCopy .= "`n"WS.Range("A20").Value " " "r"Round(WS.Range("B20").Value) "`n"WS.Range("A21").Value " " "r"Round(WS.Range("B21").Value) "`n"WS.Range("A22").Value " " "r"Round(WS.Range("B22").Value)
        Clipboard := textForCopy

        WS.Range("B2").Value := 0
        WS.Range("B3").Value := 0
        WS.Range("B4").Value := 0
        WS.Range("B5").Value := 0
        WS.Range("B6").Value := 0
        WS.Range("B7").Value := 0
        WS.Range("B8").Value := 0
        WS.Range("B9").Value := 0
        WS.Range("B10").Value := 0
        WS.Range("B11").Value := 0
        WS.Range("B12").Value := 0
        WS.Range("B13").Value := 0
        WS.Range("B14").Value := 0
        WS.Range("B15").Value := 0
        WS.Range("B16").Value := 0
        WS.Range("B17").Value := 0
        WS.Range("B18").Value := 0
        WS.Range("B19").Value := 0
        WS.Range("B20").Value := 0
        WS.Range("B21").Value := 0
        WS.Range("B22").Value := 0
        WB.Save
        XL.Quit
        ExitApp

    }
    IfMsgBox, No 
    {
        WB.Save
        XL.Quit
        ExitApp
    }
}

CheckScore(){
    GuiControl,text, Score1, %Score%
Return
}

Restart(){
    Score = 0
Return
}

Plus(){
    Score++
Return
}

Minus(){
    Score--
Return
}

CheckButtonClick(){
    if(GetKeyState("Ctrl", "P") && GetKeyState("V", "P"))
    {

        Score++
        SetTimer, CheckButtonClick, Off
        Sleep, 300
        SetTimer, CheckButtonClick, 001

    }

    ; imagePath := "C:\Users\ZISva\OneDrive\Рабочий стол\Team 1\1.jpg"
    ; hBitmap := LoadBitmapFromFile(imagePath)
    ; CopyBitmapToClipboard(hBitmap)

    ; If(Circles = 1) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C3").Value
    ; }
    ; If(Circles = 2) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C4").Value
    ; }
    ; If(Circles = 3) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C5").Value
    ; }
    ; If(Circles = 4) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C6").Value
    ; }
    ; If(Circles = 5) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C7").Value
    ; }
    ; If(Circles = 6) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C8").Value
    ; }
    ; If(Circles = 7) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C9").Value
    ; }
    ; If(Circles = 8) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C10").Value
    ; }
    ; If(Circles = 9) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C11").Value
    ; }
    ; If(Circles = 10) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C12").Value
    ; }
    ; If(Circles = 11) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C13").Value
    ; }
    ; If(Circles = 12) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C14").Value
    ; }
    ; If(Circles = 13) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C15").Value
    ; }
    ; If(Circles = 14) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C16").Value
    ; }
    ; If(Circles = 15) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C17").Value
    ; }
    ; If(Circles = 16) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C18").Value
    ; }
    ; If(Circles = 17) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C19").Value
    ; }
    ; If(Circles = 18) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C20").Value
    ; }
    ; If(Circles = 19) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C21").Value
    ; }
    ; If(Circles = 20) && WinActive("ahk_exe anty.exe", "", "ExcludeTitle")
    ; {
    ;     Clipboard := WS.Range("C22").Value
    ; }    
}

alert(){
    if(Score = 13)
    {
        MsgBox,,Done, Done, 0.5
        Score = 0
        SoundPlay, sound.mp3
    }
}
Return
;;**********************************Telegram Auto Sender**********************************************************************************

\::
    UpdateData()

    ValueOfF2 := WS.Range("F2").Value ;; Name group to send 

    RoundCountB2 := WS.Range("B2").Value
    RoundCountB3 := WS.Range("B3").Value
    RoundCountB4 := WS.Range("B4").Value
    RoundCountB5 := WS.Range("B5").Value
    RoundCountB6 := WS.Range("B6").Value
    RoundCountB7 := WS.Range("B7").Value
    RoundCountB8 := WS.Range("B8").Value
    RoundCountB9 := WS.Range("B9").Value
    RoundCountB10 := WS.Range("B10").Value
    RoundCountB11 := WS.Range("B11").Value
    RoundCountB12 := WS.Range("B12").Value
    RoundCountB13 := WS.Range("B13").Value
    RoundCountB14 := WS.Range("B14").Value
    RoundCountB15 := WS.Range("B15").Value
    RoundCountB16 := WS.Range("B16").Value
    RoundCountB17 := WS.Range("B17").Value
    RoundCountB18 := WS.Range("B18").Value
    RoundCountB19 := WS.Range("B19").Value
    RoundCountB20 := WS.Range("B20").Value
    RoundCountB21 := WS.Range("B21").Value
    RoundCountB22 := WS.Range("B22").Value

    msgForSend := accounts[count + 1] 

    count++
    Circles++

    if(Circles = 1)
    {
        global sumValue := Round(RoundCountB2 + 1)
        WS.Range("B2").Value := sumValue
        WB.Save
    }
    Else if(Circles = 2)
    {
        global sumValue := Round(RoundCountB3 + 1)
        WS.Range("B3").Value := sumValue
        WB.Save
    }
    Else if(Circles = 3)
    {
        global sumValue := Round(RoundCountB4 + 1)
        WS.Range("B4").Value := sumValue
        WB.Save
    }
    Else if(Circles = 4)
    {
        global sumValue := Round(RoundCountB5 + 1)
        WS.Range("B5").Value := sumValue
        WB.Save
    }
    Else if(Circles = 5)
    {
        global sumValue := Round(RoundCountB6 + 1)
        WS.Range("B6").Value := sumValue
        WB.Save
    }
    Else if(Circles = 6)
    {
        global sumValue := Round(RoundCountB7 + 1)
        WS.Range("B7").Value := sumValue
        WB.Save
    }
    Else if(Circles = 7)
    {
        global sumValue := Round(RoundCountB8 + 1)
        WS.Range("B8").Value := sumValue
        WB.Save
    }
    Else if(Circles = 8)
    {
        global sumValue := Round(RoundCountB9 + 1)
        WS.Range("B9").Value := sumValue
        WB.Save
    }
    Else if(Circles = 9)
    {
        global sumValue := Round(RoundCountB10 + 1)
        WS.Range("B10").Value := sumValue
        WB.Save
    }
    Else if(Circles = 10)
    {
        global sumValue := Round(RoundCountB11 + 1)
        WS.Range("B11").Value := sumValue
        WB.Save
    }
    Else if(Circles = 11)
    {
        global sumValue := Round(RoundCountB12 + 1)
        WS.Range("B12").Value := sumValue
        WB.Save
    }
    Else if(Circles = 12)
    {
        global sumValue := Round(RoundCountB13 + 1)
        WS.Range("B13").Value := sumValue
        WB.Save
    }
    Else if(Circles = 13)
    {
        global sumValue := Round(RoundCountB14 + 1)
        WS.Range("B14").Value := sumValue
        WB.Save
    }
    Else if(Circles = 14)
    {
        global sumValue := Round(RoundCountB15 + 1)
        WS.Range("B15").Value := sumValue
        WB.Save
    }
    Else if(Circles = 15)
    {
        global sumValue := Round(RoundCountB16 + 1)
        WS.Range("B16").Value := sumValue
        WB.Save
    }
    Else if(Circles = 16)
    {
        global sumValue := Round(RoundCountB17 + 1)
        WS.Range("B17").Value := sumValue
        WB.Save
    }
    Else if(Circles = 17)
    {
        global sumValue := Round(RoundCountB18 + 1)
        WS.Range("B18").Value := sumValue
        WB.Save
    }
    Else if(Circles = 18)
    {
        global sumValue := Round(RoundCountB19 + 1)
        WS.Range("B19").Value := sumValue
        WB.Save
    }
    Else if(Circles = 19)
    {
        global sumValue := Round(RoundCountB20 + 1)
        WS.Range("B20").Value := sumValue
        WB.Save
    }
    Else if(Circles = 20)
    {
        global sumValue := Round(RoundCountB21 + 1)
        WS.Range("B21").Value := sumValue
        WB.Save
    }
    Else if(Circles = 21)
    {
        global sumValue := Round(RoundCountB22 + 1)
        WS.Range("B22").Value := sumValue
        WB.Save
    }

    WinActivate, ahk_exe telegram.exe
    SetKeyDelay, 1, 1
    ControlSend,,{Esc}{Esc}, ahk_exe telegram.exe
    sleep, 200

    ControlSend,,%ValueOfF2%, ahk_exe telegram.exe
    sleep, 100
    ControlSend,,{Enter}, ahk_exe telegram.exe

    ControlSend,,%msgForSend% r%sumValue%, ahk_exe telegram.exe 
    sleep, 100
    ControlSend,,{Enter}, ahk_exe telegram.exe
    sleep, 100
    SetKeyDelay, 5

    if(Count >= WS.Range("G2").Value || Count = -1)
    {
        count := 0

    }

    If(Circles >= WS.Range("G2").Value || Circles = -1)
    {

        Circles := 0
    }

Return

AccountPluse:
    UpdateData()
    Circles++
    count++

    if(Count >= WS.Range("G2").Value || Count = -1)
    {
        count := 0

    }

    If(Circles >= WS.Range("G2").Value || Circles = -1)
    {

        Circles := 0
    }

Return

AccountMinus:
    UpdateData()
    Circles--
    count--

    if(Count >= WS.Range("G2").Value || Count = -1)
    {
        count := 0

    }

    If(Circles >= WS.Range("G2").Value || Circles = -1)
    {

        Circles := 0
    }

Return

;**********************************************Mouse tip for alt*********************************

tipForMouse(){ 
    i := Circles + 1
    If(GetKeyState("Alt", "P"))
    {
        a := Round(WS.Range("E1").Value)
        ToolTip, G: %Score%`nR: %a%`nA: %i%/%accountCount%

    }
    Else{
        ToolTip
    }
Return
}

; ***************************************Tab Functions*********************************

Tab::
    a := %A_TickCount%

    i := False
    While(i = False){
        ImageSearch, x1, y2, 0, 0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%/img/Screenshot_1.png
        if(x <> "")
        {
            i := True
            Send, t 
            sleep, 110
            Send, {Enter}
            sleep, 150
            Send, ^w
        }

        ImageSearch, x2, y2, 0, 0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%/img/Screenshot_2.png

        if(x2 <> "")
        {
            i := True
            Send, t 
            sleep, 110
            Send, {Enter}
            sleep, 170
            Send, t 
            sleep, 120
            Send, {Enter}
            sleep, 170
            Send, ^w
        }
        if((A_TickCount - a) > 1500){
            MsgBox, not founded
            Break
        }

    }
Return 
