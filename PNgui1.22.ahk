;
; AutoHotkey Version: 1.048
; Language:       English
; Platform:       Win9x/NT/XP
; Author:         H. Tenkink <tenkink@jive.nl>
;
;   Includes:
;   PN[version].ahk  -- The main program.
;   SerialPort.ahk   -- The program that takes care of the RS232 protocol.
;   Paternoster.ahk  -- The paternoster control program.

#SingleInstance force ; If it is alReady Running it will be restarted.
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode,2 ; string somewhere in titel is ok.
SetTitleMatchMode,fast  ; makes window recognition more reliable
DetectHiddenWindows, On
SetKeyDelay,0  ; Sets the speed of characters sent to CMD window.
;SetControlDelay, 150

IfNotExist, PN.ini
  {
    SoundPlay, Low.wav
    MsgBox,16,, PN.ini not found
    ExitApp,1
  }

; IniRead, OutputVar, Filename, Section, Key [, Default]
;For Paternoster initialization
IniRead, PNVersion, PN.ini, PN, PNVersion
IniRead, Init, PN.ini, PaternosterStrings, Init ; Read string from PN.ini
IniRead, Fill, PN.ini, PaternosterStrings, Fill 
IniRead, NrCariers, PN.ini, PaternosterStrings, NrCariers
IniRead, FillInit, PN.ini, PaternosterStrings, FillInit
;For paternoster movement.
IniRead, Pre, PN.ini, PaternosterStrings, Pre 
IniRead, Post, PN.ini, PaternosterStrings, Post
IniRead, End, PN.ini, PaternosterStrings, End
IniRead, Up, PN.ini, Paternosterstrings, Up 
IniRead, Stop, PN.ini, PaternosterStrings, Stop 
IniRead, Down, PN.ini, PaternosterStrings, Down 
IniRead, Fill21, PN.ini, PaternosterStrings, Fill21
IniRead, CurrentCarier, PN.ini, CurrentCarier, Carier 
IniRead, IdleTime, PN.ini, Timer, IdleTime 
;(The paternoster layout is read in PN[version].ahk)
;For external contact:
IniRead, RemoteMachine, PN.ini, Remote, RemoteMachine
IniRead, Pad, PN.ini, Remote, Pad
IniRead, Testcommand, PN.ini, Remote, Testcommand
IniRead, Start, PN.ini, Remote, start
; Serial ports intitialization
IniRead, RS232_Port, PN.ini, Laserpointer, RS232_Port
IniRead, RS232_Baud, PN.ini, Laserpointer, RS232_Baud
Gosub, InitSerialPort   ; Initialize the laser pointer
LP_Port        = %RS232_FileHandle%
IniRead, RS232_Port, PN.ini, Paternoster, RS232_Port
IniRead, RS232_Baud, PN.ini, Paternoster, RS232_Baud
Gosub, InitSerialPort    ; Initialize the paternoster
PN_Port        = %RS232_FileHandle%
IniRead, Logfile, PN.ini, Files, Logfile
IniRead, VSNReceive, PN.ini, Files, VSNReceive
IniRead, VSNShip, PN.ini, Files, VSNShip
IniRead, projectfile, PN.ini, Files, projectfile  ; Temporary file

; First check if all necessary programs are there.
IfNotExist, c:\Program Files\PuTTY\putty.exe
  {
    SoundPlay, Low.wav
    MsgBox,16,,putty.exe not found
    ExitApp,1
  }

IfNotExist, private2.ppk
  {
    SoundPlay, Low.wav
    MsgBox,16,, private2.ppk not found
    ExitApp,1
  }

IfNotExist, TrackSend.ahk
  {
    SoundPlay, Low.wav
    MsgBox,16,, TrackSend.ahk not found
    ExitApp,1
  }

IfNotExist, TrackReceive.ahk
  {
    SoundPlay, Low.wav
    MsgBox,16,, TrackReceive.ahk not found
    ExitApp,1
  }
  
IfNotExist, SerialPort.ahk
  {
    SoundPlay, Low.wav
    MsgBox,16,, SerialPort.ahk not found
    ExitApp,1
  }
  
IfNotExist, PN.ini
  {
    SoundPlay, Low.wav
    MsgBox,16,, PN.ini not found
    ExitApp,1
  }
  
Process, close , putty.exe
sleep, 100
Process, Exist, putty.exe
  If errorlevel
  {
    MsgBox,64,, Error closing putty.exe (PID %errorlevel%). 
    FileAppend, [%timestamp%]  Error0 closing putty.exe (PID %errorlevel%) at startup`n, %logfile%
    ControlSend,,Exit{Enter}, %RemoteMachine% ; kill putty nicely
  }
  
IfExist, putty.log
{
FileDelete, putty.log
if errorlevel
{
    FileAppend, [%timestamp%] Error, putty.log could not be deleted at startup `n, %logfile%
    MsgBox,16, FileDelete Error, putty.log could not be deleted.`n`n   * Terminating program *
    ExitApp, 1
  }
}

; Set some parameters
Station = 1 ;Souterain
SetPath = cd %pad%
CommandPrompt = %RemoteMachine%:~/<1>disks/update>
Run, %start%,,hide UseErrorLevel  ; start putty
if errorlevel
  MsgBox, 16, Error %ErrorLevel%, Putty will not start.
ToolTip, Starting Putty
sleep, 1000   ; wait for putty to start
StartTimePutty := A_TickCount   ; timer start
Loop   ; wait until we get a prompt.
  {
    Sleep, 500
    FileRead, retText, putty.log ; Read log file to see if putty started.
    StringGetPos, pos, retText, %RemoteMachine%:~>,R
    If not Errorlevel  ; if we see a prompt
      {
        ControlSend,,%SetPath%{Enter}, %RemoteMachine%
        ; ControlSend,,{Enter}, %RemoteMachine%
        Sleep, 200
        FileRead, retText, putty.log
        IfNotInString, retText, %SetPath%
            MsgBox, 16, Putty Error, Command: %SetPath% not Send.`n %retText%
        StringReplace retget, retText, %CommandPrompt% ,,A, UseErrorLevel    ; remove command prompt
        If ErrorLevel = 0
            MsgBox, 16, Putty Error, %CommandPrompt% not found `n %retText%.
        Break
      }
    ;MsgBox,,%retText%
    ElapsedTimePutty := A_TickCount - StartTimePutty  ; timer end
    EnvDiv, ElapsedTimePutty, 1000
    tooltip, Waiting %ElapsedTimePutty% s for Putty  , 300, 300
    SetTimer, HideToolTip, 5000

    If A_Index > 20
      {
        SoundPlay, Low.wav
        MsgBox,48, Error, Waited %ElapsedTimePutty% s for putty to start.`n Please try again.
        ExitApp
      }
  }

FormatTime, timestamp, %A_Now%, dd-MM-yy HH:mm:ss
if ElapsedTimePutty > 10000
FileAppend, [%timestamp%] Starting Putty took %ElapsedTimePutty% ms`n, %logfile%

FormatTime, timestamp, %A_Now%, dd-MM-yy HH:mm:ss
FileAppend, `n[%timestamp%]   < - StartPN %PNVersion% - >`n, %Logfile%

; Create the sub-menus for the menu bar:
Menu, ExitMenu, Add, E&xit, GuiClose
Menu, HelpMenu, Add, &Help   F1, Help
Menu, HelpMenu, Add, &About, HelpAbout

; Create the menu bar by attaching the sub-menus to it:
Menu, MyMenuBar, Add, &Exit, :ExitMenu
Menu, MyMenuBar, Add, &Help, :HelpMenu


;  Define the gui.
Gui:
  Gui, +Resize +MinSize
  Gui, font, s15 C7777aa,  Verdana
  Gui, Menu, MyMenuBar
  Gui, font, s25 C7777aa,  Verdana
  Gui, Add, Text, x20 y10 w180, %choice%  ; experiments
  Gui, font, s12, Courier New Bold
  Gui, Add, ListView, x20 y50 r24 w175 gListView, Pos|VSN
  Gui, Font, S60 C9999bb, Verdana
  Gui, Add, Text , x250 y68 w350 vAction
  Gui, font, s11 CDefault, Verdana  ; Preferred font.
  Gui, Add, Edit, x70 y535 w125 vAmount ;Readonly
  Gui, Add, Button, x20 y533 Ok, OK
  ;Gui, Add, Button, x180 y533 End, End
  ;Gui, Add, Text, x250 y185, VSN
  Gui, Add, GroupBox, x242 y187 w425 h130, VSN
  Gui, Font, S60 CDefault, Courier New Bold
  ;Gui, Font, S50 CDefault, Lucida console Bold
  Gui, Add, Edit, x250 y207 w410 Uppercase Limit18 vInput
  Gui, Font, S15 CDefault, Verdana
  Gui, Add, Button, x250 y15 w105 h50 Project, Project
  Gui, Add, Button, x365 y15 w70 h50, &Get
  Gui, Add, Button, x445 y15 w70 h50 Put, &Put
  Gui, Add, Button, x525 y15 w105 h50 Cancel, &Cancel
  ;Gui, Font, S10 CDefault, Verdana
  Gui, Add, GroupBox, x693 y117 w124 h46 border vexpedition
  Gui, Add, Button, x695 y120 w120 h40 vSouterain, Souterain
  Gui, Font, S10 CDefault, Verdana
  Gui, Add, GroupBox, x663 y13 w157 h80, Shipping
  Gui, Font, S12 CDefault, Verdana
  Gui, Add, Checkbox, x670 y30 vReceive gCheckBoxC, Receive
  Gui, Add, Checkbox, x670 y50 vRemove gCheckBoxR, Remove
  Gui, Add, Checkbox, x670 y70 vSend gCheckBoxS, Send
  Gui, Add, Button, x770 y35 w45 h40 gEnd, End
  GuiControl, Disable, End
  Gui, Add, Button, x740 y190 w75 h40 gUp, &Up
  Gui, Add, Button, x740 y240 w75 h40 gStop, &Stop
  Gui, Add, Button, x740 y290 w75 h40 gDown, &Down
  ;Gui, Add, Text, x250 y345 w120, History
  Gui, Add, GroupBox, x242 y337 w425 h195, History
  Gui, Font, S15 CDefault, Courier New Bold
  Gui, Add, ListBox, x250 y357 w410 h170 vOldVSN
  Gui, Font, S10 CDefault, Verdana
  ;Gui, Color, Default, White
  Gui, Add, GroupBox, x733 y343 w88 h165, Carier
  Gui, Font, S20 CDefault, Verdana
  Gui, Add, Edit, x740 y410 w50 limit 2 number left vCarier
  Gui, Font, S12 CDefault, Verdana
  Gui, Add, Button, x740 y360 w75 h40, Go
 ; Gui, Add, Button, x670 y440 w65 h40, Dummy
  Gui, Add, Button, x740 y460 w75 h40 Init, Init
  Gui, Font, S10 CDefault, Verdana
  Gui, Add, GroupBox, x362 y542 w305 h60, System messages
  Gui, Font, S15 C4444ff, Verdana
  Gui, Add, Edit, x370 y560 w240 vSystem, %Semaphore%   ; System messages
  Gui, Add, Edit, x620 y560 w40 vReturnedCarier, %ReturnedCarier%    ; Carier Returned from PN.
  Gui, Font, s12, MS sans serif
  Gui, Add, StatusBar,,
  Gui, Add, Progress, vMyProgress w100 h20 x638 y610 Range0-%IdleTime% ;-Smooth
  Gui, Add, Button, Default Hidden, Enter
  Gui, Show, w840 h630, PN %PNVersion%
  Gui Add, Edit, x250 y15 w120 h25 vsearchedString gIncrementalSearch hidden ; Edit6
  Gui Add, ListBox, x250 y38 w120 R15 vchoice gListBoxClick sort hidden  ; listbox2
  Gui Add, Button, gListBoxClick hidden, Ok
  
  ;GuiControl, Focus, Get
  GuiControl, disable, Input
  GuiControl, hide, expedition
anchor("h600")
anchor("")
  SetTimer, Listen, 500   ;  Listen to the Serial port every 0.5 second.
  SetTimer, idle_check, 1000
  SetTimer, Tip, 2000
  SB_SetParts( 200, 300, 70, 70, 100) ; divide the statusbar in 5 parts.
  
  ToolTip     ;stop tooltip
  GuiControl,, Carier, %CurrentCarier%  ; Update the Input.
  Gosub, Stop
 Return 

; When running idle for IdleTime restart Putty if the putty.log file is getting too big.
idle_check:
GuiControlGet, Receive
GuiControlGet, Remove
GuiControlGet, Send
FileGetSize, PuttySize, putty.log, k
Itime := round(A_TimeIdle / 1000)
GuiControl,, MyProgress, %ITime%
SB_SetIcon("file.ico",1,4)
SB_SetText(PuttySize "k", 4) ;show Putty logfile size in status bar
IfGreater, ITime, %IdleTime%
{
  IfNotEqual, Pos, Off   ; Laserpointer is still on.
  {
    Pos = Off
    ;Laser := LaserPointer(Pos)
    If Send = 0
    gosub ButtonCancel
  }
  FileGetSize, PuttySize, putty.log, k
  if PuttySize > 20  ; If putty logfile is getting too big.
     gosub RestartPutty
}
Return


GuiSize:
  Anchor("SysListView321", "h")        ; Project list
  Anchor("OK", "y")                    ; OK
  Anchor("History", "h")               ; Groupbox History
  Anchor("MySystemMessagesGroup", "y") ; GroupBox system messages
  Anchor("Edit1", "y")                 ; Amount of disks in project
  Anchor("System", "y")                ; System messages
  Anchor("Edit5", "y", true)           ; returned carier
  Anchor("OldVSN", "h")                ; History
  Anchor("msctls_progress321", "y")    ; progress bar
Return


;########################################################################
;###### Anchor   ########################################################
;########################################################################
/*
;Function: Anchor
;Defines how controls should be automatically positioned relative to the new dimensions of a window when resized.

;Parameters:
;cl - a control HWND, associated variable name or ClassNN to operate on
a - (optional) one or more of the anchors: 'x', 'y', 'w' (width) and 'h' (height),
optionally followed by a relative factor, e.g. "x h0.5"
r - (optional) true to redraw controls, recommended for GroupBox and Button types

;Examples:
;> "xy" ; bounds a control to the bottom-left edge of the window
;> "w0.5" ; any change in the width of the window will resize the width of the control on a 2:1 ratio
;> "h" ; similar to above but directly proportional to height

;Remarks:
;To assume the current window size for the new bounds of a control (i.e. resetting) simply omit the second and third parameters.
;However if the control had been created with DllCall() and has its own parent window,
;the container AutoHotkey created Gui must be made default with the +LastFound option prior to the call.
;For a complete example see anchor-example.ahk.

;License:
;- Version 4.60a <http://www.autohotkey.net/~Titan/#anchor>
- Simplified BSD License <http://www.autohotkey.net/~Titan/license.txt>
*/
Anchor(i, a = "", r = false) {
    static c, cs = 12, cx = 255, cl = 0, g, gs = 8, gl = 0, gpi, gw, gh, z = 0, k = 0xffff
    If z = 0
        VarSetCapacity(g, gs * 99, 0), VarSetCapacity(c, cs * cx, 0), z := true
    If (!WinExist("ahk_id" . i)) {
        GuiControlGet, t, Hwnd, %i%
        If ErrorLevel = 0
            i := t
        Else ControlGet, i, Hwnd, , %i%
      }
    VarSetCapacity(gi, 68, 0), DllCall("GetWindowInfo", "UInt", gp := DllCall("GetParent", "UInt", i), "UInt", &gi)
            , giw := NumGet(gi, 28, "Int") - NumGet(gi, 20, "Int"), gih := NumGet(gi, 32, "Int") - NumGet(gi, 24, "Int")
    If (gp != gpi) {
        gpi := gp
        Loop, %gl%
            If (NumGet(g, cb := gs * (A_Index - 1)) == gp) {
                gw := NumGet(g, cb + 4, "Short"), gh := NumGet(g, cb + 6, "Short"), gf := 1
                Break
              }
        If (!gf)
            NumPut(gp, g, gl), NumPut(gw := giw, g, gl + 4, "Short"), NumPut(gh := gih, g, gl + 6, "Short"), gl += gs
      }
    ControlGetPos, dx, dy, dw, dh, , ahk_id %i%
    Loop, %cl%
        If (NumGet(c, cb := cs * (A_Index - 1)) == i) {
            If a =
              {
                cf = 1
                Break
              }
            giw -= gw, gih -= gh, as := 1, dx := NumGet(c, cb + 4, "Short"), dy := NumGet(c, cb + 6, "Short")
                    , cw := dw, dw := NumGet(c, cb + 8, "Short"), ch := dh, dh := NumGet(c, cb + 10, "Short")
            Loop, Parse, a, xywh
                If A_Index > 1
                    av := SubStr(a, as, 1), as += 1 + StrLen(A_LoopField)
                        , d%av% += (InStr("yh", av) ? gih : giw) * (A_LoopField + 0 ? A_LoopField : 1)
            DllCall("SetWindowPos", "UInt", i, "Int", 0, "Int", dx, "Int", dy
                    , "Int", InStr(a, "w") ? dw : cw, "Int", InStr(a, "h") ? dh : ch, "Int", 4)
            If r != 0
                DllCall("RedrawWindow", "UInt", i, "UInt", 0, "UInt", 0, "UInt", 0x0101) ; RDW_UPDATENOW | RDW_INVALIDATE
            Return
          }
    If cf != 1
        cb := cl, cl += cs
    bx := NumGet(gi, 48), by := NumGet(gi, 16, "Int") - NumGet(gi, 8, "Int") - gih - NumGet(gi, 52)
    If cf = 1
        dw -= giw - gw, dh -= gih - gh
    NumPut(i, c, cb), NumPut(dx - bx, c, cb + 4, "Short"), NumPut(dy - by, c, cb + 6, "Short")
            , NumPut(dw, c, cb + 8, "Short"), NumPut(dh, c, cb + 10, "Short")
    Return, true
  }

Tip:
  IfWinNotActive, PN %PNVersion%
    {
      ToolTip
      Return
    }

  MouseGetPos,,,, ACtrl
  Tip =
  If ACtrl = Button6
      Tip = Abort operation

  If ACtrl = Button13
      Tip = Finish shipping operation

  If ACtrl = Button20
      Tip = Intitialize the paternoster
  If ACtrl = Edit2
    { 
    ControlGet, status, Enabled,, Edit2
    if status = 0
      Tip = Input field `nSelect an action to activate
    else
      Tip = Input field
    }
    
  If ACtrl = Edit3
      Tip = Last wanted shelf

  If ACtrl = Edit5
      Tip = Last Reported shelf

  If ACtrl = msctls_StatusBar321
    {
      MouseGetPos, XCtrl
      If XCtrl Between 5 and 200
          Tip = Read VSN

      If XCtrl Between 205 and 500
          Tip = Checked VSN

      If XCtrl Between 505 and 570
          Tip = Position

      If XCtrl Between 575 and 640
          Tip = Putty.log size
          
      If XCtrl <> %LastXCtrl%
        {
          ToolTip, %Tip%
          LastXCtrl = %XCtrl%
          SetTimer, tipof, -4000
        }
    }
  If ACtrl = msctls_Progress321
      Tip = timeout

  If ACtrl <> %LastCtrl%
    {
      ToolTip, %Tip%
      LastCtrl = %ACtrl%
      SetTimer, tipof, -4000
    }
Return

tipof:
  ToolTip
return

; ..... Define the button actions  ........................................................

~*F1::
Help:
  Run, %A_ScriptDir%\PN.chm
Return

HelpAbout:
  ;Gui, 2:+owner1  ; Make the main window (Gui #1) the owner of the "about box" (Gui #2).
  Gui +Disabled  ; Disable main window.
  Gui, 2:Add, Text,, PN  Version %PNVersion%`n`nDate: March 2012`nAutoHotkey Version: %A_AhkVersion% `nAuthor: H. Tenkink
  Gui, 2:Add, Button, Default y70, OK
  Gui, 2:Show, h100
Return

2GuiClose:
2ButtonOK:
  Gui, 1:-Disabled
  Gui, 2:Destroy
Return

ButtonEnter:
  Gosub, Input
Return

ButtonProject:
  ;MsgBox, Project
  GuiControl, Disable, OK
  GuiControl, Disable, Project
  GuiControl, Disable, Get
  GuiControl, Disable, Put
  GuiControl, Disable, Receive
  GuiControl, Disable, Remove
  GuiControl, Disable, Send
  GuiControl, Disable, Edit2
  GuiControl,, Action, Project  ; Write the big letter Caption.
  Action = Project
  Gosub, Experiment
Return

ButtonGet:
  ;MsgBox, get
  GuiControl, Disable, OK
  GuiControl, Disable, Project
  GuiControl, Disable, Get
  GuiControl, Disable, Put
  ;GuiControl, Disable, Souterain
  GuiControl, Disable, Receive
  ;GuiControl, Disable, Remove
  ;GuiControl, Disable, Send
  GuiControl,, Action, Get
  Action = Get
  GuiControl, Enable, Input
  GuiControl, Focus, Input
  SoundPlay, c:\windows\media\Start.wav  ; tik!
Return

ButtonPut:
  GuiControl, Disable, OK
  GuiControl, Disable, Project
  GuiControl, Disable, Get
  GuiControl, Disable, Put
  GuiControl, Disable, Remove
  GuiControl, Disable, Send
  Action = Put
  GuiControl,, Action, %Action%
  GuiControl, Enable, Input
  GuiControl, Focus, Input
  SoundPlay, c:\windows\media\Start.wav  ; tik!
Return

ButtonCancel:
  Canceled = 1  ; For Breaking out of confirm loop in Project.
  Gosub, Stop
  GuiControl,, Add,0
  GuiControl,, Receive,0
  GuiControl,, Remove,0
  GuiControl,, Send,0
  GuiControl, Focus, Get
  Action =
  GuiControl,, Action  ; Clear the Action field.
  GuiControl,, Input  ; Clear the Input field.
  GuiControl, Hide, ComboBox1
  Amount = 0 ; end the project
  GuiControl,,static1
  ; And clear the remaining VSNs from list
  Loop, %TotalAmount%
    {
      LV_Delete(1)   ; Delete the not selected Topmost VSN.
      Sleep,50   ; make the scrolling Visible.
    }
  ControlSetText, Edit1   ; clear Count field.
  ControlSetText, static2   ; clear projectname field.
  GuiControl, hide, Edit6   ; project edit
  GuiControl, hide, ListBox2 ; project list
  GuiControl, Enable, Edit2
  GuiControl, Enable, OK
  GuiControl, Enable, VSN
  GuiControl, Enable, Project
  GuiControl, Enable, Get
  GuiControl, Enable, Put
  GuiControl, Enable, Cancel
  GuiControl, Enable, Souterain
  GuiControl, Enable, Receive
  GuiControl, Enable, Remove
  GuiControl, Enable, Send
  GuiControl, Enable, Up
  GuiControl, Enable, Stop
  GuiControl, Enable, Down
  GuiControl, Enable, History
  GuiControl, Enable, Carier
  GuiControl, Enable, Go
  GuiControl, Enable, Init
  GuiControl, Enable, System ; Goupbox System messages
  GuiControl, Enable, ComboBox1
  GuiControl, Disable, Input
  GuiControl, Disable, End
  FileDelete, %projectfile%
SoundPlay, c:\windows\media\start.wav  ; endtune
Return

End:
  ;MsgBox, ready, REC= %Receive%`n Rem= %Remove%`n Sent= %Send%
  If %Receive%  ; Receive is checked.
    {
      IfExist, %VSNReceive%
        {
          FileRead, Received, %VSNReceive% ; This fle is written in 'InputCheck'
          Sort, Received, U  ; Remove duplicates.
          Clipboard = %Received%
          Run, TrackReceive.ahk
        }
      Gosub, ButtonSouterain 
      Gosub, ButtonCancel
    }
  Else
      If %Remove%
          Gosub ButtonCancel
      Else
          If %Send%
            {
              FileGetSize, size, %VSNShip%
              If not Errorlevel
                {
                  FileRead, Send, %VSNShip% ; This fle is written in 'InputCheck'
                  Clipboard = %Send%
                  Run, TrackSend.ahk
                }
              Gosub, ButtonCancel
            }
Return

CheckBoxC:   ;Receive
  Gui, Submit, NoHide
  If Receive = 1 ; Receive is checked.
    {
      IfExist, %VSNReceive%
        {
          FileRecycle, c:\temp\receive.bak
          ;MsgBox, rec = %Receive%
          FileCopy, %VSNReceive%, c:\temp\receive.bak
          FileRecycle, %VSNReceive%
        }
      GuiControl, Disable, OK
      GuiControl, Disable, Project
      GuiControl, Disable, remove
      GuiControl, Disable, Send
      GuiControl, Enable, End
      GuiControlGet, CtrlContents,, Souterain
      If CtrlContents = Souterain
          Gosub, Buttonsouterain
      Gosub, ButtonPut
    }
  Else  ; Receive is UnChecked
    {
      IfExist, %VSNReceive%
        {
          FileRead, Received, %VSNReceive% ; (This fle is written in 'InputCheck')
          Sort, Received, U  ; Remove duplicates.
          Clipboard = %Received%
          Run, TrackReceive.ahk
        }
      Gosub ButtonCancel
    }
Return

CheckBoxR:   ; Remove
  Gui, Submit, NoHide
  If Remove = 1
    {
      GuiControl, Disable, OK
      GuiControl, Disable, Project
      GuiControl, Disable, Put
      GuiControl, Disable, Receive
      GuiControl, Disable, Send
      GuiControl, Enable, End
      GuiControlGet, CtrlContents,, Souterain
      If CtrlContents = Souterain
      Gosub, ButtonSouterain
      Gosub, ButtonGet
      ;GuiControl,, Action, Remove
    }
  Else
      Gosub ButtonCancel
Return

CheckBoxS:   ; Send
  Gui, Submit, NoHide
  If Send = 1
    {
      Action = Send
      IfExist, %VSNShip%
        {
          FileRecycle, c:\temp\Send.bak
          FileCopy, %VSNShip%, c:\temp\Send.bak
          FileRecycle, %VSNShip%
        }
      GuiControl, Disable, OK
      GuiControl, Disable, Project
      GuiControl, Disable, Get
      GuiControl, Disable, Put
      GuiControl, Disable, Remove
      GuiControl, Disable, Receive
      GuiControl, Enable, End
      GuiControl, Enable, History
      GuiControl, Enable, Input
      GuiControl, Focus, Input
      SoundPlay, c:\windows\media\Start.wav  ; tik!
      GuiControl,, Action, %Action%
    }
  Else
    {
      IfExist, %VSNShip%
        {
          FileRead, Send, %VSNShip% ; This fle is written in 'InputCheck'
          Sort, Send, U  ; remove duplicates
          Clipboard = %Send%
          Run, TrackSend.ahk
        }
      Gosub, ButtonCancel
    }
Return

ButtonSouterain:
  GuiControlGet, CtrlContents,, Souterain

  If CtrlContents = Souterain
    {
      Station = 2
      Souterain = Expedition
      GuiControl, show, expedition
    }
  Else
    {
      Station = 1
      Souterain = Souterain
      GuiControl, hide, expedition
    }
  GuiControl,, Souterain, %Souterain%  ; rewrite the Button.
  ;GuiControl,, Souterain, border
  GuiControl,, expedition ;x690 y102 w130 h67
  GuiControl, Focus, Input
  GuiControlGet, OutputVar, Enabled, Up
  if not OutputVar
  Gosub, ButtonGo
Return

Up:
Stat = 0x3%Station%,
Upl = %Up%%Fill21%%Stat%%end%
RS232_Write(PN_Port,Upl)
CurrentCarier++
If CurrentCarier = 33
  CurrentCarier = 1
GuiControl, Focus, Input
Return

Stop:
Stat = 0x3%Station%,
Hstring = %Stop%%Fill21%%Stat%%End%
RS232_Write(PN_Port,Hstring)
Carier = %CurrentCarier%
Pos = Off
LP := LaserPointer(Pos)
;MsgBox, Stop: %Hstring%
GuiControl, Focus, Input
Return

Down:
Stat = 0x3%Station%,
Downl = %Down%%Fill21%%Stat%%end%
RS232_Write(PN_Port,Downl)
CurrentCarier--  ; subtract 1 from CurrentCarier.
If CurrentCarier = 0
  CurrentCarier = 32
GuiControl, Focus, Input
Return

ButtonGo:
  GuiControlGet, CurrentCarier,,Carier  ;Get the Input.
  If ErrorLevel
      MsgBox,16,, Go Error: Carier: %CurrentCarier% not found.
  StringLen, len, CurrentCarier
  If (len = 0)
      MsgBox,48,, Go Carier not given.%CurrentCarier%
  Else
      Answer := Paternoster(CurrentCarier, Station)
      GuiControl, Focus, Input
Return

ButtonInit:
GuiControl, Focus, Carier
   Stat = 0x3%Station%
  ;MsgBox, Station: %Station%
  GuiControlGet, Carier,, Carier
  ;MsgBox, Carier: %Carier%
  if Carier not between 1 and 32
      MsgBox,48,, Give a carier from 1 to 32.
  Else
    {
      HCarier := ASCII_to_Hex(Carier)
      InitString = %Init%%HCarier%%Fill%%NrCariers%%Stat%%FillInit%%End%
     ; MsgBox, Port: %PN_Port% InitString: %InitString%
      RS232_Write(PN_Port,InitString) ;write it to Serial port Com4.
      Semaphore = Initialized
      ReturnedCarier = %Carier%
      CurrentCarier =  %Carier%
    }
Return
;  .......... End of button actions ....................................................


GuiEscape:
GuiClose:  ; Indicate that the script should exit automatically when the window is closed.
  Gosub, ButtonCancel
  Gosub,ClosePN
  ControlSend,,Exit{Enter}, %RemoteMachine%              ; Close console window
  ;Sleep,300
  If Errorlevel
      MsgBox,48, Error %Errorlevel%, in GuiClose: putty.exe still running(0). 
  Sleep,100
  FileDelete, putty.log
  If Errorlevel
      Process, Exist, putty.exe
  If Errorlevel
      ControlSend,,Exit{Enter}, %RemoteMachine% ; kill putty nicely
  If Errorlevel
    {
      FileAppend, [%timestamp%] in GuiClose: putty.exe still Running. Error %Errorlevel% `n, %logfile%
      MsgBox,48,  Error %Errorlevel%, In GuiClose: putty.exe still running(1).
    }
  Sleep,100
  Process, Exist, putty.exe
  If Errorlevel
      ControlSend,,Exit{Enter}, %RemoteMachine% ; kill putty nicely
  If Errorlevel
    {
      FileAppend, [%timestamp%] in GuiClose: putty.exe still Running#2. Error %Errorlevel% `n, %logfile%
      MsgBox,48, Error %Errorlevel%, in GuiClose: putty.exe still Running(2).
    }
  If FileExist("putty.log")
    {
      FileDelete, putty.log
      If Errorlevel
        {
          FileAppend, [%timestamp%] Error %Errorlevel% in GuiClose, putty.log could not be Deleted `n, %logfile%
          ; MsgBox, 48, Error %Errorlevel%, In Guiclose: putty.log could not be deleted
        }
    }
  ExitApp
Return

RestartPutty:
  Process, Exist, putty.exe
  If Errorlevel
      ControlSend,,Exit{Enter}, %RemoteMachine% ; kill putty nicely
  If Errorlevel
    {
      FileAppend, [%timestamp%] in RestartPutty: putty.exe still Running(0). Error %Errorlevel% `n, %logfile%
      MsgBox,48, Error %Errorlevel%, In RestartPutty: putty.exe still Running(0).
    }
  Sleep,100

  Process, Exist, putty.exe
  If Errorlevel
      ControlSend,,Exit{Enter}, %RemoteMachine% ; kill putty nicely
  If Errorlevel
    {
      FileAppend, [%timestamp%] in RestartPutty: putty.exe still running(1). Error %Errorlevel% `n, %logfile%
      MsgBox,48, Error %Errorlevel%, in RestartPutty: putty.exe still running(1).
    }
  Sleep,100
  Process, Exist, putty.exe
  If Errorlevel
    {
      Process, Close , putty.exe  ; kill putty forcefully
      If not Errorlevel
        {
          WinClose, putty.exe  ; kill putty with another knife
          If Errorlevel
            {
              FileAppend, [%timestamp%]  Error winclosing putty.exe (PID %Errorlevel%) in restartputty`n, %logfile%
              MsgBox,64,,in RestartPutty:  Error winclosing putty.exe (PID %Errorlevel%).
            }
        }
    }

If FileExist("putty.log")
    {
      FileDelete, putty.log
      If Errorlevel
        {
          FileAppend, [%timestamp%] Error %Errorlevel% in RestartPutty, putty.log could not be Deleted `n, %logfile%
          ; MsgBox, 48,Error %Errorlevel%, in RestartPutty: putty.log could not be Deleted
        }
    }
  FileDelete, %projectfile%
  Run, %start%,,hide, %RemoteMachine%   ; start putty
  Sleep, 1000   ; wait for putty to start
  Loop   ; wait until we get a prompt.
    {
      Sleep, 100
      FileRead, retText, putty.log ; Read temp file to see if putty started.
      StringGetPos, pos, retText, %RemoteMachine%:~>,R
      If not Errorlevel  ; if we see a prompt
        {
          ControlSend,,%SetPath%{Enter}, %RemoteMachine%
          ; ControlSend,,{Enter}, %RemoteMachine%
          Sleep, 200
          FileRead, retText, putty.log
          IfNotInString, retText, %SetPath%
              MsgBox, 16, Error 2, Command: %SetPath% not Send.`n %retText%
          StringReplace retget, retText, %CommandPrompt% ,,A, UseErrorLevel    ; remove last command prompt
          If ErrorLevel = 0
              MsgBox, 16, Error 2, %CommandPrompt% not found `n %retText%.
          Break
        }
      If A_Index > 30
        {
          SoundPlay, Low.wav
          MsgBox,16, Error, putty would not Run.`n Please try again.
          ExitApp
        }
    }
  GuiControl, Focus, Input
  FormatTime, timestamp, %A_Now%, dd-MM-yy HH:mm:ss
  FileAppend, [%timestamp%] Putty restarted. LogSize: %PuttySize% k`n, %logfile%
  SoundPlay, Windows XP Ding.wav, wait
  ToolTip, Putty restarted size: %PuttySize% k
  SetTimer, HideToolTip, 5000
  WinActivate, PN
Return

Disable_All:
;GuiControl,, Remove,0  ; uncheck remove
;GuiControl, Disable, ListBox1
GuiControl, Disable, Edit1
GuiControl, Disable, OK
GuiControl, Disable, Project
GuiControl, Disable, Get
GuiControl, Disable, Put
GuiControl, Disable, Souterain
GuiControl, Disable, Receive
GuiControl, Disable, Remove
GuiControl, Disable, Up
GuiControl, Disable, Stop
GuiControl, Disable, Down
GuiControl, Disable, Carier
GuiControl, Disable, Go
GuiControl, Disable, Init
Disable_All = 1
return
  
Enable_All:
if Disable_All
{
GuiControl, Enable, ListBox1
GuiControl, Enable, Edit1
GuiControl, Enable, Edit2
GuiControl, Enable, OK
GuiControl, Enable, VSN
GuiControl, Enable, Project
GuiControl, Enable, Get
GuiControl, Enable, Put
GuiControl, Enable, Cancel
GuiControl, Enable, Souterain
GuiControl, Enable, Receive
GuiControl, Enable, Remove
GuiControl, Enable, Send
GuiControl, Enable, Up
GuiControl, Enable, Stop
GuiControl, Enable, Down
GuiControl, Enable, History
GuiControl, Enable, Carier
GuiControl, Enable, Go
GuiControl, Enable, Init
GuiControl, Enable, ComboBox1
GuiControl, Enable, Input
if Action = Get
  GuiControl, Focus, Input
  if Action = Put
  GuiControl, Focus, Input
  if Action = Confirm
  GuiControl, Focus, Input
  GuiControl, Enable, End
  Disable_All =
}
return
