;
; AutoHotkey Version: 1.0.48.05
; Language:       English
; Platform:       Win9x/NT/XP
; Author:         H. Tenkink <tenkink@jive.nl>
;
; Script Function:
;   The Paternoster is the JIVE diskpack transport and storage system.
;   It consists of 32 shelves with each 18 positions to store diskpacks.
;   It has two stations to get packs in and out. One in the Basement one in the Expedition.
;   This script is used to get diskpacks in and out of the paternoster and communicates
;   with a database on %RemoteMachine% to keep track of the positions of each pack in the Paternoster.
;   The commumication with the database runs via a Perlscript paternosterUpdate.pl.
;   This script lives on jop83/export/jive/jops/bin/tapes/user/.
;   The communication with the Perlscript goes through putty in the subroutine GetIt.
;   The logfile of putty is used to extract the answers from paternoster.pl.
;   Other required programs:
;   PN[version]2.ahk; putty.exe; PN.ppk; SerialPort.ahk --> path:
;   C:\Program Files\AutoHotkey\Extras\Scripts\
;   Includes:
;   PNGui.ahk       -- The graphical user interface.
;   SerialPort.ahk  -- The program that takes care of the RS232 protocol.
;   Paternoster.ahk -- The paternoster control program.
;   Other required files:
;   PN.ini (Paternoster carier layout, free cariers, laserpointer table, Paternoster command strings)

/*
paternoster.pl [-vsn [[ -in -shelf] | -out | -remove | -send | get ]] [-e exp] [-p]

Examples:
paternoster.pl -in     -v vsn -s shelf - insert vsn in Paternoster
paternoster.pl -out    -v vsn          - remove vsn from Paternoster set shelf = 00,00
paternoster.pl -remove -v vsn          - Remove vsn from Paternoster, set vsnstatus = removed, if shelf = 00,00
paternoster.pl -send   -v vsn          - Remove vsn from Paternoster, set vsnstatus = sent
paternoster.pl -get    -v vsn          - Get vsn and its position
paternoster.pl -e experiment           - Get all vsn's of experiment
paternoster.pl -p                      - Get all free positions separated by spaces.
paternoster.pl -experiments            - Get all the projects.

Errors:
Error 1  => in: While inserting vsn, shelf is not 00,00.
Error 2  => in: vsnstatus is Removed or vsn does not exist.
Error 3  => out: Shelf is already 00,00 or vsn does not exist.
Error 4  => remove: Vsnstatus is not Free or shelf is not 00,00 or vsn does not exist.
Error 5  => send: Vsnstatus is not 'removed' or vsn does exist.
Error 6  => get: vsn does not exist.
Error 7  => Experiment has no vsn's.
Warnings:
Get : No vsn was given.


Includes:
PNGui.ahk
Paternoster.ahk
SerialPort.ahk (in Paternoster.ahk)
*/

#Include, PNGui1.24.ahk
#Include, Paternoster1.ahk
#Include, SerialPort.ahk

Input:
  ;SetTimer, Listen, On
  GuiControl, +Default, Enter
  blip = 0
 ; MsgBox, Give input before: %CtrlContents%
  GuiControlGet, CtrlContents,, Input
 ; MsgBox, Give input after: %CtrlContents%
  StringLen, Len, CtrlContents
  If Len = 0
    {
      ;  MsgBox,,Input, VSN asked is %CtrlContents%, length = %Len%`n(VSN= %VSN% Action = %Action%) ; debug line
      ;FileAppend, [%timestamp%]  Input stringlength is 0`n, %logfile%
      CtrlContents = %VSN%   ;Put it in the Input. (VSN is the short VSN)
    }
  ;MsgBox,16,,Input=%CtrlContents% semaphore=%Semaphore%, Action= %Action%
  StringReplace, CtrlContents, CtrlContents, O, 0, All ; replace all O with 0.
 ; GuiControl, Disable, Input
  Gosub, InputCheck
  SB_SetText(CtrlContents, 1) ; display the corrected Read string in full.
  If not CtrlContents
    {
      GuiControl,, Input  ; clear the Inputfield.
      GuiControl, Enable, Input
      GuiControl, Focus, Input
      FileAppend, Input check failed. VSN read = %Badcontents%`n, %logfile%
      Return
    }

  If Action = Send
    {
      command = -Send -v %VSN%    ;Remove vsn from Paternoster, set vsnStatus = sent, Existence = History
      Gosub GetIt    ; the GetIt ThRead will get the buffer
      FormatTime, timestamp, %A_Now%, dd-MM-yy HH:mm:ss
      FileAppend, [%timestamp%] %CtrlContents% Sent`n, %logfile%
      GuiControl,, Input  ; rewrite the Input field. (Empty it.)
      EnvAdd, AmountSent, 1  ;add 1 to AmountSent
     ; IfEqual, AmountSent, 1
    ;  FileAppend, [%timestamp%]`n, %VSNShip%  ; Write a timestamp in the file VSNSend when starting.
      FileAppend, %CtrlContents%`n, %VSNShip%  ; Just put the VSN in a file for Track.
      SB_SetText(AmountSent, 3) ; display the Number of VSN's
      GuiControl,, OldVSN, %CtrlContents%  #%AmountSent% Sent||  ; Two pipes means it is preselected
      SoundGet, master_volume
      SoundSet, +50
      SoundPlay, click.wav
      SoundSet, %master_volume%
    }
  Else
    {
      StartTimefind := A_TickCount   ; timer start
      If Action <> Confirm
     ;     MsgBox, ctrlC %CtrlContents%, vsn= %VSN%
          Position := Find(VSN)   ; Goto subroutine 'Find'
      ;MsgBox, AfterFind Action = %Action% LastVSN = %LastVSN%, Position= %Position% VSN_Asked = %VSN_Asked%
      If Position  <> 00,00 ; The VSN is found.
          Gosub, Found
      Else
          Gosub, NotFound
    }
  ElapsedTimeGet := A_TickCount - StartTimefind  ; timer end

 ; GuiControl, Enable, Input
 ; GuiControl, Focus, Input
Return


NotFound:
  If Action = get
    {
      LastVSN = %VSN% NOT FOUND
      GuiControl,, Input, %LastVSN%
      SoundPlay, Low.wav, wait
      Sleep, 1000
      Confirmed = 1
      GuiControl,, OldVSN, %LastVSN% %Position%||  ; Two pipes means it is preselected. History window
      If Remove  ; if remove box in Gui is Checked.
        {
          ;MsgBox, Debug`nNotFound Remove
          command = -remove -v %VSN%   ;remove vsn from Paternoster, set vsnStatus = removed, if shelf = 00,00
          Gosub GetIt
          Act = Removed
          GuiControl,, OldVSN, %VSN_Asked% %Position% Removed||
        }
      FormatTime, timestamp, %A_Now%, dd-MM-yy HH:mm:ss
      FileAppend, [%timestamp%] %VSN% Not found %Act%`n, %logfile%
      GuiControl,, OldVSN, %LastVSN%||
    }
  Else If Action = Put
    {
      StartTimePut := A_TickCount   ; timer start
      Position := Put(VSN)    ; Now Goto 'Put'. redefine Position
      ElapsedTimePut := A_TickCount - StartTimeput  ; timer end
    }
  Else If Action = Project
    {
      GuiControl,, OldVSN, %CtrlContents% %Position% ||  ; Two pipes means it is preselected. History window
    }
  Else If Action = Confirm ; this should not happen ever.
    {
      MsgBox, Action is Confirm but not found!!!!!
      StringLeft, CtrlContents, CtrlContents, 8  ; compare shortVSN only
      SoundPlay, Low.wav,wait
      GuiControl,, OldVSN, Read: %CtrlContents%, wanted: %LastVSN% ||  ; Two pipes means it is preselected. History window
    }
  GuiControl,, Input  ; rewrite the Input field. (Empty it.)
  SoundPlay, c:\windows\media\Start.wav   ; tik!
  GuiControl, Enable, Input
  ;MsgBox, ctrlc = %CtrlContents% asked= %VSN_Asked% last = %LastVSN% action = %action%
Return

Found:
  ;MsgBox, Found! Input=%CtrlContents% Action= %Action%
  If Action = get
    {
      GuiControl,, Action, Confirm  ; Write the big letter Caption.
      Action = Confirm   ; After it is found it should be confirmed.
      ;MsgBox, Found ctrlc = %CtrlContents% asked= %VSN_Asked% LastVSN = %LastVSN%
      GuiControl,, OldVSN, %LastVSN% %Position% Found||  ; Two pipes means it is preselected.
    }
  Else If Action = Put
    {
      GuiControl,, Input, Present!
      SoundPlay, Low.wav, wait
      ;MsgBox, 48,, VSN= %LastVSN%, %Position%
      Sleep, 1000
      GuiControl,, OldVSN, %LastVSN% %Position% Present!||  ; Two pipes means it is preselected.
      StringRight, Pos, Position, 2  ; get laserposition only
      LP := LaserPointer(Pos)
    }
  Else If Action = Confirm
    {
      blip=1
      LongVSN = %CtrlContents%  ; Save longVSN for logbook
      StringLeft, CtrlContents, CtrlContents, 8  ; compare shortVSN only
      If (CtrlContents = LastVSN)  ; Confirmed!
        {
          SoundPlay, chimes.wav   ; tideling!
          ;ListVars
          ;Pause
          Position := Get(LastVSN)  ; Remove the VSN from database.
          LastVSN =       ;reset.
          GuiControl,, Action, Get
          Action =  Get
          Confirmed = 1  ; for Project.
        }
      Else
        {
          ;MsgBox, %CtrlContents% <> %LastVSN% --> not confirmed.
          SoundPlay, Low.wav,wait
          GuiControl,, OldVSN, Read: %CtrlContents%, wanted: %LastVSN% ||  ; Two pipes means it is preselected.
        }
    }
  GuiControl,, Input  ; rewrite the Input field. (Empty it.)
  GuiControl, Enable, Input
;MsgBox, ping
  SoundPlay, c:\windows\media\Start.wav
Return


Experiment:      ; Button Project clicked.
  ;MsgBox, Project= %action% input= %Input%
  If Action = Project
    {
      Gui, font, s12, Courier New Bold
      If Project =
        {
          If not Output
              Gosub, GetProjects
          GuiControl, Disable, Input
          GuiControl, show, edit6
          GuiControl, show, ListBox2
          GuiControl, -Default, Enter
          GuiControl, +Default, Ok
          GuiControl, Focus, edit6
        }
    }
Return


IncrementalSearch:      ;IncrementalSearch is defined in Edit field in Gui (Edit6).
  Gui Submit, NoHide
  len := StrLen(searchedString)   ; searchedString is defined in Edit field in Gui (Edit6).
  itemNb := 1
  Loop Parse, Output, |
    {
      StringLeft part, A_LoopField, len
      loopfield = %A_LoopField%
      If (part = searchedString)
        {
          itemNb := A_Index
          Break
        }
    }

  ToolTip %searchedString% (%itemNb%) ;%loopfield%
  SetTimer HideToolTip, 1000
  GuiControl Choose, choice, %itemNb%
Return

HideTooltip:
  SetTimer HideToolTip, Off
  ToolTip
Return

ListBoxClick:           ;ListBoxClick is defined in ListBox in Gui (ListBox2).
  Gui Submit, NoHide
  GuiControl, Enable, OK
  GuiControl, +Default, Enter
  GuiControl, -Default, Ok
  GuiControl, hide, Edit6
  GuiControl, hide, ListBox2
  GuiControl, , static1, %choice%
  command = e %choice%
  Gosub, GetIt
  IfInString, Output, Error
    {
      MsgBox,16,,Project: %RetText%
      FileAppend, [%timestamp%] Project Not found %choice%`n, %logfile%
      ExitApp, 2
    }
  IfInString, retClick, DBD ; Part of Error message from jop83
    {
      MsgBox,16,DB Error, %RetText%
      FileAppend, [%timestamp%] Database error %RetText%`n, %logfile%
      ExitApp,2
    }
  FileAppend, %RetText%, %projectfile%

  Loop, Read, %projectfile%
    {
      StringMid, VSNvar, A_LoopReadLine, 1, 8   ; Parse out VSN
      StringMid, VSNpos, A_LoopReadLine, 10, 5   ; Parse out position
      LV_Add("",VSNpos,VSNvar)
      TotalAmount = %A_Index%
    }
  LV_ModifyCol()  ; Auto-size each column to fit its contents.
  LV_ModifyCol(1, "Sort" "logical")
  If TotalAmount = 1
      ControlSetText, Edit1, %TotalAmount% Pack
  Else
      ControlSetText, Edit1, %TotalAmount% Packs
Return

ListView:
  If A_GuiEvent = D ; user attempted to drag.
      Sleep,100
Return
; So fall through to the next label.
ButtonOK:
  RowNumber = 0  ; This causes the First loop iteration to start the search at the top of the list.
  Amount = 0 ; reset the Counter

  Loop ; Just for Counting the Number of packs selected.
    {
      RowNumber := LV_GetNext(RowNumber)  ; Resume the search at the row after that found by the previous iteration.
      If not RowNumber  ; The above Returned zero, so there are no more selected rows.
          Break
      EnvAdd, Amount, 1  ;add 1 to Amount (Amount is the Number of selected VSN's)
      ; MsgBox 1#%RowNumber%, %VSN%, n: %Amount%
    }

  LV_Modify(1, "Vis")  ; Jump back to row one. (If scrolled down)
  Sleep,1000  ; Wait to show that it starts at the top.
  FormatTime, timestamp, %A_Now%, dd-MM-yy HH:mm:ss
  FileAppend, [%timestamp%] Project: %choice%`n, %logfile%
  RowNumber = 0  ; This causes the First loop iteration to start the search at the top of the list.
  Loop ; for handing out the packs.
    {
      Number := LV_GetNext(RowNumber) ; The selected row (Number) should be 1
      If not Amount  ; If the amount is zero, there are no more selected rows.
          Break
      If (Number > 1)  ; The selected row is preceeded by one or more not selected rows.
        {
          LV_Delete(1)   ; Delete the not selected Topmost VSN .
          Sleep,50   ; make the scrolling (deleting) Visible.
        }
      Else
        {
          LV_GetText(VSN, Number,2)  Get the VSN from projectlist
          If ErrorLevel
            {
              MsgBox, 21, Error Reading VSN from list
              IfMsgBox Retry
                  LV_GetText(VSN, Number,2)  Get the VSN from projectlist
              Else
                  Gosub ButtonCancel
            }
          GuiControl, Enable, Input
          ControlSetText, Edit2, %VSN%   ;Put it in the Input.
          SB_SetText(VSN, 1) ; display the VSN in the Status bar too.
          Sleep, 500
          ControlSetText, Edit1, %Amount% Packs
          EnvSub, Amount, 1  ;subtract 1 from Amount
          Sleep,500
          ;MsgBox, confirming. %VSN%
          LastVSN = %VSN%
          Confirmed =   ; reset Confirmed state.
          Canceled =  ; reset Canceled state. (set when Cancel Button Clicked)
          GuiControl, Disable, OK
          GuiControl, Disable, Project
          GuiControl, Disable, Get
          GuiControl, Disable, Put
          Action = Get
          GuiControl,, Action, Get
          GuiControlGet, CtrlContents,, Input
          If Errorlevel
              MsgBox, 16, Input Error, VSN=%VSN%, Ctrlc =%CtrlContents%
          Gosub, Input
          Loop  ; Wait for confirmation.
            {
              If (Confirmed)
                {
                  ;ListVars
                  ;Pause
                  Break
                }
              Sleep, 1600
              If (Canceled)
                {
                  ; MsgBox, Canceled `n %Canceled%
                  Break
                }
              GuiControl,, Action, Confirm  ; Write the big letter Caption.
            }
          LV_Delete(1)
        }
    }

  ; Display the window and return. The script will be notified whenever the user double clicks a row.
  ; Activate the disactivated buttons.
  GuiControl, Enable, vOK
  GuiControl, Enable, vProject
  GuiControl, Enable, vGet
  GuiControl, Enable, vPut
  SoundPlay, c:\windows\media\cash_register_x.wav  ; endtune
  Gosub, ButtonCancel
Return
;}

; #######################################################################
; ############## Function Find ##########################################
; #######################################################################
Find(VSN_Asked)
  {
    global
    command = get -v %VSN_Asked%    ;Get vsn and its position
    ask = %VSN_Asked%
    Gosub GetIt
    Position = 00,00  ; defaults to not found
    IfNotInString, retText, Error
    {
        StringSplit, Position, retText, %A_Space%   ; Split on Space.
        LastVSN = %Position1%
        StringLeft, Position2, Position2, 5  ; Take First 5 char to remove CR.
        Position = %Position2%  ; The second item is the position.
        StringSplit, Part, Position2,`,,`n  ; Split on comma, omit LF
        If Part1 Between 1 and 32
            Carier = %Part1%
        If action = Get
          {
            Answer := Paternoster(Carier,Station)   ; Position the paternoster.
            Laser := LaserPointer(Part2)  ; The second item is the position on the shelf.
          }
        ;MsgBox, Find Action is %Action%
    SB_SetText(Part1 "," Part2, 3) ; display the position.
    }
    Return, Position
  }

; #######################################################################
; ############## Function Get ###########################################
; #######################################################################
; confirmation, get a VSN from the paternoster.
Get(VSN_Asked)
  {
    global
    {
      command = out -v %VSN_Asked%   ;Get vsn from Paternoster set shelf = 00,00
      Gosub GetIt
      GuiControl,, OldVSN, %VSN_Asked% %Position% Taken||  ; Two pipes means it is preselected.
      Act = Get
    }
If Remove = 1   ; Remove box in Gui Checked
  {
    command = remove -v %VSN_Asked%   ;remove vsn from Paternoster, set vsnStatus = removed, if shelf = 00,00
    Gosub GetIt
    Act = Removed
    GuiControl,, OldVSN, %VSN_Asked% %Position% Removed||
  }
FormatTime, timestamp, %A_Now%, dd-MM-yy HH:mm:ss
FileAppend, [%timestamp%] %LongVSN% %Position% %Act% %ElapsedTimeGet% ms`n, %logfile%
SB_SetText(Position, 3) ; display the position.
SB_SetText("",6) ; clear display free places.
Return, Position
}

; #######################################################################
; ############## Function Put ###########################################
; #######################################################################

Put(VSN_Asked)
  {
    global
    command = p   ;Get all free positions separated by Spaces.
    Gosub GetIt
    Loop    ; Read PN.ini
      {
        IniRead, FreeCarier, PN.ini, FreeCariers, FreeCarier%A_Index%
        ;MsgBox, %Freecarier%
        If FreeCarier = Error   ; End of inifile section reached.
            Break
        cariers = %Cariers%`n%FreeCarier%
        NumberofCariers = %A_Index%
      }
    StringSplit, Output, retText, %A_Space%   ; Split on Space
    FreePlaces := Output0 - ((32 - NumberofCariers) * 18)-2  ; 2 for start and end?
    ;total number of cariers is 32   There are 18 positions per carier for diskpacks
    Loop, %Output0% ; (Output0 is the number of elements given by StringSplit)
      {
        StringSplit, Part, Output%A_Index%, `,   ; Split on comma.
        Carier = %Part1%
        LP     = %Part2% ; The second item is the position on the shelf.
        SB_SetText(Carier "," LP, 3) ; display the position.
        SB_SetText(FreePlaces " free", 6) ; display free places.
        IfInString, cariers, %carier%
          {
            ; IfInString, Semaphore, Initialized  ; Semaphore should be Ready or intialized
            Semaphore = Ready
            IfInString, Semaphore, Ready
              {
                command = in -v %VSN_Asked% -s %Carier%,%LP%   ; update database.
                Gosub GetIt
                Act = Put
                If Receive = 1
                    Act = Received
                ;GuiControl,, VSN, %VSN_Asked% ; rewrite the edit field.
                FormatTime, timestamp, %A_Now%, dd-MM-yy HH:mm:ss
                FileAppend, [%timestamp%] %VSN_Asked% %Carier%`,%LP% %Act% %ElapsedTimePut% ms`n, %logfile%
                GuiControl,, OldVSN, %VSN_Asked% %Carier%,%LP% Placed||  ; Two pipes means it is preselected.
                Answer := Paternoster(Carier,Station)   ; Position the paternoster.
                Laser := LaserPointer(LP)
              }
            Else
            {
               SoundPlay, Low.wav, wait
               GuiControl,, OldVSN, %VSN_Asked% Not placed||
               gui, +owndialogs
               MsgBox, 0x4030, Paternoster full, No empty places left in the paternoster.
            }
            Return, Position
          }
      }
  }

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
;:::::: GetProjects :::::::::::::::::::::::::::::::::::::::::::::::::::::
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
GetProjects:
  command = experiments   ;Get all the projects.
  Gosub GetIt
  StringReplace, Output, retText, %A_Space%, |, All
  StringReplace, Output, Output,||,, All
  GuiControl,, choice, %Output%
Return

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
;::::::  Subroutine GetIt  ::::::::::::::::::::::::::::::::::::::::::::::
;::::::  This subroutine communicates with jop83 through putty.  ::::::::
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
GetIt:
  ;  MsgBox,,Command: %command%
  succes = 0
  mail = You have new mail.
  StartTimeGetIt := A_TickCount   ; timer start
  ; First check if putty is (still) running.
  Process, Exist, putty.exe
  If not Errorlevel  ; Putty is not Running.
      Gosub RestartPutty
  ; Now send the string to the (hidden) putty window.
 ; ControlSend,,{Raw}paternosterUpdate.pl -%command%, %RemoteMachine%
 ControlSend,,{Raw}PN -%command%, %RemoteMachine%
  ControlSend,,{Enter}, %RemoteMachine%  
  While not success
  {
    Loop
      {
        Sleep, 50 ; wait for putty to finish
        FileRead, retText, putty.log
        If ErrorLevel
            MsgBox, 16, Error in GetIt, Error %ErrorLevel% in Reading Putty.log
        ; Here we Parse the file putty.log
        StringGetPos, pos, retText, PN -%Command%,R1 ; reverse search for last command
        StringTrimLeft, retText, retText, %pos%  ; remove all text up to the last command
        ;FileAppend, [retText0]`n %retText%, RetText.txt
        StringReplace retText, retText, PN -%Command% ; remove last command itself
        ;FileAppend, [retText1]`n %retText%, RetText.txt
        StringReplace retText, retText,`r ; remove empty line
        StringReplace retText, retText,`r`n ; remove {CRLF}
        ;FileAppend, [retText2]`n %retText%, RetText.txt
        StringReplace retText, retText,`r`n%CommandPrompt%  ; remove the prompt
        ;FileAppend, [retText3]`n %retText%, RetText.txt
        StringReplace retText, retText,`r`n%mail%
        ;FileAppend, [%timestamp%]`n %retText%, RetText.txt
        StringLen, len, retText  ; This is the answer from RemoteMachine
        IfInString, retText, )?
          { 
            ControlSend,,y{Enter}, %RemoteMachine%  ;on eg. CORRECT>PN -e em081a (y|n|e|a)?
            StringGetPos, pos, retText, yes,R1 ; reverse search for last command
            StringTrimLeft, retText, retText, %pos%  ; remove all text up to the last command
            StringReplace retText, retText, yes ; remove last command itself
            StringReplace retText, retText,`r ; remove empty line
            FileAppend, typed 'y' `n, %logfile%
            FileAppend, [%timestamp%]`n %retText%, RetText.txt
            succes = 1
            Break
          } 
          IfInString, retText, Error
                 {
                   SoundPlay, Low.wav, wait
                   MsgBox,16, String 'Error' found, %retText%
                   FileAppend, GetIt %retText%`n, %logfile%
                   succes = 1
                   break
                 }
        Loops = %A_Index%
        If len > 13  ; putty produced some output after removing crap.
          {
            succes = 1
            Break
          }
        If A_Index > 30  ; We can't wait forever.
          {
            MsgBox, 16, GetIt Error %RemoteMachine% slow, no response from %RemoteMachine% after command -%command%`n pos= %pos%, len = %len% `n%retText%
            FileCopy, putty.log, puttylog.bak, 1  ; overwrite last backup
            ;Gosub ButtonCancel
            ;break
            ExitApp
          }
        FileCopy, putty.log, puttylog.bak, 1  ; overwrite last backup
      }
    ElapsedTimeGetIt := A_TickCount - StartTimeGetIt  ; timer end
    FileAppend, [%timestamp%] GetIt took %ElapsedTimeGetIt% ms for command:{%command%}`n, %logfile%
    Return
  }
Return

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
;:::::: Subroutine ClosePN ::::::::::::::::::::::::::::::::::::::::::::::
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ClosePN:
  FormatTime, timestamp, %A_Now%, dd-MM-yy HH:mm:ss
  FileAppend, [%timestamp%]   > - Stop PN %PNversion% - <`n, %logfile%
  ; IniWrite, Value, Filename, Section, Key
  IniWrite, %Carier%, PN.ini, Carier, Carier
  ; Clean up temporary files
  FileDelete, %DebugFile%
  FileRecycle, %projectfile%
  FileRecycle, %project%
Return

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
;:::::: Subroutine InputCheck::::::::::::::::::::::::::::::::::::::::::::
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
InputCheck:
  ; The inputstring has 8 characters or more, but then the 8th char is a /.
  ; In that case we split the string the first part being the 'short' VSN and the rest.
  ; If the input is bad CtrlContents is made empty.
  ; vsn =
  StringLen, Len, CtrlContents
  If Len > 18
      StringLeft, CtrlContents, CtrlContents, 18
  If Len = 0
    {
      Gosub, ButtonCancel
      GuiControl,, Input  ; rewrite the Input field. (Empty it.)
      Return
    }
  If Len < 8
    {
      SoundPlay, Low.wav, wait
      VSN = %CtrlContents%
      SB_SetText("2Short: " VSN, 2)
      SB_SetIcon("surprised.ico",1,2)
      Badcontents = %CtrlContents%
      CtrlContents =
      GuiControl,, OldVSN, %VSN% (short VSN) ||  ; Two pipes means it is preselected.
Return
}
; the string CtrlContents is longer than 8 char so is an extended VSN so has two /
StringSplit, OutputArray, CtrlContents, /
If (OutputArray0 = 1)   ; The string could not be split.(Contains no /)
  {
    If Len = 8
      {
        IfInString, CtrlContents, -
            VSN = %CtrlContents%
        IfInString, CtrlContents, +
            VSN = %CtrlContents%
        If VSN
            Gosub VSN_OK
        Else
          {
            SoundPlay, Low.wav, wait
            SB_SetText("no '+' or  '-' in VSN ", 2)
            SB_SetIcon("surprised.ico",1,2)
            Badcontents = %CtrlContents%
            CtrlContents =
            Return
          }
      }
    Else
      {
        SoundPlay, Low.wav, wait
        SB_SetText("no '/' in VSN ", 2)
        SB_SetIcon("surprised.ico",1,2)
        Badcontents = %CtrlContents%
        CtrlContents =
        Return
      }
  }
Else
  {
    StringGetPos, Pos, CtrlContents, /
    If Pos <> 8
      {
        SoundPlay, Low.wav, wait
        SB_SetText("Bad / " CtrlContents, 2)
        SB_SetIcon("surprised.ico",1,2)
        Badcontents = %CtrlContents%
        CtrlContents =
        Return
      }
    Else
        VSN := Outputarray1
  }

VSN_OK:
if receive
  FileAppend, %CtrlContents%`n, %VSNreceive%  ; Save VSN for use in Track.
  StringLeft, VSN, VSN, 8  ; compare shortVSN only
  GuiControl,, Input, %VSN%
  SB_SetText(" " VSN, 2)
  SB_SetIcon("lol.ico",1,2)
Return

