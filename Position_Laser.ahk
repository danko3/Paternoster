;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         H. Tenkink <tenkink@jive.nl>
;
; Script Function:
; See the red book to get the laserpointer working at all.
; First use QPOS.EXE then QCAL.EXE ( i:\paternoster\L20\ )

;   Testing the paternoster laserpointer.
;   The laserpointer is steered by the slider position.
;   When the laserpointer tabel is used the slider moves in 230 discrete steps.
;   Each step is translated to a LaserPointer step becouse the laserpointer stepping motor
;   has less steps then the 340 values we can effectivly sent to the Laserpointer.
;   To steer the pointer without the table uncheck the 'table'
;   This might be nessesary to recreate the laserpointer table.
;   The gui darkens, the slider is rescaled to 340 steps, the laserpointer table is not needed anymore.


#SingleInstance force ; If it is alReady Running it will be restarted.
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode,2 ; string somewhere in titel is ok.
;SetTitleMatchMode,slow  ; makes window recognition more reliable

/*
pos = position in the table
Position = tabel result
*/

; Check dependencies
if not FileExist "SerialPort.ahk"
      {
        MsgBox, 16, Error, The script SerialPort.ahk cannot be found!
        SoundPlay, Low.wav, wait
      }

; Set some parameters
Station = 1 ;Souterain

RS232_Port     = COM2
RS232_Baud     = 1200
Gosub, InitSerialPort   ; Initialize the laser pointer
LP_Port        = %RS232_FileHandle%

Gui, Add, Text, x45 y154 w5 h20 , 1
Gui, Add, Text, x785 y154 w22 h20 , 230
Gui, Font, S40 CDefault, Helvetica Bold
Gui, Add, Text, x380 y29 w120 h60 vPos
Gui, Font, S10 CDefault, Console Bold
Gui, Add, Button, x155 y9 w100 h40 , Off
Gui, Add, Checkbox, x730 y9 w100 h40 checked gtable vtable, Use Table
Gui, Font, S10 CDefault, Console Bold
Gui, Add, Button, x35 y9 w100 h40 vSouterain, Souterain
;Gui, Add, Picture, w300 h-1, C:\My Pictures\Company Logo.gif
Gui, Add, Slider, x35 y120 w770 h30 gProgress vSlider range1-230 AltSubmit ToolTip
; Generated using SmartGUI Creator 4.0
Gui, Show, w830 h190, Laserpointer



Step = 1
Pos = 1
Return

GuiClose:
ExitApp

ButtonOff:
Pos = Off
LP := LaserPointer(Pos)
Pos = 1
Position = 1
Gui, Font, S40 CDefault, Helvetica Bold
GuiControl, Font, Pos
GuiControl,, Pos, %Pos%
GuiControl,, Slider, 1
return

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
  GuiControl,, expedition ;x690 y102 w130 h67
Return

Table:
  GuiControlGet, CtrlContents,, table
  ;MsgBox, table= %CtrlContents%
  If CtrlContents
    {
      Gui, Font, s9 CDefault
      GuiControl, Font, static2
      GuiControl, , static2, 230
      GuiControl, +range1-230, slider
      Gui, color, dddddd
      if not FileExist "LPtabel.ini"
      {
        MsgBox, 16, Error, Laserpointer table (LPtable.ini) cannot be found!
        SoundPlay, Low.wav, wait
      }
    }
  Else
    {
      Gui, Font, s9 CRed
      GuiControl, Font, static2
      GuiControl, , static2, 340
      GuiControl, +range1-340, slider
      Gui, color, bbbbbb
    }
    GuiControl, focus, slider
    
Return

Progress:
  Gui Submit, NoHide
  ;msgbox, table= %table%
  If table
    {
      
      IniRead, Position, LPtabel.ini, Positions%Station%, Pos%Slider%
      if (position = "ERROR")
      {
        MsgBox, 16, Error, Laserpointer table (LPtable.ini) corrupt!
        SoundPlay, Low.wav, wait
      }
    }
  Else
    {
      Position = %Slider%
    }
  Hex := ASCII_to_Hex1(Position)
  ;MsgBox, %Slider%, %Position%
  ;GuiControl,, Position, %Position%
  Gui, Font, S40 CRed, Helvetica Bold
  GuiControl, Font, Pos
  GuiControl,, Pos, %Slider%
  LP := LaserPointer(Hex)

  If (A_GuiEvent = 1)
    {
      If Slider = 230
          SoundPlay, Low.wav, wait
      Step++
    }
  Else If A_GuiEvent = 0
    {
      If pos <= 1
        {
          Pos++
          SoundPlay, Low.wav, wait
        }
      Pos -= %StepSize%
      Step--
    }
Return

;########################################################################
;###### ASCII to 'HEX' conversion  ########################################
;########################################################################
ASCII_to_Hex1(Position)
  {
    Read_Data_Num_Bytes := StrLen(Position)
    if (Read_Data_Num_Bytes = 1)
      Hex = 0x30,
    loop , %Read_Data_Num_Bytes%
      {
        StringLeft, Byte1, Position, 1 ; Read left char.
		StringTrimLeft, Position, Position, 1
        Hex = %Hex%0x3%Byte1%,
      }
    
	StringTrimRight, Hex, Hex, 1  ; remove the last comma
	;MsgBox,position: %Position%: %Hex%
    Return %Hex%
  }
  
 ; #######################################################################
; ############# Function Laser pointer ##################################
; #######################################################################

LaserPointer(Pos)
  {
    global
    LPon = 0x1B,0xFF,0x3%Station%,0x45,
    LPoff = 0x1B,0xFF,0x3%Station%,0x44
    EndofString = ,0x0D,0x0C,0x0D
    ;MsgBox, Laser: Station = %Station%

    If Pos = Off    ; The Laserpointer should be switched off.
      {
        ;MsgBox, LP: %Hexpos% `n %LP_Port%
        Hexpos = %LPoff%%EndofString%
        RS232_Write(LP_Port,HexPos)
      }
    Else
      {
        Hexpos = %LPOn%%Pos%%EndofString%
        RS232_Write(LP_Port,HexPos)
        ;MsgBox, Hex=%pos%
      }
    ;MsgBox, LaserPointer %LP_Port% `n Hex=%Hexpos%
    Return  ; %Pos%
  }
  #Include, SerialPort.ahk