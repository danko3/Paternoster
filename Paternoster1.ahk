; --------  Paternoster Functions  & Subroutines  --------
;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT/XP
; Author:         H. Tenkink <tenkink@jive.nl>
;
; Script Function:
;
; Control of the paternoster and laserpointer.
;  Paternoster  com4 9600 boud
;  Laserpointer com2 1200 boud

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
;:::::: Subroutine Listen :::::::::::::::::::::::::::::::::::::::::::::::
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

; Listen is scheduled every 0.5 seconds in PNGui.ahk

Listen:
/*
Hex codes:
0x02 = start of message
0x03 = end of message
*/

 ; Critical, On
  Read_Data := RS232_Read(PN_Port,"0xFF",RS232_Bytes_Received)
  If (RS232_Bytes_Received > 0)
    {
  ;  Critical, Off
      Read_Data_Num_Bytes := StrLen(Read_Data) / 2 ;RS232_Read() Returns 2 characters for each byte
      ;MsgBox, Read_Data = %Read_Data% `n (%Read_Data_Num_Bytes% Bytes)
      StringGetPos, Start_of_Answer, Read_Data, %SearchString%  ;(look for '02')
      If Start_of_Answer >= 0
        {
          EnvAdd, Start_of_Answer, 21 ; Look @ position after start of answer (byte 10)
          StringMid, PNCarier, Read_Data, 13, 4   ;(look for 2bytes after byte 6)
          StringMid, TestString, Read_Data, %Start_of_Answer%, 4 ;(look for 2bytes after byte 10)
          ReturnedCarier := Hex2ASCII(PNCarier)
          IniRead, Semaphore, PN.ini, TestString, %TestString% ; Read Teststring from PN.ini
          if Semaphore = ERROR
            Semaphore =
          /*
          If Carier
            {
              If (Carier <> ReturnedCarier)
              {
                  Answer := Paternoster(Carier,Station)
                  sleep,5000 ; don't keep hammering
                  ;SetTimer, Listen, Off
                  FileAppend, [%timestamp%]  %Carier% <> %ReturnedCarier%`n, %logfile%
                }
              Else If Action <> Get
                {
                  
                  If not blip
                    {
                      SoundPlay, Blip.wav, wait
                      
                    }
                }
                GuiControl,, Input  ; rewrite the Input field. (Empty it.)
            }
            */
            If not blip
            SoundPlay, Blip.wav, wait
            blip = 1
        }
    }

  Message := Hex2ASCII(TestString)
  ;control:=ahk_id %ControlHwnd%
  ;MsgBox,Control= c: %Control%, CC= %CurrentCarier%, RC= %ReturnedCarier%
  if Semaphore <> Ready
  Gui, Font, S15 Cred, Verdana
  else
  {
  Gui, Font, S15 C4444ff, Verdana ; Big blue font
  
}
  GuiControl, Font, Edit4 ; use font for system messages. (edit4)
  ControlSetText, Edit5, %ReturnedCarier%, PN ; system message (blue)
   
  ControlSetText, Edit4, %Semaphore%, PN  ; (PN is the window title.) system message
  if (Carier <> ReturnedCarier)
    if Carier
      GuiControl,, Carier, %Carier%   ;Update the Input.
    else
    {
      GuiControl,, ReturnedCarier, %Carier%   ;Update the Input.
      ;SoundPlay, Blip.wav, wait
     ; blip = 1
    }
Return  ;, %Semaphore%

; #######################################################################
; ############# Function PaterNoster ####################################
; #######################################################################
Paternoster(Carier,Station)
  {
    global
    ;MsgBox, PN Carier: %CurrentCarier% Station= %Station%
    Hex := ASCII_to_Hex(Carier)
    Stat = 0x3%Station%,
    HexPos = %Pre%%Hex%%Post%%Stat%%End%
    RS232_Write(PN_Port,Hexpos)  ; Present the desired carier.
    Return
  }


;-------------------------------------------------------------------------

; #######################################################################
; ############# Function Laser pointer ##################################
; #######################################################################

LaserPointer(Pos)
  {
    global
    LPon = 0x1B,0xFF,0x3%Station%,0x45,
    LPoff = 0x1B,0xFF,0x3%Station%,0x44,
    EndofString = ,0x0D,0x0C,0x0D
    ;MsgBox, Laser: Station = %Station%, CurrentCarier= %CurrentCarier%

    If Pos = Off    ; The Laserpointer should be switched off.
      {
        ;MsgBox, LP: %Hexpos% `n %LP_Port%
        Hexpos = %LPoff%%EndofString%
        RS232_Write(LP_Port,HexPos)
      }
    Else
      {
        IniRead, Pos, PN.ini, Laserpointer, Box%Pos% ; Read Laserpointercodes from PN.ini
        ;MsgBox, LaserPos %Pos%
        Hexpos = %LPOn%%Pos%%EndofString%
        RS232_Write(LP_Port,HexPos)
      }
    ;MsgBox, LaserPointer %Pos% `n Hex=%Hexpos%
    Return, %Pos%
  }
  #Include, SerialPort.ahk