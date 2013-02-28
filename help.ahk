
OnMessage(0x53, "WM_HELP")
Gui, Add, Button, ghelp, help
Gui, Show, w300 h300, Gui 
return

help:
Gui, +Owndialogs
MsgBox, 0x4000, help 0, gui# %A_Gui%
return


WM_HELP()
{
MsgBox, 0x4000,,Here is your help %A_Gui%
}
return

GuiEscape:
GuiClose:
ExitApp
