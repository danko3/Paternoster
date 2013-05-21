/*Includes:
PNGui.ahk
Paternoster.ahk
SerialPort.ahk (in Paternoster.ahk)
For testing only.
Can be used whithout the network and database.
For paternoster movement only.
*/
#include PNGui1.22_PN.ahk
#include Paternoster.ahk

IncrementalSearch: 
Return
ListBoxClick:
return
ListView:
  If A_GuiEvent = D ; user attempted to drag.
      Sleep,100
Return