#include <WinAPIRes.au3>
#include <array.au3>

MouseClick("left",0,728 )
Sleep(500)
Send("{UP 3}")

Local $iPos = _WinAPI_GetCaretPos()

If IsArray($iPos) Then
	_ArrayDisplay($iPos)
Else
	MsgBox(0,"",$iPos)
EndIf

;MouseClick("right", $iPos[0],$iPos[1])


Sleep(2000)
Exit