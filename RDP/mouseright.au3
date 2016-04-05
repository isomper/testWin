#include <Array.au3>
#include <MsgBoxConstants.au3>
;Local [2] = [1,2]
;Local $downarr[2] = [3,6]
;SandAright("^{esc}",10,2,6,1)
;SandAright("^{esc}",10,2,6,1)
;clickstart("^{esc}",9,2)
;点击开始
;MandAright("right",320,468,1,1,7)
;secpolAright("C:\WINDOWS\system32\secpol.msc","本地安全设置",3,4)
MscAright("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,2,1,1,5,2)
;MscAright("C:\WINDOWS\system32\rrasmgmt.msc","路由和远程访问",0,0,0,0,4,1)
;clickstart("#e",2,0,5)
;调用clickstart和Aright
Func SandAright($key,$ntab,$nUp,$nDown,$naDown)
	clickstart($key,$ntab,$nUp,$nDown)
	Aright($naDown)
	;Bright($nbDown)
	Send("{ESC}")
EndFunc
;调用Mouseright和Aright
Func MandAright($action,$x,$y,$nClicks,$nDown,$naDown,$nbDown)
	Mouseright($action,$x,$y,$nClicks,$nDown)
	Aright($naDown)
	Send("{ESC}")
EndFunc


Func _Msc($address,$name,$ntab,$nDown,$nright,$nbDown,$ncDown)
	ShellExecute($address)
	Global $pid = WinWaitActive($name)
	Send("{tab "& $ntab &"}")
	Send("{down "& $nDown &"}")
	Send("{right "& $nright &"}")
	Send("{down "& $nbDown &"}")
	Sleep(1000)
	Send("{APPSKEY}")
	Send("{down "& $ncDown &"}")
	Sleep(1000)
	;Send("{down "& $nrDown &"}")
EndFunc
;调用_Msc和Aright
Func MscAright($address,$name,$ntab,$nDown,$nright,$nbDown,$ncDown,$naDown)
	_Msc($address,$name,$ntab,$nDown,$nright,$nbDown,$ncDown)
	Aright($naDown)
	;WinClose($pid)
	Send("!{f4}")
EndFunc
Func Mouseright($action,$x,$y,$nClicks,$nDown)
	MouseClick($action,$x,$y,$nClicks)
	Send("{DOWN  "& $nDown &"}")
	;down(1)
	;Aright(7)
EndFunc
Func clickstart($key,$ntab,$nUp,$nDown);$action,$x,$y,$nClicks,
	Send($key)
	Sleep(1000)
	Send("{tab "& $ntab &"}")
	Sleep(1000)
	Send("{UP  "& $nUp &"}")
	Sleep(1000)
	Send("{APPSKEY}")
	Sleep(1000)
	For $i = 1 To $nDown
		Send("{DOWN}")
		Sleep(1000)
	Next
	Send("{ESC}")
	Sleep(1000)
EndFunc
;按右键 shift+F10
Func Aright($naDown)
	Send("{right}")
	For $i =1 to $naDown
		Send("{down}")
		Sleep(1000)
	Next
	Send("{left}")
EndFunc
Func Bright($nbDown)
	Send("{right}")
	For $i =1 to $nbDown
		Send("{down}")
		Sleep(1000)
	Next
	Send("{left}")
EndFunc
