#include <Constants.au3>
#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/05/25
作者：顾亚茹
功能说明：本地文件（word、excel、txt）与远程文件（word、excel、txt）文件之间相互复制、粘贴、剪切操作
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------



;粘贴文字
;clip_board("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\txt.sikuli","wenzi.txt - 记事本",3,@DesktopDir & "\wenben.txt","wenben.txt - 记事本",4,0)

;clip_board("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\remoatopentxt.sikuli","a.doc - Word",3,@DesktopDir & "\新建 文本文档 (4).txt","新建 文本文档 (4).txt - 记事本",4,0)


Func clip_board_down($scmd,$dfile,$dtitle,$pCount,$state)
	If $state == 0 Then
		open_remote($scmd)
		Sleep(9000)
		open_excel($dfile,$dtitle,$pCount)
	ElseIf $state == 1 Then
		open_remote($scmd)
		Sleep(9000)
		open_local($dfile,$dtitle,$pCount)
	ElseIf $state == 2 Then
		open_remote($scmd)
		Sleep(15000)
		open_tupian($dfile,$dtitle,$pCount)
	EndIf
EndFunc

Func clip_board_up($dfile,$dtitle,$pCount,$scmd,$state)
	If $state == 0 Then
		open_excel($dfile,$dtitle,$pCount)
		open_remote($scmd)
	ElseIf $state == 1 Then
		open_local($dfile,$dtitle,$pCount)
		open_remote($scmd)
	EndIf
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;打开远程文件，复制内容到剪切板
parameters:
$scmd:打开远程文件命令
#ce-----------------------------------------------------------------------------------------------------------------------
Func open_remote($scmd)
	Run(@ComSpec & " /c " & $scmd,"", @SW_HIDE,"")
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
;粘贴剪切板内容到本地文件
parameters:
$dfile: 本地文件文件                         $dtitle:本地文件标题                        $pCount:按下几下down键到粘贴按钮
#ce-----------------------------------------------------------------------------------------------------------------------

Func open_local($dfile,$dtitle,$pCount)
	ShellExecute($dfile)
	Local $spid = WinWaitActive($dtitle,"",10)
	Send("{ENTER}")
	Sleep(500)
	Send("+{right 40}")
	Sleep(500)
	Send("{APPSKEY}")
	Sleep(500)
	For $i = 0 To $pCount-1
		Send("{DOWN}")
		Sleep(500)
	Next
	Send("{ENTER}")
	Sleep(500)
	WinClose($spid)
	Send("{TAB}")
	Sleep(500)
	Send("{enter}")
EndFunc

Func open_tupian($dfile,$dtitle,$pCount)
	ShellExecute($dfile)
	Local $spid = WinWaitActive($dtitle,"",10)
	Send("{ENTER}")
	Sleep(500)
	Send("+{right 40}")
	Sleep(500)
	Send("{APPSKEY}")
	Sleep(500)
	For $i = 0 To $pCount-1
		Send("{DOWN}")
		Sleep(500)
	Next
	Send("{ENTER}")
	Sleep(15000)
	WinClose($spid)
	Send("{TAB}")
	Sleep(500)
	Send("{enter}")
EndFunc

Func open_excel($dfile,$dtitle,$pCount)
	ShellExecute($dfile)
	Local $spid = WinWaitActive($dtitle,"",10)
	Send("{ENTER}")
	Sleep(2000)
	Send("+{right}")
	Sleep(2000)
	Sleep(2000)
	Send("{APPSKEY}")
	Sleep(2000)
	For $i = 0 To $pCount-1
		Send("{DOWN}")
		Sleep(500)
	Next
	Send("{ENTER}")
	Sleep(1000)
	WinClose($spid)
EndFunc


