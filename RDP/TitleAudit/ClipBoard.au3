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
;粘贴图片
;up("C:\Users\guyaru.pc\Desktop\runsikulix -r C:\Users\guyaru.pc\Desktop\Call_up_remote_TXT.sikuli","测试.txt - 记事本",3,@DesktopDir & "\test1.docx","test1.docx - Word",2)
func clip_board($scmd,$stitle,$cCount,$dfile,$dtitle,$pCount,$state)
	If $state == 0 Then
		open_remote($scmd,$stitle,$cCount)
		open_local($dfile,$dtitle,$pCount)
	ElseIf $state == 1 Then
		open_local($dfile,$dtitle,$pCount)
		open_remote($scmd,$stitle,$cCount)
	EndIf
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;打开远程文件，复制内容到剪切板
parameters:
$scmd:打开远程文件命令                      $stitle:远程文件标题                       $cCount:按下几下down键到复制按钮
#ce-----------------------------------------------------------------------------------------------------------------------
Func open_remote($scmd,$stitle,$cCount)
	Run(@ComSpec & " /c " & $scmd,"", @SW_HIDE,"")
	Local $sPid = WinWaitActive("测试.txt - 记事本","",10)
	;ControlFocus($sPid,"","Edit1");使焦点落在编辑框上
	Send("+{right 20}")
	Sleep(1000)
	Send("{APPSKEY}")
	Sleep(1000)
	Send("{DOWN " & $cCount & "}")
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	;$text = ClipGet()
	;ConsoleWrite($text& "111111111111" & @CRLF)
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;粘贴剪切板内容到本地文件
parameters:
$dfile: 本地文件文件                         $dtitle:本地文件标题                        $pCount:按下几下down键到粘贴按钮
#ce-----------------------------------------------------------------------------------------------------------------------
Func open_local($dfile,$dtitle,$pCount)
	ShellExecute($dfile)
	WinWaitActive($dtitle,"",10)
	Sleep(1000)
	Send("{ENTER}")
	Sleep(1000)
	Send("{APPSKEY}")
	Sleep(1000)
	Send("{DOWN " & $pCount & "}")
	Send("{ENTER}")
EndFunc

