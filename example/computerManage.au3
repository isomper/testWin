;#AutoIt3Wrapper_Res_RequestedExecutionLevel=requireAdministrator
#include <Process.au3>
#cs------------------------------------------------------------
功能说明：打开计算机管理并对计算机管理窗口进行操作
#ce------------------------------------------------------------
;BlockInput(1)											;禁止用户按键盘和鼠标
;_RunDos("compmgmt.msc" )
Local $pid = ShellExecute("C:\Windows\System32\compmgmt.msc")
;Opt("WinTitleMatchMode", 2)				;windows标题匹配规则
;ConsoleWrite($pid & @CRLF)

WinWaitActive("计算机管理")
;Sleep(500)
Send("{down}")
Sleep(500)
Send("{LEFT}")
Sleep(500)
Send("{down}")
Sleep(500)
Send("{tab}")
Sleep(500)
;ConsoleWrite(1 & @CRLF)
;Send("{tab}")
;ConsoleWrite(2 & @CRLF)
;BlockInput(0)			;允许用户输入鼠标和键盘
Sleep(1000)
ProcessClose($pid)