#include <Process.au3>

#cs------------------------------------------------------------
功能说明：打开计算机管理并对计算机管理窗口进行操作
#ce------------------------------------------------------------
_RunDos("compmgmt.msc" )
Sleep(500)

WinWaitActive("计算机管理")
Sleep(500)
;Send("^F")
;Sleep(500)
Send("{DOWN}")


Sleep(5000)
WinClose("计算机管理")