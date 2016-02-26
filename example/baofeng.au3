#include <Process.au3>
#cs---------------------------------------------------------------------
功能说明：打开暴风影音播放电影，按P，按回车全屏播放10分钟
#ce--------------------------------------------------------------------
ShellExecute(@DesktopDir & "\诸神之战_bd.mp4")		;打开桌面上的诸神之战_bd.mp4

WinWaitActive("诸神之战_bd.mp4")								;锁定暴风影音窗口

Send("{ENTER}")																;按回车键使暴风影音全屏

Local $iTime = 1 * 60000												;定时1分钟
;ConsoleWrite($iTime & @CRLF)
Sleep($iTime)

Send("{ENTER}")																;按回车键使暴风影音退出全屏

Send("!{F4}")																	;按ALT + F4强制关闭暴风影音的窗口

Sleep(2000)

_RunDos("shutdown -l")													;注销远程会话




