#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/11/22
作者：张利娟，陈圆圆
功能说明：DAC测试动作流
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------


#include <Process.au3>
for $i=0 to 3

;操作1.txt文件
ShellExecute(@DesktopDir & "\1.txt")
WinWaitActive("1.txt")
Sleep(500)
Send("""This is some text.This is some text.This is some text.This is sometext.This is some text.This is some text.This is some text.This is some text.This is some text.This is some text.""")
Sleep(500)
WinClose("1.txt")
Sleep(500)
Send("!n")

;操作2.docx文件
ShellExecute(@DesktopDir & "\2.docx")
WinWaitActive("2.docx")
Sleep(500)
Send("""This is some text.This is some text.This is some text.This is sometext.This is some text.This is some text.This is some text.This is some text.This is some text.This is some text.""")
Sleep(500)
WinClose("2.docx")
Sleep(500)
Send("!n")

;打开浏览器并对浏览器进行操作
Local $pie =  ShellExecute("C:\Program Files\Internet Explorer\iexplore.exe")
Sleep(500)
Send("AutoIT")
Sleep(500)
Send("{ENTER}")
Sleep(1000)
ProcessClose($pie)

;打开计算机管理并对计算机管理窗口进行操作
Local $pid = ShellExecute("C:\Windows\System32\compmgmt.msc")
WinWaitActive("计算机管理")
Sleep(500)
Send("{down}")
Sleep(500)
Send("{LEFT}")
Sleep(500)
Send("{down}")
Sleep(500)
Send("{tab}")
Sleep(500)
Sleep(1000)
ProcessClose($pid)

;创建文本文件
Run("notepad.exe")
WinWaitActive("无标题 - 记事本")
Send("chenyuanyuan")
WinClose("无标题 - 记事本")
WinWaitActive("记事本")
Send("!n")
Next