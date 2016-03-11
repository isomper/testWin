ShellExecute(@DesktopDir & "\1.txt")
WinWaitActive("1.txt - 记事本")

Opt("SendCapslockMode", 0)		;设定CAPSLOCK带入记事本

Send("{CAPSLOCK on}")				;打开capslock变为大写状态
Sleep(1000)
Send("+a")									;输出a
Send("{CAPSLOCK off}")				;关闭capslock变为小写状态
Sleep(1000)
Send("+a")									;输出A