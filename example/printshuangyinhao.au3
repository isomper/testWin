ShellExecute(@DesktopDir & "\1.txt")
WinWaitActive("1.txt - 记事本")
Send('123')