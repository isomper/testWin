Local $iPID = Run("notepad.exe","", @SW_SHOWMAXIMIZED)
WinWaitActive("无标题 - 记事本","",2)
Send("这是一个测试脚本")
Send("{HOME}")
Sleep(200)
Send("+{TAB}")
Sleep(200)
Send("!F")
Sleep(1000)
Send("^N")
Sleep(1000)
Send("!S")
Send("test")
Send("{ENTER}")
Send("{TAB 4}")
Send("{ENTER}")

Sleep(2000)
ProcessClose($iPID)
