#cs---------------------------------------------------------------------
功能说明：操作navicat客户端，连接后台，打开查询窗口，输入sql
查询语句并执行
#ce--------------------------------------------------------------------

;打开桌面上的navicat客户端
ShellExecute(@DesktopDir & "\Navicat for MySQL")
;激活打开的navicat窗口
WinWaitActive("Navicat for MySQL")
;按一下tab，然后按回车，连接mysql	后台,选中数据库
Send("{TAB}")
Send("{ENTER}")
Sleep(1000)
Send("{DOWN}")
Send("{ENTER}")

Sleep(1000)
;输入组合键ctrl+q打开查询窗口
Send("^q")
Send("^{SPACE}")
Sleep(2000)
;输入查询sql
Send("select * from fort_account;")
Sleep(2000)
;输入组合件ctrl+r执行sql
Send("^r")
Sleep(2000)
;关闭查询窗口
Send("^w")
Sleep(2000)
;关闭navicat
Send("!{F4}")
