#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/03/22
作者：顾亚茹
功能说明：模拟键盘（大小写字母、数字、符号、功能键、组合键）输入
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

;打开桌面上的1.txt文件
ShellExecute(@DesktopDir & "\1.txt")
Local $pid = WinWaitActive("1.txt - 记事本","",2)
;设定capslock带入记事本
Opt("SendCapslockMode",0)

#cs-----------------------------------------------------------------------------------------------------------------------
定义数据源
#ce-----------------------------------------------------------------------------------------------------------------------

;定义F1功能键
Local $key_f1[1] = ["F1"]
;定义F3功能键
Local $key_f3[1] = ["F3"]
;定义alt+tab快捷键
Local $key_with_alt[1] = ["TAB"]
;定义win+(d,e,r)快捷键
Local $key_with_win[3] = ["d","e","r"]
;定义数据源定义0~9数字
Local $key_num[10] = [0,1,2,3,4,5,6,7,8,9]
;定义ctrl+(a,x,c,v,s)快捷键
Local $key_with_ctrl[5] = ["a","x","v","c"]
;定义ctrl+(f,o,g,h,p,n)快捷键
Local $key_with_ctrl_sysbol[5] = ["f","o","g","h","p"]
;定义F系列功能键
Local $key_fn[12] = ["F2","F4","F5","F6","F7","F8","F9","F10","F11","F12"]
;定义`[]\;',./~{}|:<>?"符号以及空格和回车
Local $key_sysbol[23] = ["`","[","]","\",";","'",",",".","/","~","_","-","+","=","{","}",":","<",">","?",'"',"SPACE","ENTER"]
;定义26个字母
Local $key_letter[26] = ["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"]

#cs----------------------------------------------------------------------------------------------------------------------------
调用函数
#ce----------------------------------------------------------------------------------------------------------------------------

;输入0~9
key_singal($key_num,"",0)
;输入!@#$%^&*()符号
key_singal($key_num,"+",0)
;输入符号（`[]\;',./~{}|:<>?"空格回车）
key_singal($key_sysbol,"",0)
;输入26个小写字母
key_singal($key_letter,"",0)
;输入26个大写字母
key_singal($key_letter,"",1)
;开启大写键，按下shift输入26个小写字母
key_singal($key_letter,"+",1)
;模拟ctrl+{s,x,c,v,a}快捷键
key_singal($key_with_ctrl,"^",0)
;模拟ctrl+(f,o,g,h,p,n)快捷键
key_with_sysbol($key_with_ctrl_sysbol,"^")
;输入F系列快捷键
key_singal($key_fn,"",0)
;模拟alt+tab快捷键
key_singal($key_with_alt,"!",0)
;模拟win+(r,d,e)快捷键
win_key($key_with_win)
;模拟F1键
key_with_sysbol($key_f1,"F1")
;模拟F3键
key_with_sysbol($key_f1,"F3")
WinClose($pid)
Sleep(500)
Send("{TAb}")
Sleep(500)
Send("{ENTER}")
#cs------------------------------------------------------------------------------------------------------------------------------
定义函数
#ce------------------------------------------------------------------------------------------------------------------------------

;处理有弹框的键盘模拟
func key_with_sysbol($key,$state)
	for $i in $key
		Sleep(1000)
		if $state == "^" Then
			;ctrl系列弹出框组合键
			Send("^{"& $i &"}")
		ElseIf $state== "!" Then
			;shift系列弹出框组合键
			Send("!{"& $i &"}")
		ElseIf $state =="F1" Then
			Send("{ $i }")
		ElseIf $state =="F3" Then
			Send("{ $i }")
		EndIf
		Sleep(1000)
		Send("{ESC}")
	Next
EndFunc

;处理win+key
Func win_key($key)
	for $i in $key
		Sleep(1000)
		If $i == "d" Then
			Send("#d")
		ElseIf $i == "e" Then
			Send("#e")
			Sleep(1000)
			Send("!{f4}")
		ElseIf  $i == "r" Then
			Send("#r")
			Sleep(1000)
			Send("{ESC}")
		EndIf
	Next
EndFunc

;处理没有弹出框的键盘模拟
Func key_singal($key,$state,$capslock)
	for $i in $key
		if $capslock==0 Then
			;不开启大写键情况下，模拟键盘输入
			Send($state & "{"& $i &"}")
		Else
			;开启大写键
			Send("{CAPSLOCK on}")
			Send($state & "{"& $i &"}")
		EndIf
	Next
	Sleep(1000)
	;关闭大写键
	Send("{CAPSLOCK OFF}")
EndFunc



