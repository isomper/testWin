;打开D:\CS.docx路径下的cs.docx文件
;RUN函数用法("program"[, "workingdir" [, show_flag [, opt_flag]]] )program代表应用程序的全路径
;Run("C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE D:\CS.docx")
ShellExecute(@DesktopDir & "\1.txt")
WinWaitActive("1.txt - 记事本","",2)
Opt("SendCapslockMode",0)
;定义数组
Dim $state
Local $key_num[10] = [0,1,2,3,4,5,6,7,8,9]
Local $key_sysbol[23] = ["`","[","]","\",";","'",",",".","/","~","_","-","+","=","{","}",":","<",">","?",'"',"SPACE","ENTER"]
Local $key_with_ctrl[5] = ["a","x","v","c","s"]
Local $key_with_ctrl_sysbol[6] = ["f","o","g","h","p","n"]
Local $key_with_alt[1] = ["TAB"]
Local $key_with_win[3] = ["d","e","r"]
Local $key_fn[12] = ["F2","F4","F5","F6","F7","F8","F9","F10","F11","F12"]
Local $key_f1[1] = ["F1"]
Local $key_f3[1] = ["F3"]
Local $key_letter[26] = ["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"]


#cs
;执行按下的快捷键
for $i = 0 To UBound($key_with_ctrl_sysbol)-1
	key_with_sysbol($key_with_ctrl_sysbol[$i],"^")
Next
#ce
;调用函数

key_singal($key_num,"",0);输入0~9
key_singal($key_num,"+",0);输入!@#$%^&*()符号
key_singal($key_sysbol,"",0);输入符号（`[]\;',./~{}|:<>?空格）
key_singal($key_letter,"",0);输入26个小写字母
key_singal($key_letter,"",1);输入26个大写字母
key_singal($key_letter,"+",1);开启大写键，按下shift输入26个小写字母
key_singal($key_with_ctrl,"^",0);按下ctrl+{s,x,c,v,a}
key_with_sysbol($key_with_ctrl_sysbol,"^")
key_singal($key_fn,"",0)
key_singal($key_with_alt,"!",0)
key_singal($key_with_win,"#",0)
key_with_sysbol($key_f1,"F1")
key_with_sysbol($key_f1,"F3")



func key_with_sysbol($letter,$sysbol)
	for $i in $letter
		Sleep(1000)
		if $sysbol == "^" Then

			Send("^{"& $i &"}")
			ConsoleWrite("sysbol1" & @CRLF)
			sleep(1000)
		ElseIf $sysbol== "!" Then

			Send("!{"& $i &"}")
			ConsoleWrite("sysbol2" & @CRLF)
			sleep(1000)
		ElseIf $sysbol =="F1" Then
			Send("{ $i }")
			ConsoleWrite("sysbol13" & @CRLF)
		ElseIf $sysbol =="F3" Then
			Send("{ $i }")
			ConsoleWrite("sysbol14" & @CRLF)
		EndIf
		Sleep(1000)
		Send("{ESC}")
	Next
EndFunc



Func key_singal($key,$state,$capslock)
	for $i in $key
		if $capslock==0 Then
		Sleep(1000)
		Send($state & "{"& $i &"}")
		Else
		;ConsoleWrite("xiaoxie" & @CRLF)
		Send("{CAPSLOCK on}")
		Sleep(1000)
		Send($state & "{"& $i &"}")
		EndIf
	Next
	Send("{CAPSLOCK OFF}")
EndFunc



