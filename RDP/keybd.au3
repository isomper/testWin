;打开D:\CS.docx路径下的cs.docx文件
;RUN函数用法("program"[, "workingdir" [, show_flag [, opt_flag]]] )program代表应用程序的全路径
;Run("C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE D:\CS.docx")
ShellExecute(@DesktopDir & "\1.txt")
WinWaitActive("1.txt - 记事本","",2)
Opt("SendCapslockMode",0)
;定义数组
Dim $state
Local $key_num[10] = [0,1,2,3,4,5,6,7,8,9]
Local $key_fn[15] = ["F5","F2"]
Local $key_sysbol[20] = ["`","[","]","\",";","'",",",".","/","SPACE"]
Local $key_with_ctrl[20] = ["n","o","s","x","c","v"]
Local $key_with_ctrl_sysbol[2] = ["f","o"];,"g","h","p","n"]
Local $key_with_win[10] = ["m","d","e","r","l"]
Local $key_with_alt[5] = ["tab","esc"]
Local $key_letter[26] = ["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","s","y","z"]

;调用函数
;key_singal($key_letter,"+",1)
;key_singal($key_letter,"+",0)
;key_singal($key_num,"",0)


;定义函数
#cs
Func key_with($with_,$state,$capslock)

	If $capslock==0 Then
		ConsoleWrite("xiaoxie" & @CRLF)
		key_singal($with_,$state)
	Else
		ConsoleWrite("daxie" & @CRLF)
		Send("{CAPSLOCK}")
		key_singal($with_,$state)
		Send("{CAPSLOCK}")
	EndIf
EndFunc
#ce

#cs
;执行按下的快捷键
for $i = 0 To UBound($key_with_ctrl_sysbol)-1
	key_with_sysbol($key_with_ctrl_sysbol[$i],"^")
Next
#ce

;key_singal($key_letter,"",1)
;key_singal($key_letter,"+",1)
key_with_sysbol($key_with_ctrl_sysbol,"^")


func key_with_sysbol($letter,$sysbol)
	for $i in $letter

		if $sysbol == "^" Then
			Sleep(2000)
			Send("^{"& $i &"}")
			ConsoleWrite("sysbol1" & @CRLF)
			sleep(2000)
		ElseIf $sysbol== "!" Then
			sleep(2000)
			Send("!{"& $i &"}")
			ConsoleWrite("sysbol2" & @CRLF)
			sleep(2000)
		EndIf
		;Send("{ESC}")
	Next
EndFunc



Func key_singal($key,$state,$capslock)
	for $i in $key
		if $capslock==0 Then
		;Sleep(300)
		Send($state & "{"& $i &"}")
		Else
		;ConsoleWrite("xiaoxie" & @CRLF)
		Send("{CAPSLOCK on}")
		;Sleep(1000)
		Send($state & "{"& $i &"}")
		EndIf
	Next
	Send("{CAPSLOCK OFF}")
EndFunc



