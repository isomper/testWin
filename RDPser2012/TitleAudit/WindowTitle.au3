#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/07
作者：顾亚茹
功能说明：模拟遍历windows IE浏览器、写字板、计算机管理、网络连接的标题菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#cs-----------------------------------------------------------------------------------------------------------------------
;模拟遍历windows IE浏览器、写字板、计算机管理、网络连接的标题菜单
parameters:
$path: 执行文件的路径                                       $title：窗口的标题
$key： 打开窗口标题的快捷键                                  $nDown：按下down键的次数
#ce-----------------------------------------------------------------------------------------------------------------------
Func find_chrildren($list, $key, $index)
	For $i = 0 To UBound($list) - 1
		If $key == $list[$i][0] And $index == $list[$i][1] Then
			Send("{RIGHT}")
			Sleep(500)
			For $j = 1 To $list[$i][2] - 1
				Send("{DOWN}")
				Local $tmpStr = StringFormat("%d%d", Number($index), Number($j));
				find_chrildren($arr, $key, $tmpStr)
				Sleep(500)
			Next
			Sleep(500)
			Send("{LEFT}")
		EndIf
	Next
EndFunc

Func window_title($path,$title,$key,$nDown)
	ShellExecute($path)
	Local $pid = WinWaitActive($title)
	For $i =0 To UBound($key) - 1
		Send("!{"& $key[$i] &"}")
		Sleep(500)
		For $j =0 To $nDown[$i] -1
			find_chrildren($arr, $key[$i], String($j))
			Send("{DOWN}")
			Sleep(500)
		Next
	Next
	Send("{ESC}")
	WinClose($pid)
EndFunc


