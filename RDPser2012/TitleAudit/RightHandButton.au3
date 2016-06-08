#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/07
作者：顾亚茹
功能说明：模拟遍历windows系统右键菜单，开始菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------
#cs-----------------------------------------------------------------------------------------------------------------------
;.MSC方式打开模块，遍历右键菜单
parameters:
$path:  执行文件的路径               $title：窗口的标题           $nTab:按下tab键的次数            $nDown：按下down键的次数
$nRight：按下right键的次数           $n2ndDown: 按下二级菜单down的个数        $nAppDown：appskey右键菜单功能键个数
#ce-----------------------------------------------------------------------------------------------------------------------
Func msc_appskey($path,$title,$nDown,$nRight,$n2ndDown,$nAppDown)
	ShellExecute($path)
	Local $nWnd =WinWaitActive($title)
	For $j = 1 To $nDown
		Send("{DOWN}")
		Sleep(500)
	Next
	For $b = 1 To $nRight
		Send("{RIGHT}")
		Sleep(500)
	Next
	For $b = 1 To $n2ndDown
		Send("{DOWN}")
		Sleep(500)
	Next
	Sleep(500)
	Send("{APPSKEY}")
	For $c = 1 To $nAppDown
		Send("{DOWN}")
		find_children($subItem, String($c))
		Sleep(500)
	Next
	Send("{esc}")
	Sleep(500)
	WinClose($nWnd)
EndFunc
;遍历计算机管理

Func msc_system_mgt($path,$title,$nDown,$index)
	ShellExecute($path)
	Local $nWnd=WinWaitActive($title)
	For $i =0 to UBound($nDown)-1
		For $j=0 to $nDown[$i]-1
			Send("{DOWN}")
			If  $i== 6 Then
				Send("{left 4}")
			EndIf
			Sleep(500)
		Next
		Send("{RIGHT}")
	Sleep(500)
	Next
	WinClose($nWnd)
EndFunc


#cs--------------------------------------------------------------------------------------------------------------------
;模拟点击鼠标右键
parameters:
$action:  点击鼠标（左键或者右键）        $x：当前位置横坐标                              $y:当前位置竖坐标
$nClicks：鼠标动作点击个数               $nDown：按下down键的次数
#ce-----------------------------------------------------------------------------------------------------------------------
Func mouse_right($action,$x,$y,$nClicks,$nDown)
	MouseClick($action,$x,$y,$nClicks)
	For $i =1 To $nDown
		Send("{DOWN}")
		find_children($subItem, String($i))
		Sleep(500)
	Next
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;模拟点击开始菜单
parameters:
$key:  呼起应用的快捷键                                     $nTab：按下tab键的次数
$nUp： 按下up键的次数                                       $nDown：按下down键的次数
#ce-----------------------------------------------------------------------------------------------------------------------
;遍历菜单子菜单及其右键内容
Func click_start_all($key,$nTab,$nUp,$nDown)
	Send($key)
	Sleep(500)
	;Send("{TAB "& $nTab &"}")
	;Sleep(500)
	For $i = 1 To $nUp
		Send("{up}")
		find_children($subItem, String($i))
		;second_level($arr,$nUp,$nArrowDown,$index)
		Sleep(500)
		Send("{APPSKEY}")
		Sleep(500)
		For $j =1 To $nDown[$i-1]
			Send("{DOWN}")
			Sleep(500)
		Next
		Send("{ESC}")
	Next
	Sleep(200)
EndFunc
;判断是否有子菜单
Func find_children($subList, $indexStr)
	For $i = 0 To UBound($subList) - 1
		If $subList[$i][0] == $indexStr Then
			Sleep(1000)
			Send("{right}")
			For $j = 1 To $subList[$i][1]
				Send("{DOWN}")
				Local $tmpStr = StringFormat("%d%d", Number($indexStr), Number($j));
				ConsoleWrite($tmpStr)
				ConsoleWrite("|")
				find_children($subList, $tmpStr)
				Sleep(500)
			Next
			Send("{LEFT}")
			Sleep(500)
		EndIf
	Next
EndFunc

