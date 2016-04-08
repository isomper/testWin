#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/07
作者：顾亚茹
功能说明：模拟遍历windows系统右键菜单，开始菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

;点击开始菜单并遍历二级菜单

Func start_cell($key,$nTab,$nUp,$nDown,$nArrowDown,$state)
	If $state == 1 Then
		click_start_all($key,$nTab,$nUp,$nDown)
	ElseIf $state == 2 Then
		click_start($key,$nTab,$nUp,$nDown)
	ElseIf $state == 3 Then
		click_start($key,$nTab,$nUp,$nDown)
		second_level($nArrowDown)
	EndIf
	Send("!{F4}")
EndFunc
;遍历右键子菜单
Func right_cell($action,$x,$y,$nClicks,$nDown,$nArrowDown,$state)
	If $state == 1 Then
		mouse_right($action,$x,$y,$nClicks,$nDown)
	ElseIf $state == 2 Then
		mouse_right($action,$x,$y,$nClicks,$nDown)
		second_level($nArrowDown)
	EndIf
	Send("{ESC}")
EndFunc
;遍历msc程序右键菜单的子菜单
Func msc_cell($path,$title,$nTab,$nDown,$nRight,$n2ndDown,$nAppDown,$nArrowDown,$state)
	if $state == 1 Then
		msc($path,$title,$nTab,$nDown,$nRight,$n2ndDown,$nAppDown)
	ElseIf $state==2 Then
		msc($path,$title,$nTab,$nDown,$nRight,$n2ndDown,$nAppDown)
		second_level($nArrowDown)
	EndIf
	Send("!{F4}")
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;.MSC方式打开模块，遍历右键菜单
parameters:
$path:  执行文件的路径               $title：窗口的标题           $nTab:按下tab键的次数            $nDown：按下down键的次数
$nRight：按下right键的次数           $n2ndDown: 按下二级菜单down的个数        $nAppDown：appskey右键菜单功能键个数
#ce-----------------------------------------------------------------------------------------------------------------------

Func msc($path,$title,$nTab,$nDown,$nRight,$n2ndDown,$nAppDown)
	ShellExecute($path)
	WinWaitActive($title)
	For $i = 1 To $nTab
		Send("{TAB}")
		Sleep(500)
	Next
	For $j = 1 To $nDown
		Send("{DOWN}")
		Sleep(500)
	Next
	For $a = 1 To $nRight
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
		Sleep(500)
	Next
	Sleep(500)
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
;模拟点击鼠标右键
parameters:
$action:  点击鼠标（左键或者右键）        $x：当前位置横坐标                              $y:当前位置竖坐标
$nClicks：鼠标动作点击个数               $nDown：按下down键的次数
#ce-----------------------------------------------------------------------------------------------------------------------
Func mouse_right($action,$x,$y,$nClicks,$nDown)
	MouseClick($action,$x,$y,$nClicks)
	For $i =1 To $nDown
		Send("{DOWN}")
		Sleep(500)
	Next
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;模拟点击开始菜单
parameters:
$key:  呼起应用的快捷键                                     $nTab：按下tab键的次数
$nUp： 按下up键的次数                                       $nDown：按下down键的次数
#ce-----------------------------------------------------------------------------------------------------------------------
Func click_start($key,$nTab,$nUp,$nDown)
	Send($key)
	Sleep(500)
	Send("{TAB "& $nTab &"}")
	Sleep(500)
	Send("{UP  "& $nUp &"}")
	Send("{APPSKEY}")
	Sleep(500)
	For $i =1 to $nDown
		Send("{DOWN}")
		Sleep(500)
	Next
	;Send($key)
	Sleep(500)
EndFunc
;遍历一级菜单及其右键内容
Func click_start_all($key,$nTab,$nUp,$nDown)
	Send($key)
	Sleep(500)
	Send("{TAB "& $nTab &"}")
	Sleep(500)
	For $i = 1 To $nUp
		Send("{up}")
		Sleep(500)
		Send("{APPSKEY}")
		Sleep(500)
		For $j =1 To $nDown[$i-1]
			Send("{DOWN}")
			Sleep(500)
		Next
	Send("{ESC}")
	Next
	Sleep(500)
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;遍历二级菜单
parameters:
$narrowDown: 二级菜单的数目
#ce-----------------------------------------------------------------------------------------------------------------------
Func second_level($nArrowDown)
	Send("{RIGHT}")
	For $i =1 to $nArrowDown
		Send("{DOWN}")
		Sleep(500)
	Next
	Send("{LEFT}")
EndFunc

