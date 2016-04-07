#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/07
作者：顾亚茹
功能说明：模拟遍历windows系统右键菜单，开始菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

;点击开始菜单并遍历二级菜单

Func startcell($key,$ntab,$nUp,$nDown,$narrowDown)
	clickstart($key,$ntab,$nUp,$nDown)
	second_level($narrowDown)
	Send("{ESC}")
EndFunc
;遍历右键子菜单
Func rightcell($action,$x,$y,$nClicks,$nDown,$narrowDown)
	Mouseright($action,$x,$y,$nClicks,$nDown)
	second_level($narrowDown)
	Send("{ESC}")
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;.MSC方式打开模块，遍历右键菜单
parameters:
$site:  执行文件的路径               $name：窗口的标题           $ntab:按下tab键的次数            $nDown：按下down键的次数
$nright：按下right键的次数           $n2ndDown: 按下二级菜单down的个数        $nAppDown：appskey右键菜单功能键个数
#ce-----------------------------------------------------------------------------------------------------------------------

Func _Msc($site,$name,$ntab,$nDown,$nright,$n2ndDown,$nAppDown)
	ShellExecute($site)
	Global $pid = WinWaitActive($name)
	For $i = 1 To $ntab
		Send("{tab}")
		Sleep(500)
	Next
	For $j = 1 To $nDown
		Send("{down}")
		Sleep(500)
	Next
	For $a = 1 To $nright
		Send("{right}")
		Sleep(500)
	Next
	For $b = 1 To $n2ndDown
		Send("{down}")
		Sleep(500)
	Next
	Sleep(500)
	Send("{APPSKEY}")
	For $c = 1 To $nAppDown
		Send("{down}")
		Sleep(500)
	Next
	Sleep(500)
EndFunc
;遍历msc程序右键菜单的子菜单
Func Msccell($site,$name,$ntab,$nDown,$nright,$n2ndDown,$nAppDown,$narrowDown,$state)
	if $state == 1 Then
		_Msc($site,$name,$ntab,$nDown,$nright,$n2ndDown,$nAppDown)
		WinClose($pid)
	ElseIf $state==2 Then
		_Msc($site,$name,$ntab,$nDown,$nright,$n2ndDown,$nAppDown)
		second_level($narrowDown)
		Send("!{f4}")
	EndIf
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;模拟点击鼠标右键
parameters:
$action:  点击鼠标（左键或者右键）        $x：当前位置横坐标                              $y:当前位置竖坐标
$nClicks：鼠标动作点击个数               $nDown：按下down键的次数
#ce-----------------------------------------------------------------------------------------------------------------------
Func Mouseright($action,$x,$y,$nClicks,$nDown)
	MouseClick($action,$x,$y,$nClicks)
	For $i =1 To $nDown
		Send("{down}")
		Sleep(500)
	Next
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;模拟点击开始菜单
parameters:
$key:  呼起应用的快捷键                                     $ntab：按下tab键的次数
$nUp： 按下up键的次数                                       $nDown：按下down键的次数
#ce-----------------------------------------------------------------------------------------------------------------------
Func clickstart($key,$ntab,$nUp,$nDown)
	Send($key)
	Sleep(500)
	Send("{tab "& $ntab &"}")
	Sleep(500)
	Send("{UP  "& $nUp &"}")
	Send("{APPSKEY}")
	Sleep(500)
	For $i =1 to $nDown
		Send("{down}")
		Sleep(500)
	Next
	Send($key)
	Sleep(1000)
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;遍历二级菜单
parameters:
$narrowDown: 二级菜单的数目
#ce-----------------------------------------------------------------------------------------------------------------------
Func second_level($narrowDown)
	Send("{right}")
	For $i =1 to $narrowDown
		Send("{down}")
		Sleep(1000)
	Next
	Send("{left}")
EndFunc

