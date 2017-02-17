#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2017/02/07
作者：顾亚茹
功能说明：模拟磁盘外部txt/word/excel三种文件之间的文本内容的复制、剪切、粘贴操作（剪切板操作）
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include <WinAPIFiles.au3>
#include <MsgBoxConstants.au3>

;定义临时文件路径
Global $tmpPath = "D://au3//tmp.txt"
;定义sikulix安装路径
Global $sikulixInsPath = "D:\sikulix\runsikulix -r"

#cs-----------------------------------------------------------------------------------------------------------------------
打开文件
parameters:
$filePath: 文件的路径                                       $fileTitle：窗口的标题
#ce-----------------------------------------------------------------------------------------------------------------------
Func openFile($filePath, $fileTitle)
	ShellExecute($filePath)
	Local $hWnd = WinWait($fileTitle, "", 10)
	WinActivate($hWnd)
	Send("{BACKSPACE}")
	Sleep(500)
	;return $hWnd
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
查找关键字
parameters:
$nFindTabStatus: 文件的tab状态                                  $keyWord ： 查找的关键字
取消ctrl+f关键键excel和txt，excel需要按下4tab，word需要按下1tab
0代表txt,1代表word,2代表excel
#ce-----------------------------------------------------------------------------------------------------------------------
Func searchString($nFindTabStatus, $keyWord)
	Send("^{f}")
	Sleep(1000)
	Send($keyWord)
	Sleep(1000)
	Send("{ENTER}")
	Sleep(1000)
	If $nFindTabStatus == 0 Then
		Send("{TAB 4}")
	ElseIf $nFindTabStatus ==1 then
		Send("{TAB 1}")
	ElseIf $nFindTabStatus ==2 then
		Send("{TAB 4}")
	EndIf
	Sleep(1000)
	Send("{ENTER}")
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
模拟按下down键
parameters:
$count: down键的数目
#ce-----------------------------------------------------------------------------------------------------------------------
Func passDown($count)
	For $i = 0 To $count - 1
		Send("{DOWN}")
		Sleep(100)
	Next
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
右键菜单复制、剪切、粘贴
parameters:
$countDown: down键的数目
txt粘贴需要按下4down，word粘贴需要按下2down，excel粘贴需要按下2down
#ce-----------------------------------------------------------------------------------------------------------------------
Func menuOp_opString($countDown)
	Send("{APPSKEY}")
	Sleep(1000)
	passDown($countDown)
	Sleep(500)
	Send("{ENTER}")
	Sleep(2000)
EndFunc


#cs-----------------------------------------------------------------------------------------------------------------------
快捷键复制、剪切、粘贴
parameters:
$key: 操作的快捷键方式
ctrl+v ctrl+c crtl+x
#ce-----------------------------------------------------------------------------------------------------------------------
Func hotKey_opString($key)
	Send("^{"& $key &"}")
	Sleep(1000)
EndFunc


#cs-----------------------------------------------------------------------------------------------------------------------
关闭文件
parameters:
$fileTitle: 文件标题                                       $closeTitle ： 需要关闭窗口标题
#ce-----------------------------------------------------------------------------------------------------------------------
Func close_windows($fileTitle,$closeTitle)
	WinWait($fileTitle, "", 10)
	Sleep(100)
	WinClose($fileTitle)
	Sleep(1000)
	If WinExists($closeTitle) Then
		$hWnd = WinGetHandle($closeTitle)
		ControlClick($hWnd, "", "Button2")
		Sleep(2000)
	EndIf
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
调用sikulix
parameters:
$cmd: 执行命令
local $sikulixPath = "D:\sikulix\runsikulix -r"
local $sikulixPath =@DesktopDir & "\runsikulix -r"
#ce-----------------------------------------------------------------------------------------------------------------------
Func remoteOp($cmd)
	Sleep(1000)
	$hWnd = WinWait("[CLASS:TscShellContainerClass]", "", 10)
	WinActivate($hWnd)
	Run(@ComSpec & " /c " & $sikulixInsPath & $cmd,"", @SW_HIDE,"")
	Sleep(15000)
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
内容写入文件
parameters:
$str: 写入内容
D://tmpPaste.txt ： 文件路径
#ce-----------------------------------------------------------------------------------------------------------------------
Func writeTmpTxt($str)
	Local $hFileOpen = FileOpen($tmpPath, $FO_OVERWRITE)
	If $hFileOpen <> -1 Then
		FileWriteLine($hFileOpen, $str)
	EndIf
	FileClose($hFileOpen)
EndFunc

;循环执行 下行调试数据
#cs
Local $key = ["003-twr"]
Local $opKind = [""]
For $j=0 To UBound($key) - 1
	Local $arrayKey = StringSplit($key[$j], "-")
	If $arrayKey[2] == "twr" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	EndIf
	execute_cmd($opKind)
Next
#ce

;上行调试数据
#cs
Local $key = ["001-eexrz"]
For $j=0 To UBound($key) - 1
	Local $arrayKey = StringSplit($key[$j], "-")
	If $arrayKey[2] == "eexrz" Then
		Local $opKind =  [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$cut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu,$opMenu_pasteStatus]
	EndIf
	execute_cmd($opKind)
Next
#ce
