#cs----------------------------------------------------------------------------------------------------------------------
文件名字：disk_clipboard_last.au3
创建日期：2017/02/09
作者：顾亚茹
功能说明：模拟磁盘内部txt/word/excel三种文件之间文本内容的复制、剪切、粘贴操作（剪切板操作）
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include <Array.au3>
#include <WinAPIFiles.au3>

#cs-----------------------------------------------------------------------------------------------------------------------
打开文件
parameters:
$filePath: 文件路径                                       $fileTitle：窗口的标题
#ce-----------------------------------------------------------------------------------------------------------------------
Func openFile($filePath, $fileTitle)
	ShellExecute($filePath)
	Sleep(1000)
	Local $hWnd = WinWait($fileTitle, "", 10)
	WinActivate($hWnd)
	return $hWnd
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
查找关键字
parameters:
$nFindTabStatus: 文件的tab状态                                  $keyWord ： 查找的关键字
取消ctrl+f弹出框,excel和txt,需要按下4tab,word需要按下1tab
office版本2007
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
右键菜单 剪切、复制、粘贴操作
parameters:
$countDown: down键的数目
txt粘贴需要按下4down，word粘贴需要按下2down，excel粘贴需要按下2down
txt复制需要按下3down，word复制需要按下1down，excel复制需要按下1down
#ce-----------------------------------------------------------------------------------------------------------------------
Func menuOp_String($countDown)
	Send("{APPSKEY}")
	Sleep(1000)
	passDown($countDown)
	Sleep(100)
	Send("{ENTER}")
	Sleep(1500)
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
快捷键粘贴、剪切、复制操作
parameters:
$key: 操作的快捷键方式
ctrl+x ctrl+c ctrl+v
#ce-----------------------------------------------------------------------------------------------------------------------
Func hotKey_OpString($key)
	Send("^{"& $key &"}")
	Sleep(500)
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
关闭文件
parameters:
$fileTitle: 文件标题                                       $closeTitle ： 需要关闭窗口标题
#ce-----------------------------------------------------------------------------------------------------------------------
Func close_windows($fileTitle,$closeTitle)
	WinWait($fileTitle, "", 10)
	Sleep(1000)
	WinClose($fileTitle)
	If WinExists($closeTitle) Then
		$hWnd = WinGetHandle($closeTitle)
		ControlClick($hWnd, "", "Button2")
	EndIf
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
执行本地文件到远程文件的剪切复制粘贴操作
parameters:
$opKind1: 操作类型 （eg:txt右键复制 ->txt的右键粘贴动作类型）
#ce-----------------------------------------------------------------------------------------------------------------------
Func execute_cmd($opKind1)
		$opTmp = $opKind1
		openFile($opTmp[0], $opTmp[1])
		searchString($opTmp[2],$opTmp[3])
		If $opTmp[4] == 0 Then
			menuOp_String($opTmp[5])
		ElseIf $opTmp[4] == 1 Then
			hotKey_OpString($opTmp[5])
		EndIf
		close_windows($opTmp[6],$opTmp[7])
		Sleep(2000)
		openFile($opTmp[8], $opTmp[9])
		If $opTmp[10] == 0 Then
			menuOp_String($opTmp[11])
		ElseIf $opTmp[10] == 1 Then
			hotKey_OpString($opTmp[11])
		EndIf
		close_windows($opTmp[12],$opTmp[13])
		Sleep(2000)
EndFunc


;循环执行磁盘上行操作
#cs
Local $key = ["00001-wwrfz"]
For $j=0 To UBound($key) - 1
	Local $arrayKey = StringSplit($key[$j], "-")
	If $arrayKey[2] == "wwrfz" Then
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$txtTitle,$txtCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	EndIf
	execute_cmd($opKind)
Next
#ce

;循环执行磁盘下行操作
#cs
Local $key = ["00125-wwxrz"]
For $j=0 To UBound($key) - 1
	Local $arrayKey = StringSplit($key[$j], "-")
	If $arrayKey[2] == "wwxrz" Then ;00125 右键复制-远程右键粘贴word-word
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteWordTitle,$wordCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	EndIf
	execute_cmd($opKind)
Next
#ce
