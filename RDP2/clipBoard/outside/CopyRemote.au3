#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2017/02/08
作者：顾亚茹
功能说明：模拟服务器到本地（下行）txt/word/excel三种文件之间的文本内容的复制、剪切、粘贴操作（剪切板操作）
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include <Array.au3>

Local $hFileOpen  = FileOpen("\\tsclient\D\au3\tmp.txt")
Local $sFileRead = FileRead($hFileOpen)
ConsoleWrite($sFileRead)
$filePath = StringSplit($sFileRead," ")[1]
$searchKey = StringSplit($sFileRead," ")[2]
$opMethod = StringSplit($sFileRead," ")[3]
$opStatus = StringSplit($sFileRead," ")[4]
FileClose($hFileOpen)

#cs-----------------------------------------------------------------------------------------------------------------------
打开文件
parameters:
$filePath: 执行文件的路径                                       $fileTitle：窗口的标题
#ce-----------------------------------------------------------------------------------------------------------------------
Func openFile($filePath, $fileTitle)
	ShellExecute($filePath)
	Local $hWnd = WinWait($fileTitle, "", 10)
	WinActivate($hWnd)
	return $hWnd
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
关闭文件
parameters:
$fileTitle: 文件标题                                       $closeTitle ： 需要关闭窗口标题
#ce-----------------------------------------------------------------------------------------------------------------------
Func close_windows($fileTitle, $closeTitle)
	WinWait($fileTitle, "", 10)
	Sleep(100)
	WinClose($fileTitle)
	If WinExists($closeTitle) Then
		$hWnd = WinGetHandle($closeTitle)
		ControlClick($hWnd, "", "Button2")
	EndIf
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
查找关键字
parameters:
$nFindTabStatus: 文件的tab状态                                  $keyWord ： 查找关的键字
取消ctrl+f关键键excel和txt，excel需要按下4tab，word需要按下1tab
0代表txt,1代表word,2代表excel
#ce-----------------------------------------------------------------------------------------------------------------------
Func searchString($nFindTabStatus, $keyWord)
	Send("^{f}")
	Sleep(1000)
	Send($keyWord)
	Sleep(1000)
	Send("{ENTER}")
	Sleep(500)
	If $nFindTabStatus == 0 Then
		Send("{TAB 4}")
	ElseIf $nFindTabStatus ==1 Then
		Send("{TAB 1}")
	ElseIf $nFindTabStatus ==2 Then
		Send("{TAB 4}")
	EndIf
	Send("{ENTER}")
	Sleep(500)
EndFunc

Func passDown($count)
	For $i = 0 To $count - 1
		Send("{DOWN}")
	Next
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
右键菜单复制
parameters:
$countDown: down键的数目
txt复制需要按下3down，word复制需要按下1down，excel复制需要按下2down
txt复制需要按下2down，word复制需要按下0down，excel复制需要按下1down
#ce-----------------------------------------------------------------------------------------------------------------------
Func menuOp_copyString($countDown)
	Send("{APPSKEY}")
	Sleep(1000)
	passDown($countDown)
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
右键菜单剪切、复制
parameters:
$key: 操作快捷键
ctrl+c ctrl +x
#ce-----------------------------------------------------------------------------------------------------------------------
Func hotKey_copyString($key)
	Send("^{" & $key & "}")
	Sleep(1000)
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
获取操作系统版本信息
#ce-----------------------------------------------------------------------------------------------------------------------
Func GetOSVersion()
	$objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	$colItems = $objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	For $os In $colItems
		return $os.Caption&" "&$os.Version
	Next
EndFunc

;文件类型
$fileAssert = StringSplit($filePath, ".")[2]
$fileArray = StringSplit($filePath, "\")
;文件名字
$fileName = $fileArray[UBound($fileArray) - 1]

;定义txt中按几下tab键关闭Ctrl+f窗口，定义word中按几下tab键关闭Ctrl+f窗口，定义excel中按几下tab键关闭Ctrl+f窗口
$nFindTabStatus = 0

;定义txt右键按3下到复制键，txt右键按2下到剪切键.  word右键按1下到复制键，word右键按0下到剪切键.  excel右键按1下到复制键，excel右键按0下到剪切键.
;剪切和复制的down数目
$downCount = 0

$fileTitle = ""
$closeTitle = ""
$hotKey = ""
Switch $fileAssert
	Case "txt"
		$nFindTabStatus = 0
		If Int($opStatus) == 0 Then
			$downCount = 3
		ElseIf Int($opStatus) == 1 Then
			$downCount = 2
		ElseIf Int($opStatus) == 2 Then
			$hotKey = "c"
		ElseIf Int($opStatus) == 3 Then
			$hotKey = "x"
		EndIf
		$fileTitle = $fileName & " - 记事本"
		$closeTitle = "记事本"
	Case "docx"
		$nFindTabStatus = 1
		If Int($opStatus) == 0 Then
			;复制down
			$downCount = 1
		ElseIf Int($opStatus) == 1 Then
			;剪切down
			$downCount = 0
		ElseIf Int($opStatus) == 2 Then
			$hotKey = "c"
		ElseIf Int($opStatus) == 3 Then
			$hotKey = "x"
		EndIf
		$fileTitle = $fileName & " - Microsoft Word"
		$closeTitle = "Microsoft Office Word"
	Case "doc"
		$nFindTabStatus = 1
		$copyDownCount = 1
		$cutDownCount = 0
		$fileTitle = $fileName & " - Microsoft Word"
		$closeTitle = "Microsoft Office Word"
	Case "xlsx"
		$nFindTabStatus = 2
		If Int($opStatus) == 0 Then
			$downCount = 1
		ElseIf Int($opStatus) == 1 Then
			$downCount = 0
		ElseIf Int($opStatus) == 2 Then
			$hotKey = "c"
		ElseIf Int($opStatus) == 3 Then
			$hotKey = "x"
		EndIf
		$fileTitle = "Microsoft Excel - " & $fileName
		$closeTitle = "Microsoft Office Excel"
	Case "xls"
		$nFindTabStatus = 2
		$copyDownCount = 1
		$cutDownCount = 0
		$fileTitle = "Microsoft Excel - " & $fileName
		$closeTitle = "Microsoft Office Excel"
EndSwitch

openFile($filePath, $fileTitle)
searchString($nFindTabStatus, $searchKey)
Sleep(1000)
;0代表右键复制或者剪切,1代表组合键复制或者粘贴
If Int($opMethod) == 0 Then
	menuOp_CopyString($downCount)
ElseIf Int($opMethod) == 1 Then
	hotKey_copyString($hotKey)
EndIf
close_windows($fileTitle, $closeTitle)
Sleep(1000)
Local $os_Version =  StringSplit(GetOSVersion()," ")[4]
If $os_Version == "2008" Then
	Send("{TAB}")
EndIf