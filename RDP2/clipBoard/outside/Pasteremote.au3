#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2017/02/08
作者：顾亚茹
功能说明：模拟从本地到远程服务器（上行）txt/word/excel三种文件之间的文本内容的复制、剪切、粘贴操作（剪切板操作）
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include <Array.au3>
#include <WinAPIFiles.au3>

Local $hFileOpen  = FileOpen("\\tsclient\D\au3\tmp.txt")
Local $sFileRead = FileRead($hFileOpen)
$filePath = StringSplit($sFileRead," ")[1]
$opMethod = StringSplit($sFileRead," ")[2]
$opStatus = StringSplit($sFileRead," ")[3]
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
	Sleep(1000)
	If WinExists($closeTitle) Then
		$hWnd = WinGetHandle($closeTitle)
		ControlClick($hWnd, "", "Button2")
	EndIf
EndFunc

Func passDown($count)
	For $i = 0 To $count - 1
		Send("{DOWN}")
		Sleep(500)
	Next
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
右键菜单粘贴
parameters:
$countDown: down键的数目
txt粘贴需要按下4down，word粘贴2需要按下1down，excel粘贴需要按下2down
#ce-----------------------------------------------------------------------------------------------------------------------
Func menuOp_pasteString($countDown)
	Send("{APPSKEY}")
	Sleep(1000)
	passDown($countDown)
	Sleep(500)
	Send("{ENTER}")
	Sleep(1000)
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
右键菜单粘贴
parameters:
$key: 操作快捷键
组合键ctrl+v
#ce-----------------------------------------------------------------------------------------------------------------------
Func hotKey_pasteString($key)
	Send("^{" & $key & "}")
	Sleep(500)
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

;定义txt右键按4下到粘贴键，txt右键按3下到粘贴键.  word右键按2下到粘贴键.  excel右键按2下到粘贴键.
$pasteDownCount = 0

$fileTitle = ""
$closeTitle = ""
$hotKey = ""

Switch $fileAssert
	Case "txt"
		If Int($opStatus) == 0 Then
			$pasteDownCount = 4
		ElseIf Int($opStatus)== 1 Then
			$hotKey = "v"
		EndIf
		$fileTitle = $fileName & " - 记事本"
		$closeTitle = "记事本"
	Case "docx"
		If Int($opStatus) == 0 Then
			$pasteDownCount = 2
		ElseIf Int($opStatus) == 1 Then
			$hotKey = "v"
		EndIf
		$fileTitle = $fileName & " - Microsoft Word"
		$closeTitle = "Microsoft Office Word"
	Case "doc"
		$nFindTabStatus = 1
		$cutDownCount = 0
		$fileTitle = $fileName & " - Microsoft Word"
		$closeTitle = "Microsoft Office Word"
	Case "xlsx"
		If Int($opStatus) == 0 Then
			$pasteDownCount = 2
		ElseIf Int($opStatus) == 1 Then
			$hotKey = "v"
		EndIf
		$fileTitle = "Microsoft Excel - " & $fileName
		$closeTitle = "Microsoft Office Excel"
	Case "xls"
		$nFindTabStatus = 2
		$cutDownCount = 0
		$fileTitle = "Microsoft Excel - " & $fileName
		$closeTitle = "Microsoft Office Excel"
EndSwitch

openFile($filePath, $fileTitle)
If $opMethod == 0 Then
	menuOp_pasteString($pasteDownCount)
ElseIf $opMethod == 1 Then
	hotKey_pasteString($hotKey)
EndIf
close_windows($fileTitle,$closeTitle)
Sleep(1000)
Local $os_Version =  StringSplit(GetOSVersion()," ")[4]
If $os_Version == "2008" Then
	Send("{TAB}")
EndIf