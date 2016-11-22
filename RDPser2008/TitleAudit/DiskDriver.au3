#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/26
作者：顾亚茹
功能说明：模拟文件和文件夹的新建、复制、剪切、删除及word、excel、txt文件之间相互粘贴、复制、剪切操作
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------


#include<Word.au3>
#include<Excel.au3>
#include<file.au3>
#include <MsgBoxConstants.au3>

;create_word("C:\Users\guyaru.pc\Desktop\test.doc")

#cs-----------------------------------------------------------------------------------------------------------------------
;创建文档
parameters:
$fileDir:  保存文件的路径         $path:调用程序路径       $title:当前程序标题       $instance：实例名称
#ce-----------------------------------------------------------------------------------------------------------------------
Func create_file($fileDir,$path,$title,$instance)
	If FileExists($fileDir) Then
		FileDelete($fileDir )
	EndIf
	ShellExecute($path)
	Local $nWnd = WinWaitActive($title,"",2)
	Send("^s")
	Sleep(1000)
	Local $wndSaveas = WinWaitActive("另存为","",1)
	ControlSetText($wndSaveas,"","Edit1",$fileDir)
	Sleep(1000)
	;ControlClick($wndSaveas,"",$instance)
	Opt("MouseCoordMode", 2) ;设置鼠标函数的坐标参照,相对当前激活窗口客户区坐标 Opt("MouseCoordMode", 1) ;1=absolute, 0=relative, 2=client
	$a=ControlGetPos($wndSaveas,"",$instance)   ;获取指定控件相对其窗口的坐标位置和大小
	MouseClick("left",$a[0],$a[1])
	Sleep(500)
	WinClose($nWnd)
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
;创建文件夹
parameters:
$fileDir:  保存文件的路径
#ce-----------------------------------------------------------------------------------------------------------------------
Func create_folder($fileDir)
	DirCreate($fileDir)
EndFunc

#cs-----------------------------------------------------------------------------------------------------------------------
;复制、剪切、删除文件和文件夹
parameters:
$sourceDir: 源文件的路径         $destCopyDir:目标复制路径     $destMoveDir:目标剪切路径
#ce-----------------------------------------------------------------------------------------------------------------------
Func move_file($sourceDir,$destCopyDir,$destMoveDir)

	;获取文件类型
	;Local $type = FileGetAttrib($sourceDir)
	;ConsoleWrite($type & @CRLF)
	If StringInStr(FileGetAttrib($sourceDir), "D") > 0 Then
		DirCopy($sourceDir,$destCopyDir)
		ConsoleWrite($destCopyDir & @CRLF)
		DirMove($destCopyDir,$destMoveDir)
		ConsoleWrite($destMoveDir & @CRLF)
		;DirRemove($destMoveDir & "bb")
		ConsoleWrite($destMoveDir  & @CRLF)
	ElseIf StringInStr(FileGetAttrib($sourceDir), "D") = 0 Then
		FileCopy($sourceDir,$destCopyDir)
		FileMove($destCopyDir,$destMoveDir)
		FileDelete($destMoveDir)
	EndIf
EndFunc


#cs
Func create_word($fileDir)
	If FileExists($fileDir) Then
		FileDelete($fileDir )
	EndIf
	;Local $wordProcess = _Word_Create()
	;Local $wordDoc = _Word_DocAdd($wordProcess)
	;Local $oRange = _Word_DocRangeSet($wordDoc, -1);焦点放在最近开始处
	;$oRange.Text = "这是标题. "
	;_Word_DocSaveAs($wordWnd, $fileDir);保存
	;_Word_Quit($wordProcess)
	ShellExecute("C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.EXE")
	Local $wordWnd = WinWaitActive("文档 1 - Microsoft Word","",2)
	Send("^s")
	Sleep(1000)
	Local $wndSaveas = WinWaitActive("另存为","",1)
	ControlSetText($wndSaveas,"","Edit1",$fileDir)
	Sleep(1000)
	ControlClick($wndSaveas,"","Button8")
	WinClose($wordWnd)
EndFunc

;create_excel("C:\Users\guyaru.pc\Desktop\cc.xls")
#cs-----------------------------------------------------------------------------------------------------------------------
;创建excel文件
parameters:
$fileDir:  保存文件的路径
#ce-----------------------------------------------------------------------------------------------------------------------
Func create_excel($fileDir)
	If FileExists($fileDir) Then
		FileDelete($fileDir )
	EndIf
	Local $excelProcess = _Excel_Open()
	Local $oExcel = _Excel_BookNew($excelProcess,1)
	;_Excel_BookSaveAs  ($oexcel,$fileDir,$xlWorkbookDefault , True)
	Sleep(1000)
	Send("^s")
	Sleep(1000)
	Local $wndSaveas = WinWaitActive("另存为","",2)
	Sleep(1000)
	ControlSetText($wndSaveas,"","Edit1",$fileDir)
	Sleep(1000)
	ControlClick($wndSaveas,"","Button6")
	;WinClose($hWnd)
	_Excel_Close($excelProcess)
EndFunc
;create_note(@DesktopDir & "\ss.txt")
#cs-----------------------------------------------------------------------------------------------------------------------
;创建txt文件
parameters:
$fileDir:  保存文件的路径
#ce-----------------------------------------------------------------------------------------------------------------------
Func create_note($fileDir)
	If FileExists($fileDir) Then
		FileDelete($fileDir)
	EndIf
	Run("notepad.exe")
	Local $hWnd = WinWaitActive("[CLASS:Notepad]","",10)
	;ControlFocus($hWnd,"","")
	Sleep(1000)
	Send("^s")
	Sleep(1000)
	Local $wndSaveas = WinWaitActive("另存为","",10)
	Sleep(1000)
	ControlSetText($wndSaveas,"","Edit1",$fileDir)
	Sleep(1000)
	ControlClick($wndSaveas,"","Button1")
	WinClose($hWnd)
EndFunc
#ce




