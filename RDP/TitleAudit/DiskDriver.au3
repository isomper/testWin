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
;;创建word文档
parameters:
$fileDir:  保存文件的路径
#ce-----------------------------------------------------------------------------------------------------------------------
Func create_word($fileDir)
	If FileExists($fileDir) Then
		FileDelete($fileDir )
	EndIf
	Local $wordProcess = _Word_Create()
	Local $wordDoc = _Word_DocAdd($wordProcess)
	Local $oRange = _Word_DocRangeSet($wordDoc, -1);焦点放在最近开始处
	$oRange.Text = "这是标题. "
	_Word_DocSaveAs($wordDoc, $fileDir);保存在桌面
	Sleep(2000)
	_Word_Quit($wordProcess)
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
	_Excel_BookSaveAs  ($oexcel,$fileDir,$xlWorkbookDefault , True)
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
	ControlClick($wndSaveas,"","Button2")
	WinClose($hWnd)
EndFunc
#cs-----------------------------------------------------------------------------------------------------------------------
;创建文件夹
parameters:
$fileDir:  保存文件的路径
#ce-----------------------------------------------------------------------------------------------------------------------
Func create_folder($fileDir)
	DirCreate($fileDir)
EndFunc
;move_file("C:\Users\guyaru.pc\Desktop\test.doc","D:\test.doc","E:\test.doc")
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
;clip_board("C:\Users\guyaru.pc\Desktop\test.doc","test.doc [兼容模式] - Word","C:\Users\guyaru.pc\Desktop\test.txt","test.txt - 记事本",2)
#cs-----------------------------------------------------------------------------------------------------------------------
;复制、剪切、删除文件和文件夹
parameters:
$sourceFile:: 源文件         $sourceTitle:源文件标题     $destFile:目标文件    $destTitle:目标文件    $state:状态
#ce-----------------------------------------------------------------------------------------------------------------------
Func clip_board($sourceFile,$sourceTitle,$destFile,$destTitle,$state)
	Local $sPid = ShellExecute($sourceFile)
	WinWaitActive($sourceTitle,"",10)
	$text = "测试内容"
	ClipPut($text)
	If $state ==1 Then
		Send("^V")
		Send("!{f4}")
	Else
		Local $dPid = ShellExecute($destFile)
		WinWaitActive($destTitle,"",2)
		Send("^V")
		Sleep(500)
		Send("!{f4}")
	EndIf
	WinClose($sPid)
	;ConsoleWrite($sPid &@CRLF)
	;WinClose($dPid)
EndFunc




