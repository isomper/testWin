#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2017/02/13
作者：张泽方
功能说明：磁盘外部上行剪切板业务数据
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include <Array.au3>

#include "C:\Users\zzf\Downloads\Desktop\clipBoard\outside\outside_clipboard.au3"
;#include "C:\Users\guyaru.pc\Desktop\en\clipBoard\outside\outside_clipboard.au3"
#cs-----------------------------------------------------------------------------------------------------------------------
执行本地文件到远程文件的剪切复制粘贴操作
parameters:
$opKind1: 操作类型 （eg:本地txt右键复制 ->远程txt的右键粘贴动作类型）
#ce-----------------------------------------------------------------------------------------------------------------------
Func execute_cmd($opKind1)
	$opTmp = $opKind1
	openFile($opTmp[0], $opTmp[1])
	searchString($opTmp[2],$opTmp[3])
	If $opTmp[4] == 0 Then
		menuOp_opString($opTmp[5])
	ElseIf $opTmp[4] == 1 Then
		hotKey_opString($opTmp[5])
	EndIf
	close_windows($opTmp[6],$opTmp[7])
	Sleep(1000)
	writeTmpTxt($opTmp[9] & " " & $opTmp[10]& " " & $opTmp[11])
	remoteOp($opTmp[8])
	Sleep(1000)

	;Sleep(9000)
EndFunc


;远程服务器文件名字
$remoteTxt = "c:\au3\test.txt"
$remoteWord = "c:\au3\test.docx"
$remoteExcel = "c:\au3\test.xlsx"

;远程文件粘贴的方式，0代表右键，1代表组合键 $opMethod
$opMethod_menu = 0
$opMethod_hotKey = 1

;远程服务器 ，0代表右键粘贴，1代表组合键ctrl +v 粘贴
$opMenu_pasteStatus = 0
$hotKey_pasteStatus = 1


$path = "D:\au3\"
;定义本地txt存放路径
$txt = $path & "cs.txt"
;定义txt的title标题
$txtTitle = "cs.txt - 记事本"
;定义txt中做剪切操作关闭窗口的title标题
$txtCloseTitle = "记事本"

;定义本地word存放路径
$word =  $path & "Cword.docx"
;定义word的title标题
$wordTitle = "Cword.docx - Microsoft Word"
;定义word中做剪切操作关闭窗口的title标题
$wordCloseTitle = "Microsoft Office Word"

;定义本地excel存放路径
$excel = $path & "Cexcel.xlsx"
;定义excel的title标题
$excelTitle = "Microsoft Excel - Cexcel.xlsx"
;定义excel中做剪切操作关闭窗口的title标题
$excelCloseTitle = "Microsoft Office Excel"

;使用sikuli打开远程服务器中的文件夹
;$openExeCmd = $path  & "\remotePaste1734.sikuli"
$openExeCmd = $path  & "server2003Paste.sikuli"
;$openExeCmd = $path  & "server2012Paste.sikuli"



;定义txt中按几下tab键关闭Ctrl+f窗口，定义word中按几下tab键关闭Ctrl+f窗口，定义excel中按几下tab键关闭Ctrl+f窗口
$txtTabCount = 0
$wordTabCount = 1
$excelTabCount = 2

;定义txt右键按3下到复制键
$txtCopy = 3
;txt右键按2下到剪切键
$txtCut = 2
;word右键按1下到复制键
$wordCopy = 1
;word右键按0下到剪切键
$wordCut = 0
;excel右键按1下到复制键
$excelCopy = 1
;excel右键按0下到剪切键
$excelCut = 0

;0代表使用右键粘贴方式，1代表使用快捷键粘贴
$menuState = 0
$hotKeyState = 1

;定义组合键
$copy = "c"
$cut = "x"

;字母含义说明：
;数字--> 业务序号
;w-->word           	e-->excel            	t-->txt
;f-->右键复制           	 	z-->右键粘贴           	 		j-->右键剪切
;c-->ctrl +c快捷键       	x-->ctrl +x快捷键     	 	v-->ctrl +v快捷键
;eg:001-wwrfz 代表 业务代码为001，右键复制-远程右键粘贴word-word的操作
Local $key = ["001-wwrfz","002-werfz","003-wtrfz","004-ttrfz","005-twrfz","006-terfz","007-eerfz","008-etrfz","009-ewrfz","010-wwrfv","011-werfv","012-wtrfv","013-ttrfv","014-twrfv","015-terfv","016-eerfv","017-etrfv","018-ewrfv","019-wwcrz","020-wecrz","021-wtcrz","022-ttcrz","023-twcrz","024-tecrz","025-eecrz","026-etcrz","027-ewcrz","028-wwcv","029-wecv","030-wtcv","031-ttcv","032-twcv","033-tecv","034-eecv","035-etcv","036-ewcv","037-wwrjz","038-werjz","039-wtrjz","040-ttrjz","041-twrjz","042-terjz","043-eerjz","044-etrjz","045-ewrjz","046-wwrjv","047-werjv","048-wtrjv","049-ttrjv","050-twrjv","051-terjv","052-eerjv","053-etrjv","054-ewrjv","055-wwxrz","056-wexrz","057-wtxrz","058-ttxrz","059-twxrz","060-texrz","061-eexrz","062-etxrz","063-ewxrz","064-wwxv","065-wexv","066-wtxv","067-ttxv","068-twxv","069-texv","070-eexv","071-etxv","072-ewxv"]
For $j=0 To UBound($key) - 1
	Local $arrayKey = StringSplit($key[$j], "-")
	If $arrayKey[2] == "wwrfz" Then
		Local $opKind = [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCopy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "werfz" Then
		Local $opKind = [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCopy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wtrfz" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCopy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "ttrfz" Then
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$menuState,$txtCopy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "twrfz" Then
		Local $opKind = [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$menuState,$txtCopy,$txtTitle,$txtCloseTitle,$openExeCmd, $remoteWord,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "terfz" Then
		Local $opKind = [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$menuState,$txtCopy,$txtTitle,$txtCloseTitle,$openExeCmd, $remoteExcel,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "eerfz" Then
		Local $opKind = [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCopy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "etrfz" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCopy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "ewrfz" Then
		Local $opKind = [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCopy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wwrfv" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCopy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "werfv" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCopy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "wtrfv" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCopy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "ttrfv" Then
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$menuState,$txtCopy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "twrfv" Then
		Local $opKind = [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$menuState,$txtCopy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "terfv" Then
		Local $opKind = [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$menuState,$txtCopy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "eerfv" Then
		Local $opKind = [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCopy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "etrfv" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCopy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "ewrfv" Then
		Local $opKind = [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCopy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "wwcrz" Then
		Local $opKind =   [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$copy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wecrz" Then
		Local $opKind =   [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$copy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wtcrz" Then
		Local $opKind =   [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$copy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "ttcrz" Then
		Local $opKind =   [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$copy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "twcrz" Then
		Local $opKind =   [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$copy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "tecrz" Then
		Local $opKind =   [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$copy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "eecrz" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$copy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "ewcrz" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$copy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "etcrz" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$copy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wwcv" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$copy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "wecv" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$copy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "wtcv" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$copy,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "ttcv" Then
		Local $opKind =  [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$copy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "twcv" Then
		Local $opKind =  [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$copy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "tecv" Then
		Local $opKind =  [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$copy,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "eecv" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$copy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "etcv" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$copy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "ewcv" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$copy,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "wwrjz" Then
		Local $opKind = [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "werjz" Then
		Local $opKind = [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wtrjz" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "ttrjz" Then
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$menuState,$txtCut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "twrjz" Then
		Local $opKind = [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$menuState,$txtCut,$txtTitle,$txtCloseTitle,$openExeCmd, $remoteWord,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "terjz" Then
		Local $opKind = [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$menuState,$txtCut,$txtTitle,$txtCloseTitle,$openExeCmd, $remoteExcel,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "eerjz" Then
		Local $opKind = [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "etrjz" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "ewrjz" Then
		Local $opKind = [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wwrjv" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "werjv" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "wtrjv" Then
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$menuState,$wordCut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "ttrjv" Then
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$menuState,$txtCut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "twrjv" Then
		Local $opKind = [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$menuState,$txtCut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "terjv" Then
		Local $opKind = [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$menuState,$txtCut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "eerjv" Then
		Local $opKind = [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "etrjv" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "ewrjv" Then
		Local $opKind = [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$menuState,$excelCut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "wwxrz" Then
		Local $opKind =   [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$cut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wexrz" Then
		Local $opKind =   [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$cut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wtxrz" Then
		Local $opKind =   [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$cut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "ttxrz" Then
		Local $opKind =   [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$cut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "twxrz" Then
		Local $opKind =   [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$cut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "texrz" Then
		Local $opKind =   [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$cut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu,$opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "eexrz" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$cut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteExcel,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "etxrz" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$cut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteTxt,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "ewxrz" Then
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$cut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteWord,$opMethod_menu, $opMenu_pasteStatus]
	ElseIf $arrayKey[2] == "wwxv" Then ;064
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$cut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "wexv" Then ;065
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$cut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "wtxv" Then ;066
		Local $opKind =  [$word, $wordTitle, $wordTabCount,$arrayKey[1],$hotKeyState,$cut,$wordTitle,$wordCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "ttxv" Then ;067
		Local $opKind =  [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$cut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "twxv" Then ;068
		Local $opKind =  [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$cut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "texv" Then ;069
		Local $opKind =  [$txt, $txtTitle, $txtTabCount,$arrayKey[1],$hotKeyState,$cut,$txtTitle,$txtCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey,$hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "eexv" Then ;070
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$cut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteExcel,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "etxv" Then ;071
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$cut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteTxt,$opMethod_hotKey, $hotKey_pasteStatus]
	ElseIf $arrayKey[2] == "ewxv" Then ;072
		Local $opKind =  [$excel, $excelTitle, $excelTabCount,$arrayKey[1],$hotKeyState,$cut,$excelTitle,$excelCloseTitle,$openExeCmd,$remoteWord,$opMethod_hotKey, $hotKey_pasteStatus]



	EndIf

	execute_cmd($opKind)

Next




