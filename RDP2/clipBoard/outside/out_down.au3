#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2017/02/13
作者：张泽方
功能说明：磁盘外部下行剪切板业务数据
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include <Array.au3>

#include "C:\Users\zzf\Downloads\Desktop\clipBoard\outside\outside_clipboard.au3"

#cs-----------------------------------------------------------------------------------------------------------------------
执行远程文件到本地文件的剪切复制粘贴操作
parameters:
$opKind1: 操作类型 （eg:远程txt右键复制 ->本地txt的右键粘贴动作类型）
#ce-----------------------------------------------------------------------------------------------------------------------
Func execute_cmd($opKind1)
	$opTmp = $opKind1
	writeTmpTxt($opTmp[0] & " " & $opTmp[1]& " " & $opTmp[2]& " " & $opTmp[3])
	;Sleep(3000)
	remoteOp($opTmp[4])
	Sleep(1000)
	openFile($opTmp[5], $opTmp[6])
	If $opTmp[7] == 0 Then
		menuOp_opString($opTmp[8])
	ElseIf $opTmp[7] == 1 Then
		hotKey_opString($opTmp[8])
	EndIf
	close_windows($opTmp[6], $opTmp[9])
	Sleep(1000)
EndFunc

;远程服务器文件名字
$remoteTxt = "c:\au3\test.txt"
$remoteWord = "c:\au3\test.docx"
$remoteExcel = "c:\au3\test.xlsx"

;远程文件复制或者剪切的方式，0代表右键，1代表组合键 $opMethod
$opMethod_menu = 0
$opMethod_hotKey = 1

;远程服务器 ，0代表右键复制，1代表右键剪切，2代表组合键ctrl +c 复制，3代表组合键ctrl +x 剪切$opStatus
$opMenu_copyStatus = 0
$opMenu_cutStatus = 1
$hotKey_copyStatus = 2
$hotKey_cutStatus = 3


$path = "D:\au3\"
;定义txt存放路径
$txt = $path & "cs.txt"
;定义txt的title标题
$txtTitle = "cs.txt - 记事本"
;定义txt中做剪切操作关闭窗口的title标题
$txtCloseTitle = "记事本"

;定义word存放路径
$word = $path & "Cword.docx"
;定义word的title标题
$wordTitle = "Cword.docx - Microsoft Word"
;定义word中做剪切操作关闭窗口的title标题
$wordCloseTitle = "Microsoft Office Word"

;定义excel存放路径
$excel = $path & "Cexcel.xlsx"
;定义excel的title标题
$excelTitle = "Microsoft Excel - Cexcel.xlsx"
;定义excel中做剪切操作关闭窗口的title标题
$excelCloseTitle = "Microsoft Office Excel"

;使用sikuli打开远程服务器中的文件夹
$openCmd = $path & "server2003Copy.sikuli"

;定义txt右键按4下到粘贴键.
$txtPaste = 4
;定义word右键按1下到粘贴键.
$wordPaste = 2
;定义excel右键按1下到粘贴键.
$excelPaste = 2

;定义粘贴快捷键
$hotKeyPaste = "v"



;0代表使用右键粘贴方式，1代表使用快捷键粘贴
$menuState = 0
$hotKeyState = 1

;字母含义说明：
;数字--> 业务序号
;w-->word           	e-->excel            	t-->txt
;f-->右键复制           	 	z-->右键粘贴           	 		j-->右键剪切
;c-->ctrl +c快捷键       	x-->ctrl +x快捷键     	 	v-->ctrl +v快捷键
;eg:0001-wwrfz 代表 业务代码为0001，右键复制-远程右键粘贴word-word的操作

Local $key =  ["0001-wwrfz","0002-werfz","0003-wtrfz","0004-ttrfz","0005-twrfz","0006-terfz","0007-eerfz","0008-etrfz","0009-ewrfz","0010-wwrfv","0011-werfv","0012-wtrfv","0013-ttrfv","0014-twrfv","0015-terfv","0016-eerfv","0017-etrfv","0018-ewrfv","0019-wwcrz","0020-wecrz","0021-wtcrz","0022-ttcrz","0023-twcrz","0024-tecrz","0025-eecrz","0026-etcrz","0027-ewcrz","0028-wwcv","0029-wecv","0030-wtcv","0031-ttcv","0032-twcv","0033-tecv","0034-eecv","0035-etcv","0036-ewcv","0037-wwrjz","0038-werjz","0039-wtrjz","0040-ttrjz","0041-twrjz","0042-terjz","0043-eerjz","0044-etrjz","0045-ewrjz","0046-wwrjv","0047-werjv","0048-wtrjv","0049-ttrjv","0050-twrjv","0051-terjv","0052-eerjv","0053-etrjv","0054-ewrjv","0055-wwxrz","0056-wexrz","0057-wtxrz","0058-ttxrz","0059-twxrz","0060-texrz","0061-eexrz","0062-etxrz","0063-ewxrz","0064-wwxv","0065-wexv","0066-wtxv","0067-ttxv","0068-twxv","0069-texv","0070-eexv","0071-etxv","0072-ewxv"]
Local $opKind = [""]
For $j=0 To UBound($key) - 1
	Local $arrayKey = StringSplit($key[$j], "-")
	If $arrayKey[2] == "ttrfz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrfz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terfz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "ttcrz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twcrz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "tecrz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "ttrjz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrjz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terjz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "ttxrz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twxrz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "texrz" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "ttcv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twcv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "tecv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "ttxv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twxv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "texv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "ttrfv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrfv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terfv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "ttrjv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrjv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terjv" Then
		Local $opKind = [$remoteTxt,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wwrfz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werfz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrfz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "wwrjz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werjz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrjz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "wwrfv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werfv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrfv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "wwcv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wecv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtcv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "wwxv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wexv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtxv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "wwrjv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werjv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrjv" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "wwxrz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wexrz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtxrz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "wwcrz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wecrz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtcrz" Then
		Local $opKind = [$remoteWord,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "eerfz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrfz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrfz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "eerjz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrjz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrjz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "eerjv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrjv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrjv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_cutStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "eerfv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrfv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrfv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_menu,$opMenu_copyStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "eexv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etxv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewxv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "eecv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$excel,$excelTitle,$hotKeyState,$hotKeyPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etcv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$txt,$txtTitle,$hotKeyState,$hotKeyPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewcv" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$word,$wordTitle,$hotKeyState,$hotKeyPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "eexrz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etxrz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewxrz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_cutStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	ElseIf $arrayKey[2] == "eecrz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$excel,$excelTitle,$menuState,$excelPaste,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etcrz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$txt,$txtTitle,$menuState,$txtPaste,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewcrz" Then
		Local $opKind = [$remoteExcel,$arrayKey[1],$opMethod_hotKey,$hotKey_copyStatus,$openCmd,$word,$wordTitle,$menuState,$wordPaste,$wordCloseTitle]
	EndIf

	execute_cmd($opKind)
Next