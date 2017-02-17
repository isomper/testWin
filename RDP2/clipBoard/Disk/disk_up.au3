#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2017/02/13
作者：顾亚茹
功能说明：磁盘内部（上行）业务数据
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------
#include <Array.au3>
#include <WinAPIFiles.au3>
#include "C:\au3\disk_clipboard.au3"



$path = "\\tsclient\D\au3"
;定义txt存放路径
$txt =$path & "\cs.txt"
;定义txt的title标题
$txtTitle = "cs.txt - 记事本"
;定义txt中做剪切操作关闭窗口的title标题
$txtCloseTitle = "记事本"

;定义word存放路径
$word =$path & "\Cword.docx"
;定义word的title标题，
$wordTitle = "Cword.docx - Microsoft Word"
;定义word中做剪切操作关闭窗口的title标题
$wordCloseTitle = "Microsoft Office Word"

;定义excel存放路径
$excel =$path & "\Cexcel.xlsx"
;定义excel的title标题
$excelTitle = "Microsoft Excel - Cexcel.xlsx"
;定义excel中做剪切操作关闭窗口的title标题
$excelCloseTitle = "Microsoft Office Excel"

;定义txt中按几下tab键关闭Ctrl+f窗口，定义word中按几下tab键关闭Ctrl+f窗口，定义excel中按几下tab键关闭Ctrl+f窗口
$txtTabCount = 0
$wordTabCount = 1
$excelTabCount = 2

;文件复制的方式，0代表右键，1代表组合键
$opMethod_copyMenu = 0
$opMethod_copyHotKey = 1

;定义txt右键按3下到复制键，
$txtCopy = 3
;txt右键按2下到剪切键
$txtCut = 2
;txt右键按4下到粘贴建
$txtPate = 4
;word右键按1下到复制键
$wordCopy = 1
;word右键按0下到剪切键
$wordCut = 0
;word右键按2下到粘贴建
$wordPaste = 2
;excel右键按1下到复制键
$excelCopy = 1
;excel右键按0下到剪切键
$excelCut = 0
;excel右键按2下到粘贴建
$excelPaste = 2

;定义组合键
$copy = "c"
$cut = "x"
$paste = "v"

;定义远程服务器txt存放路径
$remoteTxt = "c:\au3\disk.txt"
;定义远程服务器txt的title
$remoteTxtTitle = "disk.txt - 记事本"

;定义远程服务器word存放路径，
$remoteWord = "c:\au3\disk.docx"
;定义远程服务器word的title
$remoteWordTitle = "disk.docx - Microsoft Word"

;定义远程服务器excel存放路径
$remoteExcel = "c:\au3\disk.xlsx"
;定义远程服务器excel的title
$remoteExcelTitle = "Microsoft Excel - disk.xlsx"

;定义文件粘贴的方式，0代表右键，1代表组合键
$opMethod_pasteMenu = 0
$opMethod_pasteHotKey = 1

;字母含义说明：
;数字--> 业务序号
;w-->word           	e-->excel            	t-->txt
;f-->右键复制           	 	z-->右键粘贴           	 		j-->右键剪切
;c-->ctrl +c快捷键       	x-->ctrl +x快捷键     	 	v-->ctrl +v快捷键
;eg:00001-wwrfz 代表 业务代码为00001，右键复制-远程右键粘贴word-word的操作

Local $key = ["00001-wwrfz","00002-werfz","00003-wtrfz","00004-ttrfz","00005-twrfz","00006-terfz","00007-eerfz","00008-etrfz","00009-ewrfz","000010-wwrfv","000011-werfv","000012-wtrfv","000013-ttrfv","000014-twrfv","000015-terfv","000016-eerfv","000017-etrfv","000018-ewrfv","000019-wwcrz","000020-wecrz","000021-wtcrz","000022-ttcrz","000023-twcrz","000024-tecrz","000025-eecrz","000026-etcrz","000027-ewcrz","000028-wwcv","000029-wecv","000030-wtcv","000031-ttcv","000032-twcv","000033-tecv","000034-eecv","000035-etcv","000036-ewcv","000037-wwrjz","000038-werjz","000039-wtrjz","000040-ttrjz","000041-twrjz","000042-terjz","000043-eerjz","000044-etrjz","000045-ewrjz","000046-wwrjv","000047-werjv","000048-wtrjv","000049-ttrjv","000050-twrjv","000051-terjv","000052-eerjv","000053-etrjv","000054-ewrjv","000055-wwxrz","000056-wexrz","000057-wtxrz","000058-ttxrz","000059-twxrz","000060-texrz","000061-eexrz","000062-etxrz","000063-ewxrz","000064-wwxv","000065-wexv","000066-wtxv","000067-ttxv","000068-twxv","000069-texv","000070-eexv","000071-etxv","000072-ewxv"]
Local $opKind = [""]

For $j=0 To UBound($key) - 1
	Local $arrayKey = StringSplit($key[$j], "-")
	If $arrayKey[2] == "wwrfz" Then;00001 右键复制-远程右键粘贴word-word
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$wordTitle,$wordCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werfz" Then;00002 右键复制-远程右键粘贴word-excel
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$wordTitle,$wordCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrfz" Then;00003 右键复制-远程右键粘贴word-txt
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$wordTitle,$wordCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttrfz" Then;00004 右键复制-远程右键粘贴txt-txt
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$txtTitle,$txtCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrfz" Then;00005 右键复制-远程右键粘贴txt-word
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$txtTitle,$txtCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terfz" Then;00006 右键复制-远程右键粘贴txt-excel
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$txtTitle,$txtCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eerfz" Then;00007  右键复制-远程右键粘贴excel-excel
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$excelTitle,$excelCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrfz" Then;00008  右键复制-远程右键粘贴excel-txt
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$excelTitle,$excelCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrfz" Then;00009 右键复制-远程右键粘贴excel-word
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$excelTitle,$excelCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]

	ElseIf $arrayKey[2] == "wwrfv" Then;000010 右键复制-远程ctrl+v word-word
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$wordTitle,$wordCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werfv" Then;000011 右键复制-远程ctrl+v word-excel
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$wordTitle,$wordCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrfv" Then;000012 右键复制-远程ctrl+v word-txt
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$wordTitle,$wordCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttrfv" Then;000013 右键复制-远程ctrl+v txt-txt
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$txtTitle,$txtCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrfv" Then;000014 右键复制-远程ctrl+v txt-word
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$txtTitle,$txtCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terfv" Then;000015 右键复制-远程ctrl+v txt-excel
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$txtTitle,$txtCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eerfv" Then;000016 右键复制-远程ctrl+v excel-excel
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$excelTitle,$excelCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrfv" Then;000017 右键复制-远程ctrl+v excel-txt
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$excelTitle,$excelCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrfv" Then;000018 右键复制-远程ctrl+v excel-word
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$excelTitle,$excelCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]

	ElseIf $arrayKey[2] == "wwcrz" Then;000019 ctrl+c-远程右键粘贴word-word
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$wordTitle,$wordCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wecrz" Then;000020 ctrl+c-远程右键粘贴word-excel
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$wordTitle,$wordCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtcrz" Then;000021 ctrl+c-远程右键粘贴word-txt
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$wordTitle,$wordCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttcrz" Then;000022 ctrl+c-远程右键粘贴txt-txt
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$txtTitle,$txtCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twcrz" Then;000023 ctrl+c-远程右键粘贴txt-word
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$txtTitle,$txtCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "tecrz" Then;000024 ctrl+c-远程右键粘贴txt-excel
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$txtTitle,$txtCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eecrz" Then;000025 ctrl+c-远程右键粘贴excel-excel
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$excelTitle,$excelCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etcrz" Then;000026 ctrl+c-远程右键粘贴excel-txt
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$excelTitle,$excelCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewcrz" Then;000027 ctrl+c-远程右键粘贴excel-word
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$excelTitle,$excelCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]

	ElseIf $arrayKey[2] == "wwcv" Then;000028 ctrl+c-远程ctrl+v word-word
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$wordTitle,$wordCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wecv" Then;000029 ctrl+c-远程ctrl+v word-excel
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$wordTitle,$wordCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtcv" Then;000030 ctrl+c-远程ctrl+v word-txt
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$wordTitle,$wordCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttcv" Then;000031 ctrl+c-远程ctrl+v txt-txt
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$txtTitle,$txtCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twcv" Then;000032 ctrl+c-远程ctrl+v txt-word
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$txtTitle,$txtCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "tecv" Then;000033 ctrl+c-远程ctrl+v txt-excel
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$txtTitle,$txtCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eecv" Then;000034 ctrl+c-远程ctrl+v excel-excel
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$excelTitle,$excelCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etcv" Then;000035 ctrl+c-远程ctrl+v excel-txt
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$excelTitle,$excelCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewcv" Then;000036 ctrl+c-远程ctrl+v excel-word
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$excelTitle,$excelCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]

	ElseIf $arrayKey[2] == "wwrjz" Then;000037 右键剪切-远程右键粘贴word-word
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$wordTitle,$wordCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werjz" Then;000038 右键剪切-远程右键粘贴word-excel
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$wordTitle,$wordCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrjz" Then;000039 右键剪切-远程右键粘贴word-txt
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$wordTitle,$wordCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttrjz" Then;000040 右键剪切-远程右键粘贴txt-txt
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$txtTitle,$txtCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrjz" Then;000041 右键剪切-远程右键粘贴txt-word
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$txtTitle,$txtCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terjz" Then;000042 右键剪切-远程右键粘贴txt-excel
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$txtTitle,$txtCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eerjz" Then;000043 右键剪切-远程右键粘贴excel-excel
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$excelTitle,$excelCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrjz" Then;000044 右键剪切-远程右键粘贴excel-txt
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$excelTitle,$excelCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrjz" Then;000045 右键剪切-远程右键粘贴excel-word
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$excelTitle,$excelCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]

	ElseIf $arrayKey[2] == "wwrjv" Then;000046 右键剪切-远程ctrl+v word-word
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$wordTitle,$wordCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werjv" Then;000047 右键剪切-远程ctrl+v word-excel
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$wordTitle,$wordCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrjv" Then;000048 右键剪切-远程ctrl+v word-txt
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$wordTitle,$wordCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttrjv" Then;000049 右键剪切-远程ctrl+v txt-txt
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$txtTitle,$txtCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrjv" Then;000050 右键剪切远程ctrl+v txt-word
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$txtTitle,$txtCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terjv" Then;000051 右键剪切-远程ctrl+v txt-excel
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$txtTitle,$txtCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eerjv" Then;000052 右键剪切-远程ctrl+v excel-excel
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$excelTitle,$excelCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrjv" Then;000053 右键剪切-远程ctrl+v excel-txt
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$excelTitle,$excelCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrjv" Then;000054 右键剪切-远程ctrl+v excel-word
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$excelTitle,$excelCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]


	ElseIf $arrayKey[2] == "wwxrz" Then;000055 本地ctrl+x-远程右键粘贴word-word
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$wordTitle,$wordCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wexrz" Then;000056 本地ctrl+x-远程右键粘贴word-excel
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$wordTitle,$wordCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtxrz" Then;000057 本地ctrl+x-远程右键粘贴word-txt
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$wordTitle,$wordCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttxrz" Then;000058 本地ctrl+x-远程右键粘贴txt-txt
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$txtTitle,$txtCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twxrz" Then;000059 本地ctrl+x-远程右键粘贴txt-word
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$txtTitle,$txtCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "texrz" Then;000060 本地ctrl+x-远程右键粘贴txt-excel
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$txtTitle,$txtCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eexrz" Then;000061 本地ctrl+x-远程右键粘贴excel-excel
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$excelTitle,$excelCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteMenu,$excelPaste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etxrz" Then;000062 本地ctrl+x-远程右键粘贴excel-txt
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$excelTitle,$excelCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteMenu,$txtPate,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewxrz" Then;000063 本地ctrl+x-远程右键粘贴excel-word
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$excelTitle,$excelCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteMenu,$wordPaste,$remoteWordTitle,$wordCloseTitle]

	ElseIf $arrayKey[2] == "wwxv" Then;000064 本地ctrl+x-远程ctrl+v word-word
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$wordTitle,$wordCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wexv" Then;000065 本地ctrl+x-远程ctrl+v word-excel
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$wordTitle,$wordCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtxv" Then;000066 本地ctrl+x-远程ctrl+v word-txt
		Local $opKind =  [$word,$wordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$wordTitle,$wordCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttxv" Then;000067 本地ctrl+x-远程ctrl+v txt-txt
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$txtTitle,$txtCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twxv" Then;000068 本地ctrl+x-远程ctrl+v txt-word
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$txtTitle,$txtCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "texv" Then;000069 本地ctrl+x-远程ctrl+v txt-excel
		Local $opKind =  [$txt,$txtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$txtTitle,$txtCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eexv" Then;000070 本地ctrl+x-远程ctrl+v excel-excel
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$excelTitle,$excelCloseTitle,$remoteExcel,$remoteExcelTitle,$opMethod_pasteHotKey,$paste,$remoteExcelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etxv" Then;000071 本地ctrl+x-远程ctrl+v excel-txt
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$excelTitle,$excelCloseTitle,$remoteTxt,$remoteTxtTitle,$opMethod_pasteHotKey,$paste,$remoteTxtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewxv" Then;000072 本地ctrl+x-远程ctrl+v excel-word
		Local $opKind =  [$excel,$excelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$excelTitle,$excelCloseTitle,$remoteWord,$remoteWordTitle,$opMethod_pasteHotKey,$paste,$remoteWordTitle,$wordCloseTitle]
	EndIf

	execute_cmd($opKind)

Next