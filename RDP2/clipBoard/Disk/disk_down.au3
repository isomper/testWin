#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2017/02/13
作者：张泽方
功能说明：磁盘内部（下行）业务数据
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include "c:\au3\disk_clipboard.au3"

;定义远程服务器txt存放路径，txt的title
$remoteTxt = "c:\au3\disk.txt"
$remoteTxtTitle = "disk.txt - 记事本"

;定义远程服务器word存放路径，word的title
$remoteWord = "c:\au3\disk.docx"
$remoteWordTitle = "disk.docx - Microsoft Word"

;定义远程服务器excel存放路径，excel的title
$remoteExcel = "c:\au3\disk.xlsx"
$remoteExcelTitle = "Microsoft Excel - disk.xlsx"

;定义txt中按几下tab键关闭Ctrl+f窗口,定义word中按几下tab键关闭Ctrl+f窗口,定义excel中按几下tab键关闭Ctrl+f窗口
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


$path = "\\tsclient\D\au3"
;定义txt存放路径
$txt = $path & "\cs.txt"
;定义txt的title标题
$txtTitle = "cs.txt - 记事本"
;定义txt中做剪切操作关闭窗口的title标题
$txtCloseTitle = "记事本"

;定义word存放路径
$word = $path & "\Cword.docx"
;定义word的title标题
$wordTitle = "Cword.docx - Microsoft Word"
;定义word中做剪切操作关闭窗口的title标题
$wordCloseTitle = "Microsoft Office Word"

;定义excel存放路径
$excel =$path & "\Cexcel.xlsx"
;定义excel的title标题
$excelTitle = "Microsoft Excel - Cexcel.xlsx"
;定义excel中做剪切操作关闭窗口的title标题
$excelCloseTitle = "Microsoft Office Excel"

;文件粘贴的方式，0代表右键，1代表组合键
$opMethod_pasteMenu = 0
$opMethod_pasteHotKey = 1

;字母含义说明：
;数字--> 业务序号
;w-->word           	e-->excel            	t-->txt
;f-->右键复制           	 	z-->右键粘贴           	 		j-->右键剪切
;c-->ctrl +c快捷键       	x-->ctrl +x快捷键     	 	v-->ctrl +v快捷键
;eg:00080-wwrfz 代表 业务代码为00080,右键复制-远程右键粘贴word-word的操作

Local $key = ["00080-wwrfz","00081-werfz","00082-wtrfz","00083-ttrfz","00084-twrfz","00085-terfz","00086-eerfz","00087-etrfz","00088-ewrfz","00089-wwrjz","00090-werjz","00091-wtrjz","00092-ttrjz","00093-twrjz","00094-terjz","00095-eerjz","00096-etrjz","00097-ewrjz","00098-wwrfv","00099-werfv","00100-wtrfv","00101-ttrfv","00102-twrfv","00103-terfv","00104-eerfv","00105-etrfv","00106-ewrfv","00107-wwrjv","00108-werjv","00109-wtrjv","00110-ttrjv","00111-twrjv","00112-terjv","00113-eerjv","00114-etrjv","00115-ewrjv","00116-wwcrz","00117-wecrz","00118-wtcrz","00119-ttcrz","00120-twcrz","00121-tecrz","00122-eecrz","00123-etcrz","00124-ewcrz","00125-wwxrz","00126-wexrz","00127-wtxrz","00128-ttxrz","00129-twxrz","00130-texrz","00131-eexrz","00132-etxrz","00133-ewxrz","00134-wwcv","00135-wecv","00136-wtcv","00137-ttcv","00138-twcv","00139-tecv","00140-eecv","00141-etcv","00142-ewcv","00143-wwxv","00144-wexv","00145-wtxv","00146-ttxv","00147-twxv","00148-texv","00149-eexv","00150-etxv","00151-ewxv"]

For $j=0 To UBound($key) - 1
	Local $arrayKey = StringSplit($key[$j], "-")
	If $arrayKey[2] == "wwrfz" Then ;00080 右键复制-远程右键粘贴word-word
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$remoteWordTitle,$wordCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werfz" Then;00081 右键复制-远程右键粘贴word-excel
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$remoteWordTitle,$wordCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrfz" Then;00082 右键复制-远程右键粘贴word-txt
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$remoteWordTitle,$wordCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttrfz" Then;00083 右键复制-远程右键粘贴txt-txt
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$remoteTxtTitle,$txtCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrfz" Then;00084 右键复制-远程右键粘贴txt-word
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$remoteTxtTitle,$txtCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terfz" Then;00085 右键复制-远程右键粘贴txt-excel
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$remoteTxtTitle,$txtCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eerfz" Then;00086  右键复制-远程右键粘贴excel-excel
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$remoteExcelTitle,$excelCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrfz" Then;00087  右键复制-远程右键粘贴excel-txt
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$remoteExcelTitle,$excelCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrfz" Then;00088 右键复制-远程右键粘贴excel-word
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$remoteExcelTitle,$excelCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wwrjz" Then ;00089 右键剪切-远程右键粘贴word-word
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$remoteWordTitle,$wordCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werjz" Then;00090 右键剪切-远程右键粘贴word-excel
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$remoteWordTitle,$wordCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrjz" Then;00091 右键剪切-远程右键粘贴word-txt
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$remoteWordTitle,$wordCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttrjz" Then;00092 右键剪切-远程右键粘贴txt-txt
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$remoteTxtTitle,$txtCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrjz" Then;00093 右键剪切-远程右键粘贴txt-word
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$remoteTxtTitle,$txtCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terjz" Then;00094 右键剪切-远程右键粘贴txt-excel
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$remoteTxtTitle,$txtCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eerjz" Then;00095  右键剪切-远程右键粘贴excel-excel
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$remoteExcelTitle,$excelCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrjz" Then;00096  右键剪切-远程右键粘贴excel-txt
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$remoteExcelTitle,$excelCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrjz" Then;00097 右键剪切-远程右键粘贴excel-word
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$remoteExcelTitle,$excelCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wwrfv" Then ;00098右键复制-远程ctrl+v word-word
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$remoteWordTitle,$wordCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werfv" Then;00099 右键复制-远程ctrl+v word-excel
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$remoteWordTitle,$wordCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrfv" Then;00100 右键复制-远程ctrl+v word-txt
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCopy,$remoteWordTitle,$wordCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttrfv" Then;0010   右键复制-远程ctrl+v txt-txt
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$remoteTxtTitle,$txtCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrfv" Then;00102   右键复制-远程ctrl+v txt-word
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$remoteTxtTitle,$txtCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terfv" Then;00103 右键复制-远程ctrl+v txt-excel
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCopy,$remoteTxtTitle,$txtCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eerfv" Then;00104  右键复制-远程ctrl+v excel-excel
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$remoteExcelTitle,$excelCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrfv" Then;00105   右键复制-远程ctrl+v excel-txt
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$remoteExcelTitle,$excelCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrfv" Then;00106 右键复制-远程ctrl+v excel-word
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCopy,$remoteExcelTitle,$excelCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wwrjv" Then ;00107 右键剪切-远程ctrl+v word-word
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$remoteWordTitle,$wordCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "werjv" Then;00108 右键剪切-远程ctrl+v word-excel
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$remoteWordTitle,$wordCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtrjv" Then;00109 右键剪切-远程ctrl+v word-txt
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyMenu,$wordCut,$remoteWordTitle,$wordCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttrjv" Then;00110 右键剪切-远程ctrl+v txt-txt
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$remoteTxtTitle,$txtCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twrjv" Then;00111 右键剪切远程ctrl+v txt-word
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$remoteTxtTitle,$txtCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "terjv" Then;00112 右键剪切-远程ctrl+v txt-excel
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyMenu,$txtCut,$remoteTxtTitle,$txtCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eerjv" Then;00113  右键剪切-远程ctrl+v excel-excel
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$remoteExcelTitle,$excelCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etrjv" Then;00114  右键剪切-远程ctrl+v excel-txt
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$remoteExcelTitle,$excelCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewrjv" Then;00115  右键剪切-远程ctrl+v excel-word
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyMenu,$excelCut,$remoteExcelTitle,$excelCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wwcrz" Then ;00116 右ctrl+c-远程右键粘贴word-word
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteWordTitle,$wordCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wecrz" Then;ctrl+c-远程右键粘贴word-excel
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteWordTitle,$wordCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtcrz" Then;00118 ctrl+c-远程右键粘贴word-txt
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteWordTitle,$wordCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttcrz" Then;00119 ctrl+c-远程右键粘贴txt-txt
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteTxtTitle,$txtCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twcrz" Then;00120 ctrl+c-远程右键粘贴txt-word
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteTxtTitle,$txtCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "tecrz" Then;00121 ctrl+c-远程右键粘贴txt-excel
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteTxtTitle,$txtCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eecrz" Then;00122  ctrl+c-远程右键粘贴excel-excel
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteExcelTitle,$excelCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etcrz" Then;00123  ctrl+c-远程右键粘贴excel-txt
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteExcelTitle,$excelCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewcrz" Then;00124  ctrl+c-远程右键粘贴excel-word
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteExcelTitle,$excelCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wwxrz" Then ;00125 本地ctrl+x-远程右键粘贴word-word
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteWordTitle,$wordCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wexrz" Then;00126 本地ctrl+x-远程右键粘贴word-excel
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteWordTitle,$wordCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtxrz" Then;00127 本地ctrl+x-远程右键粘贴word-txt
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteWordTitle,$wordCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttxrz" Then;00128 本地ctrl+x-远程右键粘贴txt-txt
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteTxtTitle,$txtCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twxrz" Then;00129 本地ctrl+x-远程右键粘贴txt-word
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteTxtTitle,$txtCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "texrz" Then;00130 本地ctrl+x-远程右键粘贴txt-excel
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteTxtTitle,$txtCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eexrz" Then;00131  本地ctrl+x-远程右键粘贴excel-excel
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteExcelTitle,$excelCloseTitle,$excel,$excelTitle,$opMethod_pasteMenu,$excelPaste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etxrz" Then;00132  本地ctrl+x-远程右键粘贴excel-txt
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteExcelTitle,$excelCloseTitle,$txt,$txtTitle,$opMethod_pasteMenu,$txtPate,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewxrz" Then;00133 本地ctrl+x-远程右键粘贴excel-word
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteExcelTitle,$excelCloseTitle,$word,$wordTitle,$opMethod_pasteMenu,$wordPaste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wwcv" Then ;00134  ctrl+c-远程ctrl+v word-word
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteWordTitle,$wordCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wecv" Then;00135 ctrl+c-远程ctrl+v word-excel
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteWordTitle,$wordCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtcv" Then;00136 ctrl+c-远程ctrl+v word-txt
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteWordTitle,$wordCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttcv" Then;00137 ctrl+c-远程ctrl+v txt-txt
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteTxtTitle,$txtCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twcv" Then;00138  ctrl+c-远程ctrl+v txt-word
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteTxtTitle,$txtCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "tecv" Then;00139  ctrl+c-远程ctrl+v txt-excel
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteTxtTitle,$txtCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eecv" Then;00140  ctrl+c-远程ctrl+v excel-excel
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteExcelTitle,$excelCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etcv" Then;00141  ctrl+c-远程ctrl+v excel-txt
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteExcelTitle,$excelCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewcv" Then;00142   ctrl+c-远程ctrl+v excel-word
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$copy,$remoteExcelTitle,$excelCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wwxv" Then ;00143 本地ctrl+x-远程ctrl+v word-word
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteWordTitle,$wordCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "wexv" Then;00144 本地ctrl+x-远程ctrl+v word-excel
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteWordTitle,$wordCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "wtxv" Then;00145 本地ctrl+x-远程ctrl+v word-txt
		Local $opKind =  [$remoteWord,$remoteWordTitle,$wordTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteWordTitle,$wordCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ttxv" Then;00146 本地ctrl+x-远程ctrl+v txt-txt
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteTxtTitle,$txtCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "twxv" Then;00147 本地ctrl+x-远程ctrl+v txt-word
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteTxtTitle,$txtCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	ElseIf $arrayKey[2] == "texv" Then;00148 本地ctrl+x-远程ctrl+v txt-excel
		Local $opKind =  [$remoteTxt,$remoteTxtTitle,$txtTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteTxtTitle,$txtCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "eexv" Then;00149  本地ctrl+x-远程ctrl+v excel-excel
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteExcelTitle,$excelCloseTitle,$excel,$excelTitle,$opMethod_pasteHotKey,$paste,$excelTitle,$excelCloseTitle]
	ElseIf $arrayKey[2] == "etxv" Then;00150  本地ctrl+x-远程ctrl+v excel-txt
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteExcelTitle,$excelCloseTitle,$txt,$txtTitle,$opMethod_pasteHotKey,$paste,$txtTitle,$txtCloseTitle]
	ElseIf $arrayKey[2] == "ewxv" Then;00151 本地ctrl+x-远程ctrl+v excel-word
		Local $opKind =  [$remoteExcel,$remoteExcelTitle,$excelTabCount,$arrayKey[1],$opMethod_copyHotKey,$cut,$remoteExcelTitle,$excelCloseTitle,$word,$wordTitle,$opMethod_pasteHotKey,$paste,$wordTitle,$wordCloseTitle]
	EndIf

	execute_cmd($opKind)

Next