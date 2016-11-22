#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/05/26
作者:张泽方
功能说明：打开并且关闭指定目录下的所有文件夹（2000形近字）
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------
#Include <Array.au3>
#Include "C:\Documents and Settings\Administrator\桌面\RDP\TitleAudit\OperationDir.au3"

;遍历2000形近字文件夹
Dim $dirList,$dirArry
;获取所有文件夹的绝对路径
$dirList=file_list_ex("C:\Documents and Settings\Administrator\桌面\2000FromNearWord")
;拆分绝对路径放在数组里
$dirArry=StringSplit($dirList,"|")

for $i=1 To UBound($dirArry);$dirArry[0]
	;打开文件夹
	ShellExecute($dirArry[$i])
	Sleep(1000)
	;关闭文件夹
	Send("!{f4}")
Next