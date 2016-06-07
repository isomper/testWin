#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/03/31
作者：顾亚茹
功能说明：打开并且关闭指定目录下的所有文件夹
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------
#Include <Array.au3>

Dim $dirList,$dirArry
;获取所有文件夹的绝对路径
$dirList=file_list_ex("E:\2000FromNearWord")
;拆分绝对路径放在数组里
$dirArry=StringSplit($dirList,"|")

for $i=1 to $dirArry[0]
	;打开文件夹
	ShellExecute($dirArry[$i])
	Sleep(1000)
	;关闭文件夹
	Send("!{f4}")
Next


;获取所有文件夹的绝对路径
Func file_list_ex($sDir)
		If StringInStr(FileGetAttrib($sDir),"D")=0 Then Return SetError(1,0,"")
		Local $oFSO = ObjCreate("Scripting.FileSystemObject")
		Local $objDir
				Local $Sources=$sDir
		Local $aDir = StringSplit($sDir, "|", 2)
		Local $iCnt = 0
		;Local $sFiles = "",$Attributes,$DateCreated,$DateLastAccessed,$DateLastModified,$ShortName,$ShortPath,$Size,$Type
		Do
				$objDir = $oFSO.GetFolder($aDir[$iCnt])
				For $aItem In $objDir.SubFolders
						;扩展应用改下这句, 如指定文件夹 If StringInStr($aItem.Name, "XXX") Then
						$sDir &= "|" & $aItem.Path
						;文件夹层数可以通过 StringReplace($aItem.Path, "\", "", 0, 1)的@extended值来判断
				Next
				$iCnt += 1
				If UBound($aDir) <= $iCnt Then $aDir = StringSplit($sDir, "|", 2)
		Until UBound($aDir) <= $iCnt
		Return $sDir
EndFunc