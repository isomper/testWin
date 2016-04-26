#Include <Array.au3>
Dim $dirlist,$dirarry
$dirlist=_FileListEx("c:\root")
$dirarry=StringSplit($dirlist,"|")
_ArrayDisplay($dirarry)

Func _FileListEx($sDir)
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