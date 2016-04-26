#include <Constants.au3>
#include <ButtonConstants.au3>
#include <WindowsConstants.au3>

_Main()
Func _Main()
	Local $sDir
	$sDir=InputBox("提示", "请输入需要创建的文件夹路径", @ScriptDir)

		If DirCreate($sDir)=1 Then
			MsgBox(0,"","文件夹创建成功")
		Else
			MsgBox(0,"","文件夹创建失败，请检查输入内容及父路径是否只读")
		EndIf

EndFunc