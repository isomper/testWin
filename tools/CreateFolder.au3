#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/27
作者：张利娟
功能说明：自动创建文件夹
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include <Constants.au3>
#include <ButtonConstants.au3>
#include <WindowsConstants.au3>


create_folder()
Func create_folder()
	Local $sdir
	$sdir=InputBox("提示", "请输入需要创建的文件夹路径", @ScriptDir)

		If DirCreate($sdir)=1 Then
			MsgBox(0,"","文件夹创建成功")
		Else
			MsgBox(0,"","文件夹创建失败，请检查输入内容及父路径是否只读")
		EndIf

EndFunc