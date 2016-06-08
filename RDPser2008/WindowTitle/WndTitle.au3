

#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/05/30
作者：顾亚茹\Wind
功能说明：遍历windows 计算机管理标题菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------


#include "E:\RDPser2008\TitleAudit\WindowTitle.au3"

;计算机管理窗口标题
Local $alt_key[4]=["F","A","V","H"]
Local $listdown[4]=[1,3,5,3]
Local $arr[1][3]
$arr[0][0] = "A"
$arr[0][1] = 1
$arr[0][2] = 1
window_title("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$alt_key,$listdown)

;记事本窗口标题
Local $alt_key[5]=["F","E","O","V","H"]
Local $listdown[5]=[6,10,1,0,1]
Local $arr[0][0]
window_title("C:\Windows\System32\notepad.exe","无标题 - 记事本",$alt_key,$listdown)

