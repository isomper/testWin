#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/28
作者：张泽方
功能说明：包含WindowTitle.au3（windows IE浏览器、写字板、计算机管理、网络连接的标题菜单）
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------


#include "WindowTitle.au3"
#include "RightHandButton.au3"
#include  "ServiceWindow.au3"

;Windows窗口标题
Local $alt_key[5]=["F","A","V","W","H"]
Local $listdown[5]=[1,5,5,4,3]
window_title("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$alt_key,$listdown)

;ie窗口标题
Local $alt_key[6]=["F","E","V","A","T","H"]
Local $listdown[6]=[12,4,10,4,5,4]
window_title("C:\Program Files\Internet Explorer\iexplore.exe","没有启用 Internet Explorer 增强的安全配置 - Microsoft Internet Explorer",$alt_key,$listdown)

;记事本窗口标题
Local $alt_key[6]=["F","E","V","I","O","H"]
Local $listdown[6]=[12,12,4,1,3,1]
window_title("C:\Program Files\Windows NT\Accessories\wordpad.exe","文档 - 写字板",$alt_key,$listdown)
