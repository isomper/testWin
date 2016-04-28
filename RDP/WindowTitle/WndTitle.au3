#include "WindowTitle.au3"
#include "RightHandButton.au3"
#include  "ServiceWindow.au3"
;Windows窗口标题
Local $alt_key[5]=["F","A","V","W","H"]
Local $listdown[5]=[1,5,5,4,3]
window_title("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$alt_key,$listdown)
