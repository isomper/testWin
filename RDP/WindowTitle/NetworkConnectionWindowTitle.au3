#include "WindowTitle.au3"
#include "RightHandButton.au3"
#include  "ServiceWindow.au3"


;网络连接窗口标题
Local $alt_key[7]=["F","E","V","A","T","N","H"]
Local $listdown[7]=[9,8,11,4,3,5,2]
window_title("C:\WINDOWS\system32\ncpa.cpl","网络连接",$alt_key,$listdown)