#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/05/30
作者：顾亚茹
功能说明：遍历windows IE浏览器、写字板、计算机管理、网络连接、我的文档的标题菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------


#include "C:\Documents and Settings\Administrator\桌面\RDP\TitleAudit\WindowTitle.au3"

;计算机管理窗口标题
Local $alt_key_mgt[5]=["F","A","V","W","H"]
Local $listdown_mgt[5]=[1,5,5,4,3]
Local $arr[1][3]
$arr[0][0] = "A"
$arr[0][1] = 1
$arr[0][2] = 2
window_title("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$alt_key_mgt,$listdown_mgt)


;ie窗口标题
Local $alt_key_ie[6]=["F","E","V","A","T","H"]
Local $listdown_ie[6] = [12,4,9,4,5,5]
Local $arr[11][3]
$arr[0][0] = "F"
$arr[0][1] = "0"
$arr[0][2] = 5

$arr[1][0] = "F"
$arr[1][1] = "8"
$arr[1][2] = 3

$arr[2][0] = "V"
$arr[2][1] = "0"
$arr[2][2] = 5

$arr[3][0] = "V"
$arr[3][1] = "2"
$arr[3][2] = 8

$arr[4][0] = "V"
$arr[4][1] = "3"
$arr[4][2] = 4

$arr[5][0] = "V"
$arr[5][1] = "6"
$arr[5][2] = 5

$arr[6][0] = "V"
$arr[6][1] = "7"
$arr[6][2] = 5

$arr[7][0] = "V"
$arr[7][1] = "74"
$arr[7][2] = 32

$arr[8][0] = "A"
$arr[8][1] = "2"
$arr[8][2] = 4

$arr[9][0] = "T"
$arr[9][1] = "0"
$arr[9][2] = 5

$arr[10][0] = "T"
$arr[10][1] = "1"
$arr[10][2] = 2
window_title("C:\Program Files\Internet Explorer\iexplore.exe","没有启用 Internet Explorer 增强的安全配置 - Microsoft Internet Explorer",$alt_key_ie,$listdown_ie)

;记事本窗口标题
Local $alt_key[6]=["F","E","V","I","O","H"]
Local $listdown[6]=[12,12,4,1,3,1]
Local $arr[0][0]
window_title("C:\Program Files\Windows NT\Accessories\wordpad.exe","文档 - 写字板",$alt_key,$listdown)

;网络连接窗口标题
Local $alt_key_network[7]=["F","E","V","A","T","N","H"]
Local $listdown_network[7]=[9,8,11,4,3,5,2]
Local $arr[5][3]
$arr[0][0] = "V"
$arr[0][1] = 0
$arr[0][2] = 5

$arr[1][0] = "V"
$arr[1][1] = 2
$arr[1][2] = 6

$arr[2][0] = "V"
$arr[2][1] = 8
$arr[2][2] = 8

$arr[3][0] = "V"
$arr[3][1] = 10
$arr[3][2] = 5

$arr[4][0] = "A"
$arr[4][1] = 2
$arr[4][2] = 4
window_title("C:\WINDOWS\system32\ncpa.cpl","网络连接",$alt_key_network,$listdown_network)

;我的文档窗口标题
Local $alt_key[6]=["F","E","V","A","T","H"]
Local $listdown[6]=[7,6,11,4,3,2]
Local $arr[8][3]
$arr[0][0] = "F"
$arr[0][1] = "0"
$arr[0][2] = 2

$arr[1][0] = "F"
$arr[1][1] = "1"
$arr[1][2] = 19

$arr[2][0] = "F"
$arr[2][1] = "6"
$arr[2][2] = 8

$arr[3][0] = "V"
$arr[3][1] = 0
$arr[3][2] = 5

$arr[4][0] = "V"
$arr[4][1] = "2"
$arr[4][2] = 8

$arr[5][0] = "V"
$arr[5][1] = 8
$arr[5][2] = 8

$arr[6][0] = "V"
$arr[6][1] = "10"
$arr[6][2] = 5

$arr[7][0] = "A"
$arr[7][1] = "2"
$arr[7][2] = 4

window_title("C:\windows\explorer.exe","我的文档",$alt_key,$listdown)
