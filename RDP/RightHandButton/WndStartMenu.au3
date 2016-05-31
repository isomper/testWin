#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/28
作者：张泽方
功能说明：包含RightHandButton.au3,遍历windows系统开始菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------


#include "WindowTitle.au3"
#include "RightHandButton.au3"
#include  "ServiceWindow.au3"



;遍历一级菜单及其右键内容，和遍历二级菜单

Local $listdown[20]=[1,1,1,1,1,1,2,2,2,2,9,10,13,1,1,1,1,12,12,1]
Local $subItem[8][2]
$subItem[0][0] = "8"
$subItem[0][1] = 1
$subItem[1][0] = "9"
$subItem[1][1] = 21
$subItem[2][0] = "10"
$subItem[2][1] = 31
$subItem[3][0] = "108"
$subItem[3][1] = 1
$subItem[4][0] = "1012"
$subItem[4][1] = 21
$subItem[5][0] = "1015"
$subItem[5][1] = 1
$subItem[6][0] = "1018"
$subItem[6][1] = 1
$subItem[7][0] = "1024"
$subItem[7][1] = 2
click_start_all("^{esc}",0,20,$listdown)
Send("{ESC}")

