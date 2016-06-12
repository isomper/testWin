#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/28
作者：张泽方
功能说明：包含RightHandButton.au3遍历windows系统桌面右键菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------



#include "C:\Documents and Settings\Administrator\桌面\RDP\TitleAudit\RightHandButton.au3"


;模拟点击鼠标右键
Local $subItem[3][2]
$subItem[0][0] = "1"
$subItem[0][1] = 8
$subItem[1][0] = "5"
$subItem[1][1] = 2
$subItem[2][0] = "6"
$subItem[2][1] = 19
mouse_right("right",350,743,1,7)
Send("{ESC}")