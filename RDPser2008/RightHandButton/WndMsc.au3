#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/28
作者：张泽方
功能说明：RightHandButton.au3,遍历windows系统右键菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include "E:\RDPser2008\TitleAudit\RightHandButton.au3"

;计算机管理
#cs
;打开计算机模块，遍历右键菜单（计算机管理-系统工具-任务计划程序右键菜单和二级菜单）
Local $subItem[1][2]
$subItem[0][0] = "7"
$subItem[0][1] = 1
msc_appskey("C:\Windows\System32\compmgmt.msc","计算机管理",2,0,0,9)

;遍历事件查看器右键和二级菜单
Local $subItem[2][2]
$subItem[0][0] = "4"
$subItem[0][1] = 1
$subItem[1][0] = "6"
$subItem[1][1] = 1
msc_appskey("C:\WINDOWS\system32\compmgmt.msc","计算机管理",3,0,0,6)


;遍历共享文件夹右键菜单和二级菜单
Local $subItem[2][2]
$subItem[0][0] = "1"
$subItem[0][1] = 1
$subItem[1][0] = "2"
$subItem[1][1] = 5
msc_appskey("C:\WINDOWS\system32\compmgmt.msc","计算机管理",4,0,0,3)
#ce

;遍历计算机管理菜单
Local $listdown[5] = [4,4,2,3,5]
msc_system_mgt("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$listdown,2)


