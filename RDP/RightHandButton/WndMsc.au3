#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/28
作者：张泽方
功能说明：RightHandButton.au3,遍历windows系统右键菜单
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include "C:\Documents and Settings\Administrator\桌面\RDP\TitleAudit\RightHandButton.au3"

;计算机管理
;打开计算机模块，遍历右键菜单（计算机管理-系统工具-事件查看器-应用程序右键菜单和二级菜单）
Local $subItem[1][2]
$subItem[0][0] = "5"
$subItem[0][1] = 6
msc_appskey("C:\WINDOWS\system32\compmgmt.msc","计算机管理",2,1,1,11)


;遍历计数器日志右键和二级菜单
Local $subItem[1][2]
$subItem[0][0] = "3"
$subItem[0][1] = 5
msc_appskey("C:\WINDOWS\system32\compmgmt.msc","计算机管理",5,1,1,7)


;遍历计算机管理（本地）右键菜单和二级菜单
Local $subItem[1][2]
$subItem[0][0] = "2"
$subItem[0][1] = 1
msc_appskey("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,0,0,6)


;遍历计算机管理菜单（存储功能以上）
Local $listdown[7] = [2,8,3,6,2,4,3]
msc_system_mgt("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$listdown)
;遍历计算机管理菜单（存储功能以下）
Local $listdown[4]=[1,11,3,4]
service("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$listdown)

