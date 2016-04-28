#include "WindowTitle.au3"
#include "RightHandButton.au3"
#include  "ServiceWindow.au3"
;计算机管理
;打开计算机模块，遍历右键菜单（计算机管理-系统工具-事件查看器-应用程序右键菜单和二级菜单）
msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,2,1,1,5,6,2)
;打开计算机模块，遍历右键菜单（计算机管理-系统工具-事件查看器-应用程序右键菜单）
Msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,2,1,1,11,0,1)
;计算管理模块菜单
msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,0,1,11,0,0,1)
;打开计算机模块，遍历右键菜单（计算管理-系统工具-事件查看器右键菜单及二级菜单）
msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,0,0,2,2,2,2)
;打开计算机模块，遍历右键菜单（计算管理-系统工具右键菜单及二级菜单）
msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,0,0,1,1,5,2)