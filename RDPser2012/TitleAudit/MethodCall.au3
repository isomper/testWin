#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/07
作者：顾亚茹
功能说明：包含WindowTitle.au3（windows IE浏览器、写字板、计算机管理、网络连接的标题菜单）、RightHandButton.au3（遍历windows系统右键菜单，开始菜单）
、ServiceWindow.au3（查看和关闭windows服务的属性信息）
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------


#include "WindowTitle.au3"
#include "RightHandButton.au3"
#include  "ServiceWindow.au3"
;Local $alt_key[5]=["F","A","V","W","H"]
;Local $listdown[5]=[1,5,5,4,3]
;ie C:\Program Files\Internet Explorer\iexplore.exe
;SandAright("^{esc}",10,2,6,1)
;SandAright("^{esc}",10,2,6,1)
;clickstart("^{esc}",9,2)
;点击开始
;MandAright("right",320,468,1,1,7)
;secpolAright("C:\WINDOWS\system32\secpol.msc","本地安全设置",3,4)
;MscAright("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,2,1,1,5,2)
;MscAright("C:\WINDOWS\system32\rrasmgmt.msc","路由和远程访问",0,0,0,0,4,1)
;clickstart("#e",2,0,5)
;调用clickstart和Aright
;WdnTitle("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$alt_key,$listdown)

;计算机管理
;msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,2,1,1,5,6,2)
;Msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,2,1,1,11,0,1)
;msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,0,1,11,0,0,1)
;msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,0,0,2,2,2,2)
;msc_cell("C:\WINDOWS\system32\compmgmt.msc","计算机管理",0,0,0,1,1,5,2)

;模拟点击鼠标右键
;mouse_right("right",601,359,2,7)
;right_cell("right",601,359,2,1,8,2)
;right_cell("right",601,359,2,5,16,2)

;遍历一级菜单及其右键内容
;Local $listdown[20]=[1,1,1,1,1,1,2,2,2,2,9,10,13,1,1,1,1,12,12,1]
;click_start_all("^{esc}",0,20,$listdown)
;点击开始菜单并遍历二级菜单
;start_cell("^{esc}",0,9,0,22,3)
;start_cell("^{esc}",0,10,0,32,3)
;//
;click_start("^{esc}",0,8,2)
;click_start_all("^{esc}",0,10,2)
;click_start("^{esc}",0,11,9)
;click_start("^{esc}",0,14,12)

;Windows窗口标题
;Local $alt_key[5]=["F","A","V","W","H"]
;Local $listdown[5]=[1,5,5,4,3]
;window_title("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$alt_key,$listdown)

;记事本窗口标题
;Local $alt_key[6]=["F","E,"V","I","O""H"]
;Local $listdown[6]=[13,13,5,2,4,2]
;window_title("C:\Program Files\Windows NT\Accessories\wordpad.exe","文档 - 写字板",$alt_key,$listdown)

;service服务
;service_attributes("C:\WINDOWS\system32\services.msc","服务",119)







