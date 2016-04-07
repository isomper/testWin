#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/07
作者：顾亚茹
功能说明：包含WndTitle.au3（windows IE浏览器、写字板、计算机管理、网络连接的标题菜单）、mouseright.au3（遍历windows系统右键菜单，开始菜单）
、services.au3（查看和关闭windows服务的属性信息）
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------


#include "WndTitle.au3"
#include "mouseright.au3"
#include  "services.au3"
Local $alt_key[5]=["F","A","V","W","H"]
Local $listdown[5]=[1,5,5,4,3]
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
WdnTitle("C:\WINDOWS\system32\compmgmt.msc","计算机管理",$alt_key,$listdown)