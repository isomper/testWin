#include "WindowTitle.au3"
#include "RightHandButton.au3"
#include  "ServiceWindow.au3"

;点击开始菜单并遍历二级菜单
start_cell("^{esc}",0,9,0,22,3)
start_cell("^{esc}",0,10,0,32,3)
;遍历一级菜单及其右键内容
Local $listdown[20]=[1,1,1,1,1,1,2,2,2,2,9,10,13,1,1,1,1,12,12,1]
click_start_all("^{esc}",0,20,$listdown)

;//开始菜单需要的参数，不可删除，无需执行
;click_start("^{esc}",0,8,2)
;click_start_all("^{esc}",0,10,2)
;click_start("^{esc}",0,11,9)
;click_start("^{esc}",0,14,12)