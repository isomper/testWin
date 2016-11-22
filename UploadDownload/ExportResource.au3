#cs----------------------------------------------------
创建日期：2016/09/29
作者：张泽方
功能说明：通过堡垒系统在资源页面进行资源导出操作
修改日期：
修改人：
修改内容
#ce-----------------------------------------------------
;以只导出资源账号excel的方式导出
MouseClick("left",1022,421) ;选择是否导出资源账号
Sleep(1000)
Send("{down 1}")
MouseClick("left",1221,728)  ;点击确定
Sleep(1000)
MouseClick("left",1052,540) ;点击保存
Sleep(1000)
MouseClick("left",763,512)  ;选择文件存放路径-桌面
Sleep(1000)
MouseClick("left",1281,907)  ;点击确定


;以加密的文件方式导出

MouseClick("left",1572,328) ;定位导出按钮
Sleep(1000)
MouseClick("left",952,422) ;选择导出资源账号
Sleep(1000)
Send("{down}")
Sleep(1000)
Send("{ENTER}")
Sleep(2000)
Send("{tab}")
Sleep(1000)
Send("{down 1}")
MouseClick("left",1221,728)  ;点击确定
Sleep(1000)
MouseClick("left",1052,540) ;点击保存
Sleep(1000)
MouseClick("left",763,512)  ;选择文件存放路径-桌面
Sleep(1000)
MouseClick("left",1281,907)  ;点击确定

