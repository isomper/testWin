#cs----------------------------------------------------
创建日期：2016/09/29
作者：张泽方
功能说明：通过堡垒页面点击浏览下载文件功能
修改日期：
修改人：
修改内容
#ce-----------------------------------------------------
MouseClick("left",1033,545)  ;定位保存按钮
Sleep(1000)
MouseClick("left",760,512)  ;定位保存路径（桌面）位置
Send("{ENTER}")
MouseClick("left",1279,917) ;点击保存