MouseClick("left",77,167) ;选择文件存在路径位置
Sleep(1000)
Send("{ENTER}")
MouseClick("left",547,533) ;定位到上传输入框
Send("{0.png}") ;输入要上传的文件
Send("{down 1}")   ;选择第一个
Sleep(1500)
Send("{ENTER}")
Sleep(1000)

