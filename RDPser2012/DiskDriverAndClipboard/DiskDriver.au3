#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/04/26
作者：顾亚茹
功能说明：模拟文件和文件夹的新建、复制、剪切、删除及word、excel、txt文件之间相互粘贴、复制、剪切操作
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include<Word.au3>
#include<Excel.au3>
#include<file.au3>
#include "C:\Users\Administrator\Desktop\RDPser2012\TitleAudit\DiskDriver.au3"


;新建文件夹
create_folder("D:\disk_driver")

;新建文本文档
create_file("D:\disk_driver\xinjian.txt","C:\Windows\System32\notepad.exe","无标题 - 记事本","Button1")


;模拟创建word文件操作
;create_word("D:\disk_driver\xinjian.docx")
create_file("D:\disk_driver\xinjian.docx","C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.EXE","文档 1 - Microsoft Word","Button8")


;模拟创建excel文件操作
create_file("D:\disk_driver\xinjian.xlsx","C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE","Microsoft Excel","Button6")

;模拟操作文件复制和剪切的过程
move_file("D:\disk_driver","D:\aa\disk_driver","D:\bb\aa\disk_driver")






