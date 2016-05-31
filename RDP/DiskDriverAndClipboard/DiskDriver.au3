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
#include "C:\Documents and Settings\Administrator\桌面\RDP\TitleAudit\DiskDriverAndClipboard.au3"

;模拟在桌面新建文件夹操作
create_note(@DesktopDir & "\ss.txt")

;模拟新建文件夹操作
create_folder("C:\Documents and Settings\Administrator\桌面\a")

;模拟操作文件复制和剪切的过程
move_file("C:\Documents and Settings\Administrator\桌面\a","C:\Documents and Settings\Administrator\桌面\2\a","C:\Documents and Settings\Administrator\桌面\3\a")



