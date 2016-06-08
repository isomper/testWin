#include <Constants.au3>
#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/06/06
作者：张泽方
功能说明：本地文件（word、excel、txt）与远程文件（word、excel、txt）文件之间相互复制、粘贴、剪切操作
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include <Constants.au3>
#include "C:\Users\zzf\Desktop\RDPser2012\TitleAudit\Clipboard.au3"

;word复制粘贴到txt(从本地到远程服务器)
clip_board_up(@DesktopDir & "\wenjianjia\a.docx","a.docx - Microsoft Word",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\CP2012\ptxt.sikuli",1)
Sleep(10000)

;word复制粘贴到word(从本地到远程服务器)
clip_board_up(@DesktopDir & "\wenjianjia\a.docx","a.docx - Microsoft Word",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\CP2012\pword.sikuli",1)
Sleep(10000)

;word复制粘贴到excel(从本地到远程服务器)
clip_board_up(@DesktopDir & "\wenjianjia\a.docx","a.docx - Microsoft Word",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\CP2012\pxls.sikuli",1)
Sleep(10000)


;txt复制粘贴到txt(从本地到远程服务器)
clip_board_up(@DesktopDir & "\wenjianjia\wenben.txt","wenben.txt - 记事本",3,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\CP2012\ptxt.sikuli",1)
Sleep(10000)


;txt复制粘贴到txt(从远程服务器到本地)
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\CP2012\ctxt.sikuli",@DesktopDir & "\wenjianjia\wenben.txt","wenben.txt - 记事本",4,2012,1)


;word复制粘贴到txt(从远程服务器到本地)
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\cp2012\cword.sikuli",@DesktopDir & "\wenjianjia\wenben.txt","wenben.txt - 记事本",4,2012,1)


;word复制粘贴图片到word(从远程服务器到本地)
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\CP2012\tupian.sikuli",@DesktopDir & "\wenjianjia\tupian11.docx","tupian11.docx - Microsoft Word",2,2012,2)

