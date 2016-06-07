#include <Constants.au3>
#cs----------------------------------------------------------------------------------------------------------------------
创建日期：2016/05/25
作者：顾亚茹
功能说明：本地文件（word、excel、txt）与远程文件（word、excel、txt）文件之间相互复制、粘贴、剪切操作
修改日期：
修改人：
修改内容：
#ce-----------------------------------------------------------------------------------------------------------------------

#include <Constants.au3>
#include "C:\Users\zzf\Desktop\RDP\TitleAudit\Clipboard.au3"

;word复制粘贴到文本txt,(从本地到远程服务器)
clip_board_up(@DesktopDir & "\a.docx","a.docx - Microsoft Word",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Ptxt.sikuli",1)
Sleep(10000)

;word复制粘贴到word,(从本地到远程服务器)
clip_board_up(@DesktopDir & "\a.docx","a.docx - Microsoft Word",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Pword.sikuli",1)
Sleep(10000)

;word复制粘贴到excel,(从本地到远程服务器)
clip_board_up(@DesktopDir & "\a.docx","a.docx - Microsoft Word",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\paste111.sikuli",1)
Sleep(10000)

;excel复制粘贴到excel,(从本地到远程服务器)
clip_board_up(@DesktopDir & "\abc.xlsx","Microsoft Excel - abc.xlsx",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\paste111.sikuli",0)
Sleep(10000)

;txt复制粘贴到txt,(从本地到远程服务器)
clip_board_up(@DesktopDir & "\wenben.txt","wenben.txt - 记事本",3,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Ptxt.sikuli",1)
Sleep(10000)


;txt复制粘贴到txt,(从远程服务器到本地)
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Ctxt.sikuli",@DesktopDir & "\wenben.txt","wenben.txt - 记事本",4,1)

;word复制粘贴到word,(从远程服务器到本地)
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Cword.sikuli",@DesktopDir & "\a.docx","a.docx - Microsoft Word",2,1)

;word复制粘贴到txt,(从远程服务器到本地)
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Cword.sikuli",@DesktopDir & "\wenben.txt","wenben.txt - 记事本",4,1)

;word复制图片到word（从远程到本地）
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\remotetupian.sikuli",@DesktopDir & "\tupian11.docx","tupian11.docx - Microsoft Word",2,2)
#ce