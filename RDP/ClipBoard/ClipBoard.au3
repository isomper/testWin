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


;模拟下行文本与word文字复制粘贴操作
;clip_board("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\txt.sikuli","wenzi.txt - 记事本",3,@DesktopDir & "\a.docx","a.docx - Word",2,0)

;模拟下行文本与文本文字复制粘贴操作
;clip_board("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\txt.sikuli","wenzi.txt - 记事本",3,@DesktopDir & "\wenben.txt","wenben.txt - 记事本",4,0)

;模拟下行文本与excel文字复制粘贴操作
;clip_board("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\txt.sikuli","wenzi.txt - 记事本",3,@DesktopDir & "\abc.xlsx","abc.xlsx - excel",2,0)

;模拟下行word与excel图片复制粘贴操作
;clip_board("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\tupian.sikuli","wen.docx - Word",1,@DesktopDir & "\abc.xlsx","abc.xlsx - Excel",2,0)