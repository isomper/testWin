#include <Constants.au3>
#cs----------------------------------------------------------------------------------------------------------------------
�������ڣ�2016/05/25
���ߣ�������
����˵���������ļ���word��excel��txt����Զ���ļ���word��excel��txt���ļ�֮���໥���ơ�ճ�������в���
�޸����ڣ�
�޸��ˣ�
�޸����ݣ�
#ce-----------------------------------------------------------------------------------------------------------------------

#include <Constants.au3>
#include "C:\Users\zzf\Desktop\RDP\TitleAudit\Clipboard.au3"

;word����ճ�����ı�txt,(�ӱ��ص�Զ�̷�����)
clip_board_up(@DesktopDir & "\a.docx","a.docx - Microsoft Word",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Ptxt.sikuli",1)
Sleep(10000)

;word����ճ����word,(�ӱ��ص�Զ�̷�����)
clip_board_up(@DesktopDir & "\a.docx","a.docx - Microsoft Word",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Pword.sikuli",1)
Sleep(10000)

;word����ճ����excel,(�ӱ��ص�Զ�̷�����)
clip_board_up(@DesktopDir & "\a.docx","a.docx - Microsoft Word",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\paste111.sikuli",1)
Sleep(10000)

;excel����ճ����excel,(�ӱ��ص�Զ�̷�����)
clip_board_up(@DesktopDir & "\abc.xlsx","Microsoft Excel - abc.xlsx",1,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\paste111.sikuli",0)
Sleep(10000)

;txt����ճ����txt,(�ӱ��ص�Զ�̷�����)
clip_board_up(@DesktopDir & "\wenben.txt","wenben.txt - ���±�",3,"C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Ptxt.sikuli",1)
Sleep(10000)


;txt����ճ����txt,(��Զ�̷�����������)
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Ctxt.sikuli",@DesktopDir & "\wenben.txt","wenben.txt - ���±�",4,1)

;word����ճ����word,(��Զ�̷�����������)
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Cword.sikuli",@DesktopDir & "\a.docx","a.docx - Microsoft Word",2,1)

;word����ճ����txt,(��Զ�̷�����������)
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\Cword.sikuli",@DesktopDir & "\wenben.txt","wenben.txt - ���±�",4,1)

;word����ͼƬ��word����Զ�̵����أ�
clip_board_down("C:\Users\zzf\Desktop\runsikulix -r C:\Users\zzf\Desktop\remotetupian.sikuli",@DesktopDir & "\tupian11.docx","tupian11.docx - Microsoft Word",2,2)
#ce