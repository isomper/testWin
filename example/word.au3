#cs------------------------------------------------------------
功能说明：打开word进行相关操作
#ce------------------------------------------------------------
#include <Word.au3>

#cs----------------------------start--------------------------------
打开已存在的word

;打开桌面的test.docx
ShellExecute("test.docx","","C:\Users\yy\Desktop")
WinWaitActive("test.docx")
Send("this is a test")		;在名字为test的word中输入内容
#ce----------------------------end---------------------------------

#cs----------------------------start--------------------------------
在桌面创建个新word
#ce----------------------------end---------------------------------

;在桌面创建一个名字为123.docx的word文件
Local $iWord = _Word_Create()		;创建word进程

_Word_DocAdd($iWord)					;追加一个空word页

;Local $iDoc = _Word_DocOpen($iWord, "C:\Users\yy\Desktop\123.docx")

;_Word_DocSaveAs($iDoc, "C:\Users\yy\Desktop\123.docx")

Sleep(1000)										;睡1秒

_Word_Quit($iWord)						;关闭word






