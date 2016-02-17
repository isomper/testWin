MouseClick("right",833,221)
sleep(500)
For $iNum = 8 To 1 Step -1
	ConsoleWrite($iNum & @LF)
	sleep(500)
	Send("{DOWN}")
	If $iNum = 8 Then
		Send("{RIGHT}")
		For $i = 7 To 1 Step -1
			sleep(500)
			Send("{DOWN}")
		Next
	EndIf
	Send("{LEFT}")
Next

MouseClick("left",1288,292)