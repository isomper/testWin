#cs-----------------------------------------------------------------------------------------------------------------------
获取操作系统版本信息
#ce-----------------------------------------------------------------------------------------------------------------------
Func GetOSVersion()
	$objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	$colItems = $objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	For $os In $colItems
		return $os.Caption&" "&$os.Version
	Next
EndFunc
Local $os_Version =  StringSplit(GetOSVersion()," ")[4]