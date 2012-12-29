const HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set StdOut = WScript.StdOut
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")
 
 
strKeyPath = "SOFTWARE\Microsoft\Internet Explorer\New Windows\Allow"
strValueName = "192.5.0.2"
dwValue = 0
oReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue