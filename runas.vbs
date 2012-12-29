set WshShell = CreateObject("WScript.Shell")
WshShell.Run "runas /user:administrador ""c:\windows\system32\notepad.exe c:\boot.ini"""
WScript.Sleep 100
WshShell.Sendkeys "admadm~"