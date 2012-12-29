 

' Process.vbs
' Free Sample VBScript to discover which processes are running
' Author Guy Thomas http://computerperformance.co.uk/
' Version 1.4 - December 2005
' -------------------------------------------------------'
Option Explicit
Dim objWMIService, objProcess, colProcess
Dim strComputer, strList

strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

Set colProcess = objWMIService.ExecQuery _
("Select * from Win32_Process")

For Each objProcess in colProcess
strList = strList & vbCr & _
objProcess.Name
Next

WSCript.Echo strList
WScript.Quit

' End of List Process Example VBScript