' OperatingSystem.vbs
' VBScript to document your Operating System
' Author Guy http://computerperformance.co.uk/
' Version 1.3 - June 2005
' -------------------------------------------------------'
Option Explicit
Dim objWMIService, objItem, colItems
Dim strComputer, strList

On Error Resume Next
strComputer = "."

' WMI Connection to the object in the CIM namespace
Set objWMIService = GetObject("winmgmts:\\" _
& strComputer & "\root\cimv2")

' WMI Query to the Win32_OperatingSystem
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")

' For Each... In Loop (Next at the very end)
For Each objItem in colItems
WScript.Echo "Computer: " & objItem.CSName
WScript.Echo "Processor: " & objItem.Description
WScript.Echo "Manufacturer: " & objItem.Manufacturer
WScript.Echo "Operating System: " & objItem.Caption
WScript.Echo "Version: " & objItem.Version
WScript.Echo "Service Pack: " & objItem.CSDVersion
WSCript.Echo ""
Next
WSCript.Quit

' End of WMI Win32_OperatingSystem VBScript