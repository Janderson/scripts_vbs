' LogicalDisk.vbs
' Sample VBScript to interrogate a Logical disk with WMI
' Author Guy Thomas http://computerperformance.co.uk/
' Version 1.8 - November 2005
' -------------------------------------------------------------'
Option Explicit
Dim objWMIService, objItem, colItems, strComputer

On Error Resume Next
strComputer = "."

Set objWMIService = GetObject _
("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_LogicalDisk")

For Each objItem in colItems
Wscript.Echo "Computer: " & objItem.SystemName & VbCr & _
" ==================================" & VbCr & _
"Drive Letter: " & objItem.Name & vbCr & _
"Description: " & objItem.Description & vbCr & _
"Volume Name: " & objItem.VolumeName & vbCr & _
"Drive Type: " & objItem.DriveType & vbCr & _
"Media Type: " & objItem.MediaType & vbCr & _
"VolumeSerialNumber: " & objItem.VolumeSerialNumber & vbCr & _
"Size: " & Int(objItem.Size /1073741824) & " GB" & vbCr & _
"Free Space: " & Int(objItem.FreeSpace /1073741824) & _
" GB" & vbCr & _
"Quotas Disabled: " & objItem.QuotasDisabled & vbCr & _
"Supports DiskQuotas: " & objItem.SupportsDiskQuotas & vbCr & _
"Supports FileBasedCompression: " & _
objItem.SupportsFileBasedCompression & vbCr & _
"Compressed: " & objItem.Compressed & vbCr & _
""
Next

WSCript.Quit

' End of Sample DiskDrive VBScript

