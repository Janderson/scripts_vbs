 
' EventIDSecurity.vbs
' Sample WMI to find and Event ID in the Security Log
' Author Guy Thomas http://computerperformance.co.uk/
' Version 1.7 - May 2006
' -----------------------------------------------------------'
Option Explicit

Dim objFSO, objFolder, objFile, objWMI, objItem ' Objects
Dim strComputer, strFileName, strFileOpen, strFolder, strPath
Dim intEvent, intNumberID, intRecordNum, colLoggedEvents

' --------------------------------------------------------
' Set the folder and file name
strComputer = "."
strFileName = "\Event672.txt"
strFolder = "c:"
strPath = strFolder & strFileName

' Set numbers
intNumberID = 672 ' Event ID Number
intRecordNum = 0

' -----------------------------------------------------
' Section to create folder and hold file.
' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check that the strFolder folder exists
If objFSO.FolderExists(strFolder) Then
Set objFolder = objFSO.GetFolder(strFolder)
Else
Set objFolder = objFSO.CreateFolder(strFolder)
WScript.Echo "Just created " & strFolder
End If

If objFSO.FileExists(strFolder & strFileName) Then
Set objFolder = objFSO.GetFolder(strFolder)
Else
Set objFile = objFSO.CreateTextFile(strFolder & strFileName)
Wscript.Echo "Just created " & strFolder & strFileName
End If
' --------------------------------------------------
' Two tiny but vital commands (Try script without)
set objFile = nothing
set objFolder = nothing

' ----------------------------------------------------
' Write the information to the file
Wscript.Echo " Press OK and Wait 30 seconds (ish)"
Set strFileOpen = objFSO.CreateTextFile(strPath, True)

' ----------------------------------------------------------
' WMI Core Section
Set objWMI = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate,(Security)}!\\" _
& strComputer & "\root\cimv2")
Set colLoggedEvents = objWMI.ExecQuery _
("Select * from Win32_NTLogEvent Where Logfile = 'Security'" )

' ----------------------------------------------------------
' Next section loops through ID properties

For Each objItem in colLoggedEvents
If objItem.EventCode = intNumberID Then
If objItem.EventType=5 then
strFileOpen.WriteLine("Category: " & objItem.Category _
& " string " & objItem.CategoryString)
strFileOpen.WriteLine("ComputerName: " & objItem.ComputerName)
strFileOpen.WriteLine("Logfile: " & objItem.Logfile _
& " source " & objItem.SourceName)
strFileOpen.WriteLine("EventCode: " & objItem.EventCode)
strFileOpen.WriteLine("EventType: " & objItem.EventType)
strFileOpen.WriteLine("Type: " & objItem.Type)
strFileOpen.WriteLine("User: " & objItem.User)
strFileOpen.WriteLine("Message: " & objItem.Message)
strFileOpen.WriteLine (" ")
intRecordNum = intRecordNum +1
End If
End If
Next
Wscript.Echo "Check " & strPath & " for " & intRecordNum & " events"

WScript.Quit

' End of Guy's FSO sample VBScript