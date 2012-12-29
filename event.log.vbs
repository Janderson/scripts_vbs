' EventLogFSO.vbs
' Sample VBScript to Create a file ready for WMI
' Author Guy Thomas http://computerperformance.co.uk/
' Version 1.5 - November 2005
' -----------------------------------------------------------'
Option Explicit

Dim objFSO, objFolder, objFile ' Objects
Dim strComputer, strFileName, strFolder, strPath ' strings

' --------------------------------------------------------
' Set the folder and file name
strComputer = "."
strFileName = "\Event672.txt"
strFolder = "c:\"
strPath = strFolder & strFileName

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
Set strFileName = objFSO.CreateTextFile(strPath, True)
strFileName.WriteLine("Computer to test " & strComputer)
Wscript.Echo "Check " & strPath

WScript.Quit

' End of Guy's FSO sample VBScript