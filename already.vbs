' Already.vbs Windows Logon Script
' VBScript to map a network drive.
' Author Guy Thomas http://computerperformance.co.uk/
' Version 1.7 - April 24th 2005
' ------------------------------------------------------'
Option Explicit
Dim strDriveLetter, strRemotePath
Dim objNetwork, objShell
Dim CheckDrive, AlreadyConnected, intDrive
' The section sets the variables.
strDriveLetter = "W:"
strRemotePath = "\\alan\drivers"

' This sections creates two objects:
' objShell and objNetwork and counts the drives
Set objShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network")
Set CheckDrive = objNetwork.EnumNetworkDrives()

' This section deals with a For ... Next loop
' See how it compares the enumerated drive letters
' with strDriveLetter
On Error Resume Next
AlreadyConnected = False
For intDrive = 0 To CheckDrive.Count - 1 Step 2
If CheckDrive.Item(intDrive) =strDriveLetter _
Then AlreadyConnected =True
Next

' This section uses the If = then, else logic
' This tests to see if the Drive is already mapped.
' If yes then disconnects
If AlreadyConnected = True then
objNetwork.RemoveNetworkDrive strDriveLetter
objNetwork.MapNetworkDrive strDriveLetter, strRemotePath

' The first message box
objShell.PopUp "Drive " & strDriveLetter & _
"Disconnected, then connected successfully."
Else
objNetwork.MapNetworkDrive strDriveLetter, strRemotePath
objShell.PopUp "Drive " & strDriveLetter & _
" connected successfully." End if
WScript.Quit

' Guy's Script ends here 