'**************************************************
'Taking Screenshot using word object
Set oWordBasic = CreateObject("Word.Basic")
oWordBasic.SendKeys "{prtsc}" 
oWordBasic.AppClose "Microsoft Word"
Set oWordBasic = Nothing
WScript.Sleep 2000
 
'Opening Paint Application
set WshShell = CreateObject("WScript.Shell")
WshShell.Run "mspaint"
WScript.Sleep 2000
 

'Activating Paint Application
WshShell.AppActivate "untitled - Paint"
WScript.Sleep 1000
 
'Paste the captured Screenshot
WshShell.SendKeys "^v"
WScript.Sleep 500
 
'Save Screenshot
WshShell.SendKeys "^s"
WScript.Sleep 500
WshShell.SendKeys "c:\test.bmp"
WScript.Sleep 500
WshShell.SendKeys "{ENTER}"
 
'Release Objects
Set WshShell=Nothing
WScript.Quit
'**************************************************