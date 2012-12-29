Set objwShell = CreateObject("Wscript.Shell")
ptdesk = objwShell.SpecialFolders("Desktop")
Set objShortCut = objwShell.CreateShortcut(ptdesk & "\PRIMUS.lnk")
objShortCut.TargetPath = "c:\windows\regedit.exe"
objShortCut.Description = "PRIMUS"
objShortCut.WindowStyle = iWinStyle
objShortCut.HotKey = sHotKey
objShortCut.IconLocation = "c:\windows\explorer.exe"
objShortCut.WorkingDirectory = "c:\primus"
objShortCut.Save


    'Check to see if there is a backslash at
sub CreateDesktopINI (strPath)
    '     the end of the path
	'oShell.Run "copy " & strPath & " %USERPROFILE%\Desktop\", 0, True
End sub
