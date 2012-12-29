With CreateObject("Wscript.Shell").Exec("ipconfig /all").StdOut  
  Do Until.AtEndOfStream    
    msgbox(ReadLine,vbsyesno)
  Loop
End With
