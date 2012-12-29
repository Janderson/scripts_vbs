On Error Resume Next
'----------------------------------------------------------'
'        OBJETOS GLOBAIS                       '
'----------------------------------------------------------'

Set WshNetwork = WScript.CreateObject("wscript.network")
Set WshShell = WScript.CreateObject("wscript.Shell")

'----------------------------------------------------------'
'          MAPEANDO IMPRESSORAS               '
'----------------------------------------------------------'

WshNetwork.AddWindowsPrinterConnection "\\192.0.0.81\laserA3"
WshNetwork.AddWindowsPrinterConnection "\\192.0.0.81\laserA4"
wshNetWork.SetDefaultPrinter "\\192.0.0.81\laserA3"
