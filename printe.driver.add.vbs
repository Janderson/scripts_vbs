strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objDriver = objWMIService.Get("Win32_PrinterDriver")
objDriver.Name = "NewPrinter Model 2900"
objDriver.SupportedPlatform = "Windows NT x86"
objDriver.Version = "3"
objDriverPath = "C:\Scripts\NewPrinter.dll"
objInfname = "C:\Scripts\NewPrinter.inf"
intResult = objDriver.AddPrinterDriver(objDriver)
Wscript.Echo intResult