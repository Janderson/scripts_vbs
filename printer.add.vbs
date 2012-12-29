Set objWMIService = GetObject("winmgmts:")
Set objDriver = objWMIService.Get("Win32_PrinterDriver")
objDriver.Name = "Apple LaserWriter 8500"
objDriver.SupportedPlatform = "Windows NT x86"
objDriver.Version = "3"
errResult = objDriver.AddPrinterDriver(objDriver)