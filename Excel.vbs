Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("z:\New_ports.xls")
objExcel.Visible = True
intRow = 2
Do Until objExcel.Cells(intRow,1).Value = ""
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:")
	Set objNewPort = objWMIService.Get("Win32_TCPIPPrinterPort").SpawnInstance_
	objNewPort.Name = objExcel.Cells(intRow,5).Value
	objNewPort.Protocol = 1
	objNewPort.HostAddress = objExcel.Cells(intRow,1).Value
	objNewPort.PortNumber = "9100"
	objNewPort.SNMPEnabled = False
	objNewPort.Put_
	intRow = intRow + 1
Loop
objExcel.Quit

