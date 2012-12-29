Dim objFileSystem, objInputFile
Dim strInputFile, inputData, strData

Const OPEN_FILE_FOR_READING = 1

' generate a filename base on the script name, here readfile.in
strOutputFile = "./" & Split(WScript.ScriptName, ".")(0) & ".in"

Set objFileSystem = CreateObject("Scripting.fileSystemObject")
Set objInputFile = objFileSystem.OpenTextFile(strOutputFile, _
  OPEN_FILE_FOR_READING)

' read everything in an array
inputData = Split(objInputFile.ReadAll, vbNewline)

For each strData In inputData
    WScript.Echo strData
Next

objInputFile.Close
Set objFileSystem = Nothing

WScript.Quit(0)