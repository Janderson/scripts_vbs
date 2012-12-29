'Create File Object
Set objfileobject = CreateObject("Scripting.FilesystemObject")
Set objFile = objfileobject.CreateTextFile("C:\resultante.txt")
Set objfiletoread = objfileobject.getfile("C:\filetoget.txt")
Set objtextstream = objfiletoread.openastextstream
do while not objtextstream.atendofstream
rdtxt=objtextstream.readline
	if InStr(rdtxt, "aqui") then ' pode ser strcomp para ficar mais preciso
		msgbox "!!!! encontrado  !!!!!"
		objFile.writeline("Nome: janderson")
		objFile.writeline("Setor Tecnologia da Informação")
	else 
		msgbox rdtxt
		objFile.writeline(rdtxt)
	end if
loop

objfiletoread.close
objFile.close