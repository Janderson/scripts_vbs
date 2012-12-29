function fnShellExecuteVB()
   dim objShell
   set objShell = CreateObject("Shell.Application")
   objShell.ShellExecute "nome_do_aplicativo.exe", "", "", "open", 1
   set objShell = nothing
end function

fnShellExecuteVB