Dim strNome
Dim intIdade
Dim intResultado
Dim intOpcao

strNome = "Vinicius"
intIdade = 19
intResultado = msgbox("Ol�, meu nome � " & strNome & " e tenho " & intIdade & " anos.")

intOpcao = msgbox("Voc� est� entendendo o artigo sobre WSH?",vbYesNo)
if intOpcao = vbNo then 
   msgbox "Ok, vamos melhorar a explica��o ent�o..."
end if
