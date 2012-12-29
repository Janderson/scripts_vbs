Dim strNome
Dim intIdade
Dim intResultado
Dim intOpcao

strNome = "Vinicius"
intIdade = 19
intResultado = msgbox("Olá, meu nome é " & strNome & " e tenho " & intIdade & " anos.")

intOpcao = msgbox("Você está entendendo o artigo sobre WSH?",vbYesNo)
if intOpcao = vbNo then 
   msgbox "Ok, vamos melhorar a explicação então..."
end if
