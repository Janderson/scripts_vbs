dim a,b
str="CN=almir.snato,OU=SANTO,dc=com"
a=Split(str,",")
b = split(a(0),"=")
msgbox b(1)
