<%Set FSO = Server.CreateObject("Scripting.FileSystemObject")
caminho = Server.MapPath("txt/teste.txt") 'especifique aqui o caminho onde ficar�/est� o TXT
Set GRAVAR = FSO.CreateTextFile(caminho,true)
'Foi criado o objeto e logo ap�s busca o txt em caminho para gravar, se n�o achar, vai cria-lo (por causa da marca��o TRUE)

gravar.write ("equipamento,controle,marca,modelo,periodicidade,data;equipamento,controle,marca,modelo,periodicidade,data;equipamento,controle,marca,modelo,periodicidade,data")
gravar.close
response.write "GRAVADO!"
'apos abrir o TXT, gravar� a linha com o texto "TESTE DE GRAVA��O" a confirma��o no cliente aparecer� como "GRAVADO"
%>


equipamento,controle,marca,modelo,periodicidade,data,tipo_manut,;
equipamento,controle,marca,modelo,periodicidade,data;
equipamento,controle,marca,modelo,periodicidade,data
