<%Set FSO = Server.CreateObject("Scripting.FileSystemObject")
caminho = Server.MapPath("txt/teste.txt") 'especifique aqui o caminho onde ficará/está o TXT
Set GRAVAR = FSO.CreateTextFile(caminho,true)
'Foi criado o objeto e logo após busca o txt em caminho para gravar, se não achar, vai cria-lo (por causa da marcação TRUE)

gravar.write ("equipamento,controle,marca,modelo,periodicidade,data;equipamento,controle,marca,modelo,periodicidade,data;equipamento,controle,marca,modelo,periodicidade,data")
gravar.close
response.write "GRAVADO!"
'apos abrir o TXT, gravará a linha com o texto "TESTE DE GRAVAÇÃO" a confirmação no cliente aparecerá como "GRAVADO"
%>


equipamento,controle,marca,modelo,periodicidade,data,tipo_manut,;
equipamento,controle,marca,modelo,periodicidade,data;
equipamento,controle,marca,modelo,periodicidade,data
