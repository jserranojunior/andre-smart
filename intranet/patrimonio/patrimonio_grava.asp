<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<%
ok = request("ok")

origem = zero(1)
dt_gravacao = right(year(now),2)&zero(month(now))&zero(day(now))&zero(hour(now))&zero(minute(now))
usuario =  session("cd_codigo")
nome_arquivo = request("nome_arquivo")
nome_gravacao = origem&"_"&dt_gravacao&"_"&zero(usuario)&"-"&nome_arquivo
grava_txt = request("conteudo")

dt_hoje = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)%>



<%if ok = "" Then%>
<body>
<table border="1">
	<form action="patrimonio_grava.asp" method="post">
	<input type="hidden" name="conteudo" value="<%=grava_txt%>">
	<tr><td colspan="2">Salvar cópia</td></tr>
	<tr>
		<td>Nome do arquivo</td>
		<td><input type="text" name="nome_arquivo" size="30" maxlength="30"></td>
	</tr>
	<tr><td colspan="2"><input type="submit" name="ok" value="salva"></td></tr>
	</form>
</table>
<%else
Set FSO = Server.CreateObject("Scripting.FileSystemObject")


caminho = Server.MapPath("../gravados/"&nome_gravacao&".txt") 'especifique aqui o caminho onde ficará/está o TXT
Set GRAVAR = FSO.CreateTextFile(caminho,true)
'Foi criado o objeto e logo após busca o txt em caminho para gravar, se não achar, vai cria-lo (por causa da marcação TRUE)
	
	'*** grava no banco de dados ***
	'if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
		'xsql = "up_grava_bkp_insert @cd_origem="&origem&", @dt_gravacao='"&mid(dt_gravacao,3,4)&"/"&mid(dt_gravacao,5,6)&"/20"&mid(dt_gravacao,1,2)&" "&mid(dt_gravacao,7,8)&":"&mid(dt_gravacao,9,10)&"', @cd_usuario='"&usuario&"', @nm_arquivo='"&nome_gravacao&"'"
		'xsql = "up_grava_bkp_insert @cd_origem="&origem&", @dt_gravacao='"&mid(dt_gravacao,3,4)&"/"&mid(dt_gravacao,5,6)&"/20"&mid(dt_gravacao,1,2)&"', @cd_usuario='"&usuario&"', @nm_arquivo='"&nome_gravacao&"'"
		xsql = "up_grava_bkp_insert @cd_origem="&origem&", @dt_gravacao='"&month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)&":"&second(now)&"', @cd_usuario='"&usuario&"', @nm_arquivo='"&nome_gravacao&"'"
		dbconn.execute(xsql)
	'end if
	
	
gravar.write (request("conteudo"))
gravar.close
response.write "GRAVADO!"&nome_gravacao&""
'apos abrir o TXT, gravará a linha com o texto "TESTE DE GRAVAÇÃO" a confirmação no cliente aparecerá como "GRAVADO"
%>

<!--body onload="window.close();"-->
<%end if%>
<%'=replace(grava_txt,"$","<br>"%>
</body>
