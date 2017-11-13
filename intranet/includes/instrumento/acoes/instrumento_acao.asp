<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<!--#include file="../../includes/inc_open_connection.asp"-->
<%
janela = request("janela")
tipo = request("tipo")
cd_inst_codigo = request("cd_inst_codigo")
acao = request("acao")
caminho = request("caminho")

cd_unidade = request("cd_unidade")
cd_equipamento = request("cd_equipamento")
nm_modelo = request("nm_modelo")
cd_marca = request("cd_marca")
nm_pn = request("nm_pn")
cd_categoria = request("cd_categoria")
cd_qtd = request("cd_qtd")
nm_descricao = request("nm_descricao")

cd_user = session("cd_codigo")
dt_cadastro = "'"&month(now)&"/"&day(now)&"/"&year(now)&"'" '"&minute(now)&":"&hour(now)&"'"

' ****  AÇÕES  ****
	'Instrumento
	If acao = "inserir" AND tipo = "instrumento" Then
		'Insere dados
		xsql = "up_instrumento_insert @nm_pn='"&nm_pn&"', @cd_unidade="&cd_unidade&", @cd_equipamento="&cd_equipamento&", @nm_modelo="&nm_modelo&", @cd_marca="&cd_marca&", @nm_descricao="&nm_descricao&", @cd_categoria="&cd_categoria&", @cd_user="&cd_user&", @dt_cadastro="&dt_cadastro&""
		'xsql = "up_instrumento_insert @nm_pn="", @cd_unidade="", @cd_equipamento="", @nm_modelo="", @cd_marca="", @nm_descricao="", @cd_categoria="", @cd_user="", @dt_cadastro="" "
		dbconn.execute(xsql)
		response.write("instrumento insere OK")
			
		'response.redirect(caminho&"../../instrumento.asp?tipo=cadastro&mensagem=Inserido com sucesso")
	end if

'@nm_pn, @cd_unidade, @cd_equipamento, @nm_modelo, @cd_marca, @nm_descricao, @cd_categoria, @cd_user, @dt_cadastro

%>
<html>
<head>
	<title>Untitled</title>
</head>

<body>
<br>
ação = <%=acao%><br>
tipo = <%=tipo%><br>
<br>
User = <%=cd_user%><br>
Data = <%=dt_cadastro%><br>
<br>

Janela = <%=janela%><br><br>

cd_inst_codigo = <%=cd_inst_codigo%><br>
acao = <%=acao%><br>
<br>
cd_unidade = <%=cd_unidade%><br>
cd_equipamento = <%=cd_equipamento%><br>
nm_modelo = <%=nm_modelo%><br>
cd_marca <%=cd_marca%><br>
nm_pn = <%=nm_pn%><br>
cd_categoria = <%=cd_categoria%><br>
cd_qtd = <%=cd_qtd%><br>
nm_descricao = <%=nm_descricao%><br>

</body>
</html>
