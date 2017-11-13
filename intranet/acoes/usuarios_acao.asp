
  <!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
   <!--#include file="../includes/inc_area_restrita.asp"-->

<%

cd_cod_usuario = request("cd_cod_usuario")

acao = request("acao")
nm_nome = request("nm_nome")
nm_email = request("nm_email")
nm_senha = request("nm_senha")

menu_1 = request("menu_1")
menu_2 = request("menu_2")
menu_3 = request("menu_3")

cd_status = request("cd_status")

'nm_email = encrypt(nm_email)
'nm_senha = encrypt(nm_senha)

'*************************************************************************************************
'************************ 1.   Cadastra    *******************************************************
'*************************************************************************************************
If acao = "1" AND cd_cod_usuario = "" Then
'Insere dados - Usuário
xsql = "up_usuario_insert @nm_nome='"&nm_nome&"', @nm_email='"&nm_email&"', @nm_senha='"&nm_senha&"', @dt_ultimo_acesso='"&dt_ultimo_acesso&"', @nr_qtd_erro='"&nr_qtd_erro&"', @nr_qtd_acasso='"&nr_qtd_acasso&"'"
dbconn.execute(xsql)

		'Verifica e pega o código do usuário para inserção das permissões
		xsql2 = "SELECT * FROM TBL_ADM_usuario WHERE nm_nome='"&nm_nome&"' AND nm_email = '"&nm_email&"'"
		Set rs2 = dbconn.execute(xsql2)
		cd_cod = rs2("cd_codigo")
		
		'Registra permissões
		xsql = "up_usuario_permissoes_insert @cd_cod='"&cd_cod&"', @menu_1='"&menu_1&"', @menu_2='"&menu_2&"', @menu_3='"&menu_3&"'"
		dbconn.execute(xsql)

'response.redirect("../usuarios.asp?cd_codigo=&msg=ok")
'response.redirect("ADM_cadastro.asp?cd_codigo="&cd_codigo&"&msg=ok")
'response.Write("Usuário cadastrado com Sucesso!")

'*************************************************************************************************
'************************ 2.  Modifica     *******************************************************
'*************************************************************************************************
Elseif acao = "2" AND cd_cod_usuario <> "" Then
'Modifica os dados
xsql = "up_usuario_update @cd_codigo='"&cd_cod_usuario&"', @nm_senha='"&nm_senha&"', @cd_status='"&cd_status&"'"
dbconn.execute(xsql)

		'*** Verifica se o usuário tem permissoes ***
		xsql_vpermissoes ="SELECT * FROM tbl_ADM_usuario_permissoes where cd_usuario='"&cd_cod_usuario&"'"
		Set rs_vpermissoes = dbconn.execute(xsql_vpermissoes)
			if rs_vpermissoes.EOF Then 
				'Registra permissões
				xsql = "up_usuario_permissoes_insert @cd_cod='"&cd_cod_usuario&"', @menu_1='"&menu_1&"', @menu_2='"&menu_2&"', @menu_3='"&menu_3&"'"
				dbconn.execute(xsql)
				msn="inseriu"
			Elseif not rs_vpermissoes.EOF then
				'Altera permissões
				xsql = "up_usuario_permissoes_update @cd_cod='"&cd_cod_usuario&"', @menu_1='"&menu_1&"', @menu_2='"&menu_2&"', @menu_3='"&menu_3&"'"
				dbconn.execute(xsql)
				msn="alterou"
			end if

response.redirect("../usuarios.asp?tipo=lista")
'response.write("Usuário modificado com sucesso")

'******************************* Exclui *************************************************

Elseif acao = "3" AND cd_cod_usuario <> "" Then
'Apaga Usuário
xsql = "up_usuario_delete @cd_codigo='"&cd_cod_usuario&"'"
action = "apagado com sucesso"
dbconn.execute(xsql)
	'Apaga permissões
		xsql = "up_usuario_permissoes_excluir @cd_usuario='"&cd_cod_usuario&"'"
		dbconn.execute(xsql)
		msn = "excluído"
response.redirect("../usuarios.asp?tipo=lista")


End If

'='"&Month()&"-"&day()&"-"&year()&"'


'response.redirect("manutencao.asp?cd_codigo="&cd_codigo&"&campo="&campo&"")

'response.redirect("manutencao.asp?cd_codigo=&campo=dt_os")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>

<body>
MSN - <%=msn%><br>
Ação - <%=acao%><br>
Codigo - <%=cd_codigo%><br>
Usuário - <%=cd_cod_usuario%><br><br>

nome - <%=nm_nome%><br>
email - <%=nm_email%><br>
senha - <%=nm_senha%><br><br>

Menu 1 = <%=menu_1%><br>
Menu 2 = <%=menu_2%><br>
Menu 3 = <%=menu_3%><br>

</body>
</html>
