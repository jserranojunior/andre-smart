
  <!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
   <!--#include file="../includes/inc_area_restrita.asp"-->

<%

acao = request("acao")
cd_tipo = request("cd_tipo")
cd_funcionario = request("cd_funcionario")
cd_apaga = request("cd_apaga")

cd_pat_codigo = request("cd_pat_codigo")
cd_equipamento = request("cd_equipamento")
cd_pn = request("cd_pn")
cd_ns = request("cd_ns")
cd_patrimonio = request("cd_patrimonio")

nm_motivo = request("nm_motivo")
cd_unidade = request("cd_unidade")
cd_status = request("cd_status")
cd_marca =request("cd_marca")

cd_dia = request("cd_dia")
cd_mes = request("cd_mes")
cd_ano = request("cd_ano")

cd_hora = hour(now)
cd_minuto = minute(now)
cd_segundo = second(now)


	dt_atualizacao = cd_mes&"/"&cd_dia&"/"&cd_ano&" "&cd_hora&":"&cd_minuto&":"&cd_segundo
	dt_data = dt_atualizacao
	dt_descarte = dt_atualizacao





voltar = request("voltar")

if cd_apaga <> "" Then
	acao = "apaga"
end if

'Cargo
cd_funcao = request("cd_funcao")

' ****  AÇÕES  ****
	
	'Cargo
	If acao = "inserir" AND cd_tipo = "cargo" Then
	'Insere dados
	xsql = "up_funcionario_cargo_insert @cd_funcionario='"&cd_funcionario&"', @cd_cargo='"&cd_funcao&"', @dt_atualizacao='"&dt_atualizacao&"'"
	dbconn.execute(xsql)
	'response.redirect("assistencia.asp?action="&action&"")
	
	'ElseIf acao = "editar" AND cd_tipo = "cargo" Then
	''Modifica os dados
	'xsql = "up_fornecedores_update @cd_pn='"&cd_pn&"', @nm_equipamento='"&nm_equipamento&"'"
	'action = "editado com sucesso"
	'dbconn.execute(xsql)
	''response.redirect("assistencia.asp?action="&action&"")
	
	
	
	'Unidade
	ElseIf acao = "inserir" AND cd_tipo = "unidade" Then
	'Insere dados
	xsql = "up_funcionario_unidade_insert @cd_funcionario='"&cd_funcionario&"', @cd_unidade='"&cd_unidade&"', @dt_atualizacao='"&dt_atualizacao&"'"
	dbconn.execute(xsql)
	'response.redirect("assistencia.asp?action="&action&"")
		xsql="Select nm_sigla from TBL_unidades Where cd_codigo='"&cd_unidade&"'"
		Set rs = dbconn.execute(xsql)
		nm_sigla = rs("nm_sigla")
			'Atualiza a sigla da unidade na tabela de Funcionarios
			xsql = "UPDATE TBL_funcionario SET nm_sigla='"&nm_sigla&"' WHERE cd_codigo='"&cd_funcionario&"'"
			dbconn.execute(xsql)
	
	'Status
	ElseIf acao = "inserir" AND cd_tipo = "status" Then
	'Insere dados
	xsql = "up_funcionario_status_insert @cd_funcionario='"&cd_funcionario&"', @cd_status='"&cd_status&"', @dt_atualizacao='"&dt_atualizacao&"'"
	dbconn.execute(xsql)
	'response.redirect("assistencia.asp?action="&action&"")
	
	
	'Patrimonio
	Elseif acao = "inserir" AND cd_tipo = "patrimonio" Then
	'Insere dados
	xsql = "up_patrimonio_insert @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @cd_patrimonio='"&cd_patrimonio&"', @cd_unidade='"&cd_unidade&"', @dt_data='"&dt_data&"'"
	dbconn.execute(xsql)
	response.write("patrimonio insere OK")
	response.redirect("../patrimonio_cadastro.asp?mensagem=Inserido com sucesso")
	
	ElseIf acao = "editar" AND cd_tipo = "patrimonio" Then
	'Modifica os dados
	xsql = "up_patrimonio_update @cd_pat_codigo='"&cd_pat_codigo&"', @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @cd_patrimonio='"&cd_patrimonio&"', @cd_unidade='"&cd_unidade&"', @dt_data='"&dt_data&"'"
	dbconn.execute(xsql)
	response.write("patrimonio edita OK")
	response.redirect("../patrimonio_cadastro.asp?mensagem=Alterado com sucesso")
	
			ElseIf acao = "descartar" AND cd_tipo = "patrimonio" Then
			'Modifica os dados
			xsql = "up_patrimonio_descarte @cd_pat_codigo='"&cd_pat_codigo&"', @dt_descarte='"&dt_data&"', @nm_motivo='"&nm_motivo&"'"
			dbconn.execute(xsql)
			response.write("Descarte OK")
			response.redirect("../patrimonio_descarte.asp?mensagem=Alterado com sucesso")
			
	
	Elseif acao = "apaga" AND cd_tipo = "patrimonio" Then
	'Apaga os dados
	xsql = "up_patrimonio_delete @cd_apaga='"&cd_apaga&"'"
	dbconn.execute(xsql)
	response.write("patrimonio apaga OK")
	response.redirect("../patrimonio_cadastro.asp?mensagem=Apagado com sucesso")
	
	
	
	'Adicionais
	ElseIf acao = "inserir" AND cd_tipo > "0" Then
	'Insere dados
	xsql = "up_funcionario_noturno_insert @cd_funcionario='"&cd_funcionario&"', @dt_mes='"&cd_mes&"', @dt_ano='"&cd_ano&"', @cd_tipo='"&cd_tipo&"'"
	dbconn.execute(xsql)
	'response.redirect(voltar)
	
	Else 
	action = "erro"
	End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>

<body onLoad=window.close()>
<br>
Mensagem = <%=action%><br>
Ação: <%=acao%><br>
Tipo: <%=cd_tipo%><br><br>
Funcionario: <%=cd_funcionario%><br><br>
Funcao: <%=cd_funcao%><br>
Mes/Ano: <%=cd_mes%>/<%=cd_ano%><br>
Apagar: <%=cd_apaga%><br>

<br>
Num fab.: <%=cd_pn%><br>
Serie: <%=cd_ns%><br>
Eqp.: <%=cd_equipamento%><br>
Unidade: <%=cd_unidade%><br>
Marca: <%=cd_marca%><br>
<br>
Descarte: <%=dt_descarte%><br>
Motivo: <%=nm_motivo%>





</body>
</html>
