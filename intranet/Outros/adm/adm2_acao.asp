
  <!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
   <!--include file="includes/inc_area_restrita.asp"-->

<%

acao = request("acao")
campo = request("campo")

tipo = request("tipo")
janela = request("janela")

'Avaliações
nm_avaliacao = request("avaliacao")

'Fornecedores
nm_fornecedor = request("nm_fornecedor")
nm_contato = request("nm_contato")
nm_contato2 = request("nm_contato2")
nm_endereco = request("nm_endereco")
'nm_telefone_cod1 = request("nm_telefone_cod1")
'nm_telefone1 = request("nm_telefone1")
'nm_telefone_cod2 = request("nm_telefone_cod2")
'nm_telefone2 = request("nm_telefone2")
nm_numero = request("nm_numero")
nm_bairro = request("nm_bairro")
nm_cidade = request("nm_cidade")
nm_estado = request("nm_estado")

cd_prazo = request("cd_prazo")

'Equipamentos
cd_pn = request("cd_pn")
nm_equipamento = request("nm_equipamento")
'nm_marca = request("nm_marca")

'Especialidades
nm_especialidade = request("nm_especialidade")

'Marcas
nm_marca = request("nm_marca")

'Unidades
nm_sigla = request("nm_sigla")
nm_unidade = request("nm_unidade")

'Cargo
cd_cargo = request("cd_cargo")
nm_cargo = request("nm_cargo")
cd_especificidade = request("cd_especificidade")

'Salario
cd_salario = request("cd_salario")

'Ajustes
cd_codigo = request("cd_codigo")
dia = request("dia")
mes = request("mes")
ano = request("ano")
hora = hour(now)
minuto = minute(now)
nm_status = request("nm_status")
cd_funcao = request("cd_funcao")
cd_unidade = request("cd_unidade")

'Data
dt_data = mes&"/"&dia&"/"&ano

' ****  AÇÕES  ****
	
'********************* 
'***  Fornecedores ***
'*********************
	If acao = "inserir" AND tipo = "forn" Then
	'Insere dados
	'xsql = "up_fornecedor_insert @nm_fornecedor='"&nm_fornecedor&"', @nm_contato='"&nm_contato&"', @nm_contato2='"&nm_contato2&"', @nm_endereco='"&nm_endereco&"', @nm_telefone_cod1='"&nm_telefone_cod1&"', @nm_telefone1='"&nm_telefone1&"', @nm_telefone_cod2='"&nm_telefone_cod2&"', @nm_telefone2='"&nm_telefone2&"'"
	xsql = "up_fornecedor_insert @nm_fornecedor='"&nm_fornecedor&"', @nm_contato='"&nm_contato&"', @nm_contato2='"&nm_contato2&"', @nm_endereco='"&nm_endereco&"', @nm_numero = '"&nm_numero&"', @nm_bairro='"&nm_bairro&"', @nm_cidade='"&nm_cidade&"',@nm_estado='"&nm_estado&"', @cd_prazo='"&cd_prazo&"'"
	action = "Fornecedor inserido com sucesso"
	dbconn.execute(xsql)
	redirecionamento = "../fornecedores.asp?action="&action&""
	
	ElseIf acao = "editar" AND tipo = "forn" Then
	'Modifica os dados
	xsql = "up_fornecedor_update @cd_codigo='"&cd_codigo&"', @nm_fornecedor='"&nm_fornecedor&"', @nm_contato='"&nm_contato&"', @nm_contato2='"&nm_contato2&"', @nm_endereco='"&nm_endereco&"', @nm_numero = '"&nm_numero&"', @nm_bairro='"&nm_bairro&"', @nm_cidade='"&nm_cidade&"',@nm_estado='"&nm_estado&"', @cd_prazo='"&cd_prazo&"'"
	action = "Fornecedor editado com sucesso"
	dbconn.execute(xsql)
	'response.redirect("../fornecedores.asp?action="&action&"")
	redirecionamento = "../fornecedores.asp?action="&action&""
	
	ElseIf acao = "excluir" AND tipo = "forn" Then
	xsql = "up_fornecedor_delete @cd_codigo='"&cd_codigo&"'"
	dbconn.execute(xsql)
	action = "Fornecedor excluído com sucesso"
	'response.redirect("../fornecedores.asp?action="&action&"")
	redirecionamento = "../fornecedores.asp?action="&action&""
	
'********************
'*** Equipamentos ***
'********************
	ElseIf acao = "inserir" AND tipo = "equip" Then
	'Insere dados
	xsql = "up_equipamento_insert @cd_pn='"&cd_pn&"', @nm_equipamento='"&nm_equipamento&"'"
	action = "Assistência inserida com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		response.redirect("../equipamentos.asp?action="&action&"")
		end if
	
	ElseIf acao = "editar" AND tipo = "equip" Then
	'Modifica os dados
	xsql = "up_manutencao_update @cd_pn='"&cd_pn&"', @nm_equipamento='"&nm_equipamento&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	response.redirect("equipamento.asp?action="&action&"")



	'Especialidades
	Elseif acao = "inserir" AND tipo = "espec" Then
	'Insere dados
	xsql = "up_especialidade_insert @nm_especialidade='"&nm_especialidade&"'"
	action = "Especialidade inserida com sucesso"
	dbconn.execute(xsql)
	response.redirect("especialidade.asp?action="&action&"")
	
	ElseIf acao = "editar" AND tipo = "espec" Then
	'Modifica os dados
	xsql = "up_espec_update @nm_especialidade='"&nm_especialidade&"'"
	action = "editada com sucesso"
	dbconn.execute(xsql)
	response.redirect("especialidade.asp?action="&action&"")



'**************
'*** Marcas ***
'**************
	'Marcas
	Elseif acao = "inserir" AND tipo = "marca" Then
	'Insere dados
	xsql = "up_marca_insert @nm_marca='"&nm_marca&"'"
	action = "Especialidade inserida com sucesso"
	dbconn.execute(xsql)
	redirecionamento = "../marcas.asp?action="&action&""
	
	ElseIf acao = "editar" AND tipo = "marca" Then
	'Modifica os dados
	xsql = "up_marca_update @nm_marca='"&nm_marca&"'"
	action = "editada com sucesso"
	dbconn.execute(xsql)
	redirecionamento = "../marcas.asp?action="&action&""
	
	
	'Unidades
	Elseif acao = "inserir" AND tipo = "unidade" Then
	'Insere dados
	xsql = "up_unidade_insert @nm_sigla='"&nm_sigla&"', @nm_unidade='"&nm_unidade&"'"
	action = "unidade inserida com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		response.redirect("../unidades.asp?action="&action&"")
		end if
	
	
	
	'Cargo
	Elseif acao = "inserir" AND tipo = "cargo" Then
	'Insere dados
	xsql = "up_cargo_insert @nm_cargo='"&nm_cargo&"', @cd_especificidade='"&cd_especificidade&"'"
	action = "cargo inserido com sucesso"
	dbconn.execute(xsql)
	
			'Salário
		'lista o codigo do cargo recém registrado.
		strsql = "Select * From TBL_funcoes Where nm_funcao = '"&nm_cargo&"'"
		Set rs = dbconn.execute(strsql)
		nova_funcao = rs("cd_codigo")
		'dt_data = "01/01/2008"
			'registra o salário referente ao novo cargo
			'response.write("insere valor do salário")
			xsql = "up_cargo_salario_insert @cd_cargo = '"&nova_funcao&"', @cd_valor = '"&cd_salario&"', @cd_tipo = 1, @dt_data = '"&dt_data&"'"
			dbconn.execute(xsql)
			response.redirect("../adm_cargos.asp")
	
	Elseif acao = "editar" AND tipo = "cargo" Then
	'Atualiza dados (insere valor e data novos)
	'xsql = "up_cargo_valorupdate @nm_cargo='"&nm_cargo&"', @cd_especificidade='"&cd_especificidade&"'"
	'action = "cargo inserido com sucesso"
	'dbconn.execute(xsql)
	response.write("insere valor do salário")
	xsql = "up_cargo_salario_insert @cd_cargo = '"&cd_cargo&"', @cd_valor = '"&cd_salario&"', @cd_tipo = 1, @dt_data = '"&dt_data&"'"
	dbconn.execute(xsql)
			
	'response.write("ok, mano!")
	response.redirect("../adm_cargos.asp")
	
	'Ajuste
	Elseif acao = "editar" AND tipo = "ajuste" Then
	xsql = "up_relatorio_cooperativa_ajuste @cd_codigo='"&cd_codigo&"', @cd_funcao='"&cd_funcao&"', @nm_status='"&nm_status&"', @cd_unidade='"&cd_unidade&"', @mes='"&mes&"', @ano='"&ano&"'"
	dbconn.execute(xsql)
	'response.redirect("../produtividade/unidades.asp?action="&action&"")
	'response.write("Ajuste ok")%>
	<%
	Else 
	action = "erro"
	End If
	
	if janela = "" Then
	response.redirect(redirecionamento)
	end if

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
Tipo: <%=tipo%><br><br>

Código: <%=cd_codigo%><br>
Dia: <%=dia%><br>
Mes: <%=mes%><br>
Ano: <%=ano%><br>
data: <%=dt_data%>
Status: <%=nm_status%><br>
Funcao: <%=cd_funcao%><br>
Unidade: <%=cd_unidade%><br>
cargo: <%=nm_cargo%><br>
espec: <%=cd_especificidade%><br><br>

nova funçao = <%=nova_funcao&cd_cargo%><br>
(salario = <%=cd_salario%><br>
<br>
fornecedor = <%=nm_fornecedor%><br>
nm_contato = <%=nm_contato%><br>
nm_contato2 = <%=nm_contato2%><br>
nm_endereco = <%=nm_endereco%><br>
nm_numero = <%=nm_numero%><br>
nm_bairro = <%=nm_bairro%><br>
nm_cidade = <%=nm_cidade%><br>
nm_estado = <%=nm_estado%><br><br>

Marca = <%=nm_marca%><br>





</body>
</html>
