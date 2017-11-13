
  <!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
   <!--include file="includes/inc_area_restrita.asp"-->

<%cd_user = session("cd_codigo")

acao = request("acao")
campo = request("campo")

modo = request("modo")
jan = request("jan")
metodo = request("metodo")

'Data Atual
data_atual = Month(now)&"/"&Day(now)&"/"&Year(now)&" "&Hour(now)&":"&Minute(now)&":"&Second(now)

'*** Variaveis compartilhadas ***
cd_status = request("cd_status")

nm_endereco = request("nm_endereco")
nm_numero = request("nm_numero")
nm_cep = request("nm_cep1")&"-"&request("nm_cep2")
'nm_cep = nm_cep&"-"&request("nm_cep2")
nm_bairro = request("nm_bairro")
nm_cidade = request("nm_cidade")
nm_estado = request("nm_estado")
nm_endereco_completo = nm_endereco&"!"&nm_numero&"#"&nm_cep&"$"&nm_bairro&"["&nm_cidade&"]"&nm_estado
nm_contato = request("nm_contato")
nm_telefone = request("nm_telefone")

nm_cnpj = request("nm_cnpj")
nm_ie = request("nm_ie")
cd_empresa = request("cd_empresa")

cd_status = request("cd_status")
'********************************

'Avaliações
nm_avaliacao = request("avaliacao")

'Fornecedores
cd_fornecedor = request("cd_fornecedor")
nm_fornecedor = request("nm_fornecedor")
cd_prazo = request("cd_prazo")



'Equipamentos
cd_pn = request("cd_pn")
nm_equipamento = request("nm_equipamento")
cd_equipamento = request("cd_equipamento")
cd_tipo = request("cd_tipo")
cd_ns = request("cd_ns")

'Unidades
cd_unidade = request("cd_unidade")
nm_sigla = request("nm_sigla")
nm_unidade = request("nm_unidade")
nm_unidade_nome = request("nm_unidade_nome")
cd_unidade_tipo = request("cd_unidade_tipo")
dt_inicio_contrato = "'"&request("dt_i_contrato_mes")&"/"&request("dt_i_contrato_dia")&"/"&request("dt_i_contrato_ano")&"'"
dt_fim_contrato = request("dt_f_contrato_mes")&"/"&request("dt_f_contrato_dia")&"/"&request("dt_f_contrato_ano")
	if trim(dt_fim_contrato) = "//" Then
		dt_fim_contrato = "NULL"
	else	
		dt_fim_contrato = "'"&dt_fim_contrato&"'"
		cd_status = "0"
	end if
nm_cor = request("nm_cor")
cd_prazo_faturamento = request("cd_prazo_faturamento")

'Especialidades
cd_especialidade = request("cd_especialidade")
nm_especialidade = request("nm_especialidade")


'Marcas
cd_marca = request("cd_marca")
nm_marca = request("nm_marca")

'Patrimonio
cd_patrimonio = request("cd_patrimonio")
dt_descarte = request("dt_descarte")
nm_motivo = request("nm_motivo")

'Situacao
nm_situacao = request("nm_situacao")

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

'Assistencias Médicas
cd_convenio = request("cd_convenio")
nm_convenio = request("nm_convenio")

'Medicos
cd_medico = request("cd_medico")
cd_crm = request("cd_crm")
nm_medico = request("nm_medico")
status = request("status")

'Rack
cd_rack = request("cd_rack")
nm_rack = request("nm_rack")

'Procedimento
cod_amb = request("cod_amb")
nm_procedimento = request("nm_procedimento")
cd_procedimento = request("cd_procedimento")

'Material
cd_material = request("cd_material")
nm_material = request("nm_material")
cd_material = request("cd_material")

'Empresa
nm_empresa = request("nm_empresa")
dt_abertura_dia = request("dt_abertura_dia")
dt_abertura_mes = request("dt_abertura_mes")
dt_abertura_ano = request("dt_abertura_ano")
dt_fechamento_dia = request("dt_fechamento_dia")
dt_fechamento_mes = request("dt_fechamento_mes")
dt_fechamento_ano = request("dt_fechamento_ano")

'Data
dt_datacad = request("dt_datacad")
dt_datatu = request("dt_datatu")
dt_data = mes&"/"&dia&"/"&ano






if cd_user <> 460 then

' ****  AÇÕES  *******************************************************************************
	
	'*******************
	'*** Equipamento ***
	'*******************
	If acao = "inserir" AND modo = "equipamento" then
	'Insere dados
	xsql = "up_equipamento_insert @nm_equipamento='"&nm_equipamento&"', @cd_pn='"&cd_pn&"',@cd_tipo= '"&cd_tipo&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "Equipamento inserido com sucesso"
	dbconn.execute(xsql)
		if jan = "" Then
		response.redirect("../ferramentas.asp?tipo=equipamento&status=1")
		end if
		
	Elseif acao = "editar" AND modo = "equipamento" then
	'Altera dados
		If status = 1 then
			xsql = "up_equipamento_update @cd_equipamento='"&cd_equipamento&"', @nm_equipamento='"&nm_equipamento&"', @cd_pn='"&cd_pn&"', @cd_tipo='"&cd_tipo&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Equipamento inserido com sucesso"
			dbconn.execute(xsql)
			response.write("editado!!!")
				if jan = "" Then
				response.redirect("../ferramentas.asp?tipo=equipamento&status=1")
				end if
				
		elseif status = 0 then
			xsql = "up_equipamento_inativar @cd_equipamento='"&cd_equipamento&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O equipamento foi reativado"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if jan = "" Then
				response.redirect("../ferramentas.asp?tipo=equipamento&status=1")
				end if
		end if
		
	Elseif acao = "excluir" AND modo = "equipamento" then
	'Exclui dados
		if status = 2 then
			xsql = "up_equipamento_delete @cd_equipamento='"&cd_equipamento&"'"
			action = "Equipamento exluído com sucesso"
			response.write("apagado!!!")
			dbconn.execute(xsql)
				if jan = "" Then
				response.redirect("../ferramentas.asp?tipo=equipamento&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_equipamento_inativar @cd_equipamento='"&cd_equipamento&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O equipamento está inativo"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if jan = "" Then
				response.redirect("../ferramentas.asp?tipo=equipamento&status=1")
				end if
		end if
	
	
	'***************
	'*** Unidade ***
	'***************
	Elseif acao = "inserir" AND modo = "unidade" then
		'Insere dados
		xsql = "up_unidade_insert @nm_unidade_nome='"&nm_unidade_nome&"', @nm_unidade='"&nm_unidade&"', @nm_sigla='"&nm_sigla&"',@nm_endereco= '"&nm_endereco_completo&"', @dt_inicio_contrato="&dt_inicio_contrato&", @nm_cnpj='"&nm_cnpj&"', @nm_ie='"&nm_ie&"', @cd_hospital='"&cd_unidade_tipo&"', @cd_cliente_empresa='"&cd_empresa&"', @nm_cor='"&nm_cor&"', @cd_prazo_faturamento='"&cd_prazo_faturamento&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
		dbconn.execute(xsql)
		'Após incluir unidade, fecha a janela
		
	Elseif acao = "editar" AND modo = "unidade" then
		xsql = "up_unidade_update @nm_unidade_nome='"&nm_unidade_nome&"', @cd_unidade='"&cd_unidade&"', @nm_unidade='"&nm_unidade&"', @nm_sigla='"&nm_sigla&"',@nm_endereco= '"&nm_endereco_completo&"',@nm_contato='"&nm_contato&"', @nm_telefone='"&nm_telefone&"',@dt_inicio_contrato="&dt_inicio_contrato&", @dt_fim_contrato="&dt_fim_contrato&", @nm_cnpj='"&nm_cnpj&"', @nm_ie='"&nm_ie&"', @cd_hospital='"&cd_unidade_tipo&"', @cd_cliente_empresa='"&cd_empresa&"', @nm_cor='"&nm_cor&"', @cd_status='"&cd_status&"', @cd_prazo_faturamento='"&cd_prazo_faturamento&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
		dbconn.execute(xsql)
		'response.redirect("../unidade_cad.asp?cd_unidade="&cd_unidade&"&status="&cd_status&"&acao=editar&jan=1")
		'Após incluir unidade, fecha a janela
	
	Elseif acao = "ativar" AND modo = "unidade" then
		xsql = "up_unidade_ativar_desativar @cd_unidade='"&cd_unidade&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
		action = "Unidade foi reativada"
		dbconn.execute(xsql)
		response.redirect("../../ferramentas.asp?tipo=unidade&status=0")
		
	Elseif acao = "inativar" AND modo = "unidade" then
		xsql = "up_unidade_ativar_desativar @cd_unidade='"&cd_unidade&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
		action = "Unidade foi inativada"
		dbconn.execute(xsql)
		response.redirect("../../ferramentas.asp?tipo=unidade&status=1")
		
	Elseif acao = "excluir" AND modo = "unidade" then
		xsql = "up_unidade_delete @cd_unidade='"&cd_unidade&"'"
		action = "Unidade exluída com sucesso"
		dbconn.execute(xsql)
		response.redirect("../../ferramentas.asp?tipo=unidade&status=1")
	
	
	'*******************
	'*** Fornecedor ***
	'*******************
	ElseIf acao = "inserir" AND modo = "fornecedor" then
	'Insere dados
	xsql = "up_fornecedor_insert @nm_fornecedor='"&nm_fornecedor&"', @nm_endereco='"&nm_endereco&"', @nm_numero='"&nm_numero&"', @nm_bairro='"&nm_bairro&"', @nm_cep='"&nm_cep&"', @nm_cidade='"&nm_cidade&"', @nm_estado='"&nm_estado&"', @nm_cnpj='"&nm_cnpj&"', @nm_ie='"&nm_ie&"', @cd_prazo='"&cd_prazo&"', @cd_status = 1"
	action = "Equipamento inserido com sucesso"
	dbconn.execute(xsql)
		'if jan = "" Then
		'response.redirect("../ferramentas.asp?tipo=fornecedor&status=1")
		'end if
	
	Elseif acao = "editar" AND modo = "fornecedor" then
	'Altera dados
		If cd_status = 1 then
			xsql = "up_fornecedor_update @cd_fornecedor='"&cd_fornecedor&"', @nm_fornecedor='"&nm_fornecedor&"', @nm_endereco='"&nm_endereco&"', @nm_numero='"&nm_numero&"', @nm_bairro='"&nm_bairro&"', @nm_cep='"&nm_cep&"', @nm_cidade='"&nm_cidade&"', @nm_estado='"&nm_estado&"', @nm_cnpj='"&nm_cnpj&"', @nm_ie='"&nm_ie&"', @cd_prazo='"&cd_prazo&"', @cd_status='"&cd_status&"'"
			dbconn.execute(xsql)
				if jan = "" Then
				'response.redirect("../ferramentas.asp?tipo=fornecedores&status=1")
				end if
		end if
	
	
	ELSEIf acao = "ativar" AND modo = "fornecedor" then
		'elseif cd_status = 0 then
			xsql = "up_fornecedor_inativar @cd_fornecedor='"&cd_fornecedor&"', @cd_status=1"
			action = "O equipamento foi reativado"
			dbconn.execute(xsql)
				if jan = "" Then
				response.redirect("../../ferramentas.asp?tipo=fornecedores&status=0")
				end if
		'end if
	elseif acao = "desativar" AND modo = "fornecedor" then
		'elseif cd_status = 0 then
			xsql = "up_fornecedor_inativar @cd_fornecedor='"&cd_fornecedor&"', @cd_status=0"
			action = "O equipamento foi desativado"
			dbconn.execute(xsql)
				if jan = "" Then
				response.redirect("../../ferramentas.asp?tipo=fornecedores&status=1")
				end if
	
	
	Else 
	action = "erro"
	End If
	
		if jan = "" Then
		'response.redirect(redirecionamento)
		end if%>
	
<%end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>cadastros</title>
</head>
<%'if jan = "1" Then%>
	<body onLoad="window.close()">
<%'else%>
	<!--body-->
<%'end if%>
<br>
Mensagem = <%=action%><br>
Ação: <%=acao%><br>
modo: <%=modo%><br>
jan: <%=jan%><br>
metodo: <%=metodo%><br>
<br>

Código: <%=cd_codigo%><br>
Dia: <%=dia%><br>
Mes: <%=mes%><br>
Ano: <%=ano%><br>
data: <%=dt_data%><br>
Status: <%=nm_status%><br>
Funcao: <%=cd_funcao%><br>
Unidade: (<%=cd_unidade%>):<%=nm_sigla%> - <%=nm_unidade%><br>
cargo: <%=nm_cargo%><br>
espec: <%=cd_especificidade%><br><br>

nova funçao = <%=nova_funcao&cd_cargo%><br>
(salario = <%=cd_salario%><br>
<br>
cod_fornecedor = <%=cd_fornecedor%><br>
fornecedor = <%=nm_fornecedor%><br>
nm_contato = <%=nm_contato%><br>
nm_contato2 = <%=nm_contato2%><br>
nm_endereco = <%=nm_endereco%><br>
nm_numero = <%=nm_numero%><br>
nm_bairro = <%=nm_bairro%><br>
nm_cidade = <%=nm_cidade%><br>
nm_estado = <%=nm_estado%><br><br>
cd_status = <%=cd_status%><br>
cnpj = <%=nm_cnpj%><br>
I.E = <%=nm_ie%><br>

<br>

Marca = <%=nm_marca%><br><br>

Rack = <%=nm_rack%><br>
especialidade = <%=nm_especialidade%><br><br>

Cod. AMB = <%=cod_amb%><br>
Procedimento = <%=nm_procedimento%><br>
cod_ proc = <%=cd_procediemnto%><br>

Material = <%=nm_material%><br><br>

Medico = <%=cd_medico%> *(<%=cd_crm%>) - <%=nm_medico%> - <%=status%><br><br>
Convenio = <%=cd_convenio%> - <%=nm_convenio%> - <%=status%><br><br>
Material = <%=cd_material%> - <%=nm_material%><br><br>




Data = <%=data_atual%><br>
User = <%=cd_user%><br><br>


cd_unidade = <%=cd_unidade%><br>
nm_sigla = <%=nm_sigla%><br>
nm_unidade = <%=nm_unidade%><br>
cd_unidade_tipo = <%=cd_unidade_tipo%><br>
dt_inicio_contrato = <%=dt_inicio_contrato%><br>
dt_fim_contrato = <%=dt_fim_contrato%><br>
nm_cor = <%=nm_cor%><br>

action = <%=action%><br>

asd

</body>
</html>
