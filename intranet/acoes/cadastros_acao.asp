
  <!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
   <!--include file="includes/inc_area_restrita.asp"-->

<%

acao = request("acao")
campo = request("campo")

modo = request("modo")
janela = request("janela")
metodo = request("metodo")

'Avaliações
nm_avaliacao = request("avaliacao")

'Fornecedores
nm_fornecedor = request("nm_fornecedor")
nm_contato = request("nm_contato")
nm_contato2 = request("nm_contato2")
nm_endereco = request("nm_endereco")
nm_numero = request("nm_numero")
nm_bairro = request("nm_bairro")
nm_cidade = request("nm_cidade")
nm_estado = request("nm_estado")

cd_prazo = request("cd_prazo")

'Equipamentos
cd_pn = request("cd_pn")
nm_equipamento = request("nm_equipamento")
cd_equipamento = request("cd_equipamento")
cd_tipo = request("cd_tipo")
cd_ns = request("cd_ns")

'Especialidades
cd_especialidade = request("cd_especialidade")
nm_especialidade = request("nm_especialidade")
cd_status = request("cd_status")

'Marcas
cd_marca = request("cd_marca")
nm_marca = request("nm_marca")

'Patrimonio
cd_patrimonio = request("cd_patrimonio")
dt_descarte = request("dt_descarte")
nm_motivo = request("nm_motivo")

'Unidades
cd_unidade = request("cd_unidade")
nm_sigla = request("nm_sigla")
nm_unidade = request("nm_unidade")

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

'Aparelhos
cd_aparelho = request("cd_aparelho")
nm_modelo = request("nm_modelo")
nm_apelido = request("nm_apelido")

'Data
dt_datacad = request("dt_datacad")
dt_datatu = request("dt_datatu")
dt_data = mes&"/"&dia&"/"&ano

cd_status = request("cd_status")

'Data Atual

data_atual = Month(now)&"/"&Day(now)&"/"&Year(now)&" "&Hour(now)&":"&Minute(now)&":"&Second(now)
cd_user = session("cd_codigo")




' ****  AÇÕES  *******************************************************************************

'*********************
'***  Fornecedores ***
'*********************
	If acao = "inserir" AND modo = "forn" Then
	'Insere dados
	'xsql = "up_fornecedor_insert @nm_fornecedor='"&nm_fornecedor&"', @nm_contato='"&nm_contato&"', @nm_contato2='"&nm_contato2&"', @nm_endereco='"&nm_endereco&"', @nm_telefone_cod1='"&nm_telefone_cod1&"', @nm_telefone1='"&nm_telefone1&"', @nm_telefone_cod2='"&nm_telefone_cod2&"', @nm_telefone2='"&nm_telefone2&"'"
	xsql = "up_fornecedor_insert @nm_fornecedor='"&nm_fornecedor&"', @nm_contato='"&nm_contato&"', @nm_contato2='"&nm_contato2&"', @nm_cep='', @nm_cnpj='',@nm_ie='', @nm_endereco='"&nm_endereco&"', @nm_numero = '"&nm_numero&"', @nm_bairro='"&nm_bairro&"', @nm_cidade='"&nm_cidade&"',@nm_estado='"&nm_estado&"', @cd_prazo='"&cd_prazo&"'"
	action = "Fornecedor inserido com sucesso"
	dbconn.execute(xsql)
	redirecionamento = "../fornecedores.asp?action="&action&""
	
	ElseIf acao = "editar" AND modo = "forn" Then
	'Modifica os dados
	xsql = "up_fornecedor_update @cd_codigo='"&cd_codigo&"', @nm_fornecedor='"&nm_fornecedor&"', @nm_contato='"&nm_contato&"', @nm_contato2='"&nm_contato2&"', @nm_endereco='"&nm_endereco&"', @nm_numero = '"&nm_numero&"', @nm_bairro='"&nm_bairro&"', @nm_cidade='"&nm_cidade&"',@nm_estado='"&nm_estado&"', @cd_prazo='"&cd_prazo&"'"
	action = "Fornecedor editado com sucesso"
	dbconn.execute(xsql)
	'response.redirect("../fornecedores.asp?action="&action&"")
	redirecionamento = "../fornecedores.asp?action="&action&""
	
	ElseIf acao = "excluir" AND modo = "forn" Then
	xsql = "up_fornecedor_delete @cd_codigo='"&cd_codigo&"'"
	dbconn.execute(xsql)
	action = "Fornecedor excluído com sucesso"
	'response.redirect("../fornecedores.asp?action="&action&"")
	redirecionamento = "../fornecedores.asp?action="&action&""


'********************
'*** Equipamentos ***
'********************
'	ElseIf acao = "inserir" AND modo = "equip" Then
'	'Insere dados
'	xsql = "up_equipamento_insert @cd_pn='"&cd_pn&"', @nm_equipamento='"&nm_equipamento&"', @cd_tipo = '"&cd_tipo&"', @cd_user='"&cd_user&"',@data_atual='"&data_atual&"'"
'	action = "Equipamemto inserido com sucesso"
'	dbconn.execute(xsql)
'		if janela = "" Then
'		response.redirect("../equipamentos.asp?action="&action&"")
'		end if
'	
'	ElseIf acao = "editar" AND modo = "equip" Then
'	'Modifica os dados
'	xsql = "up_manutencao_update @cd_pn='"&cd_pn&"', @nm_equipamento='"&nm_equipamento&"', @cd_tipo = '"&cd_tipo&"', @cd_user='"&cd_user&"',@data_atual='"&data_atual&"'"
'	action = "editado com sucesso"
'	dbconn.execute(xsql)
'	response.redirect("equipamento.asp?action="&action&"")


'******************
'*** Patrimonio ***
'******************
	ElseIf acao = "inserir" AND modo = "patrimonio" Then
	'Insere dados
	xsql = "up_equipamento_insert @cd_pn='"&cd_pn&"', @nm_equipamento='"&nm_equipamento&"'"
	action = "Equipamemto inserido com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		response.redirect("../equipamentos.asp?action="&action&"")
		end if
	
	ElseIf acao = "editar" AND modo = "patrimio" Then
	'Modifica os dados
	xsql = "up_manutencao_update @cd_pn='"&cd_pn&"', @nm_equipamento='"&nm_equipamento&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	response.redirect("equipamento.asp?action="&action&"")


'**************
'*** Marcas ***
'**************
	Elseif acao = "inserir" AND modo = "marca" Then
	'Insere dados
	xsql = "up_marca_insert @nm_marca='"&nm_marca&"'"
	action = "Especialidade inserida com sucesso"
	dbconn.execute(xsql)
	redirecionamento = "../marcas.asp?action="&action&""
	
	
	ElseIf acao = "editar" AND modo = "marca" Then
	'Modifica os dados
	xsql = "up_marca_update @nm_marca='"&nm_marca&"'"
	action = "editada com sucesso"
	dbconn.execute(xsql)
	redirecionamento = "../marcas.asp?action="&action&""


'****************
'*** Unidades ***
'****************
	Elseif acao = "inserir" AND modo = "unidade" Then
	'Insere dados
	xsql = "up_unidade_insert @nm_sigla='"&nm_sigla&"', @nm_unidade='"&nm_unidade&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "unidade inserida com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		'response.redirect("../unidades.asp?action="&action&"")
		response.redirect("../ferramentas.asp?tipo=unidade&status=1")
		end if
	
	Elseif acao = "editar" AND modo = "unidade" then
	'Altera dados
		If status = 1 then
			xsql = "up_unidade_update @cd_unidade = '"&cd_unidade&"', @nm_sigla='"&nm_sigla&"', @nm_unidade='"&nm_unidade&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Unidade inserida com sucesso"
			dbconn.execute(xsql)
			response.write("editada!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=unidade&status=1")
				end if
		
		elseif status = 0 then
			xsql = "up_unidade_inativar @cd_unidade='"&cd_unidade&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "A unidade foi reativada"
			response.write("Inativa!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=unidade&status=1")
				end if
		end if
	
	Elseif acao = "excluir" AND modo = "unidade" then
	'Exclui dados
		if status = 2 then
			xsql = "up_unidade_delete @cd_unidade='"&cd_unidade&"'"
			action = "Unidade exluído com sucesso"
			response.write("apagada!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=unidade&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_unidade_inativar @cd_unidade='"&cd_unidade&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "A unidade está inativa"
			response.write("Inativa!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=unidade&status=1")
				end if
		end if

'*****************
'*** Situações ***
'*****************
	Elseif acao = "inserir" AND modo = "situacao" Then
	'Insere dados
	xsql = "up_situacao_insert @nm_situacao='"&nm_situacao&"'"
	action = "Situacao inserida com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		response.redirect("../unidades.asp?action="&action&"")
		end if
	
	
'***************************************************************************************************
	''Especialidades
	'Elseif acao = "inserir" AND modo = "especialidade" Then
	''Insere dados
	'xsql = "up_especialidade_insert @nm_especialidade='"&nm_especialidade&"'"
	'action = "Especialidade inserida com sucesso"
	'dbconn.execute(xsql)
	'response.redirect("especialidade.asp?action="&action&"")
	
	'ElseIf acao = "editar" AND modo = "especialidade" Then
	''Modifica os dados
	'xsql = "up_espec_update @nm_especialidade='"&nm_especialidade&"'"
	'action = "editada com sucesso"
	'dbconn.execute(xsql)
	'response.redirect("especialidade.asp?action="&action&"")

'*****************************************************************************************************
	'Cargo
	Elseif acao = "inserir" AND modo = "cargo" Then
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
			xsql = "up_cargo_salario_insert @cd_cargo = '"&nova_funcao&"', @cd_valor = '"&cd_salario&"', @cd_modo = 1, @dt_data = '"&dt_data&"'"
			dbconn.execute(xsql)
			response.redirect("../adm_cargos.asp")
	
	Elseif acao = "editar" AND modo = "cargo" Then
	'Atualiza dados (insere valor e data novos)
	'xsql = "up_cargo_valorupdate @nm_cargo='"&nm_cargo&"', @cd_especificidade='"&cd_especificidade&"'"
	'action = "cargo inserido com sucesso"
	'dbconn.execute(xsql)
	response.write("insere valor do salário")
	xsql = "up_cargo_salario_insert @cd_cargo = '"&cd_cargo&"', @cd_valor = '"&cd_salario&"', @cd_modo = 1, @dt_data = '"&dt_data&"'"
	dbconn.execute(xsql)
			
	'response.write("ok, mano!")
	response.redirect("../adm_cargos.asp")
	
	'Ajuste
	Elseif acao = "editar" AND modo = "ajuste" Then
	xsql = "up_relatorio_cooperativa_ajuste @cd_codigo='"&cd_codigo&"', @cd_funcao='"&cd_funcao&"', @nm_status='"&nm_status&"', @cd_unidade='"&cd_unidade&"', @mes='"&mes&"', @ano='"&ano&"'"
	dbconn.execute(xsql)
	'response.redirect("../produtividade/unidades.asp?action="&action&"")
	'response.write("Ajuste ok")
	
	'*****************
	'*** Convenios ***
	'*****************
	Elseif acao = "inserir" AND modo = "convenio" Then
	
	'Insere dados
	xsql = "up_convenios_insert @nm_convenio='"&nm_convenio&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "Convenio inserido com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		response.redirect("../ferramentas.asp?tipo=convenio&status=1")
		end if
	
	Elseif acao = "editar" AND modo = "convenio" then
	'Altera dados
		If status = 1 then
			xsql = "up_convenios_update @cd_convenio='"&cd_convenio&"', @nm_convenio='"&nm_convenio&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Convenio inserido com sucesso"
			dbconn.execute(xsql)
			response.write("editado!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=convenio&status=1")
				end if
		
		elseif status = 0 then
			xsql = "up_convenios_inativar @cd_convenio='"&cd_convenio&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O cadastro do convenio foi reativado"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=convenio&status=1")
				end if
		end if
		
	Elseif acao = "excluir" AND modo = "convenio" then
	'Exclui dados
		if status = 2 then
			xsql = "up_convenios_delete @cd_convenio='"&cd_convenio&"'"
			action = "Convenio exluído com sucesso"
			response.write("apagado!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=convenio&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_convenios_inativar @cd_convenio='"&cd_convenio&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Convenio está inativo"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=convenio&status=1")
				end if
		end if

	'***************
	'*** Medicos ***
	'***************
	Elseif acao = "inserir" AND modo = "medico" then
	'Insere dados
	xsql = "up_medicos_insert @cd_crm='"&cd_crm&"', @nm_medico='"&nm_medico&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "Medico inserido com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		response.redirect("../ferramentas.asp?tipo=medico&status=1")
		end if
	
	Elseif acao = "editar" AND modo = "medico" then
	'Altera dados
		If status = 1 then
			xsql = "up_medicos_update @cd_medico='"&cd_medico&"', @nm_medico='"&nm_medico&"', @cd_crm='"&cd_crm&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Medico inserido com sucesso"
			dbconn.execute(xsql)
			response.write("editado!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=medico&status=1")
				end if
		
		elseif status = 0 then
			xsql = "up_medicos_inativar @cd_medico='"&cd_medico&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O cadastro do medico foi reativado"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=medico&status=1")
				end if
		end if
	
	Elseif acao = "excluir" AND modo = "medico" then
	'Exclui dados
		if status = 2 then
			xsql = "up_medicos_delete @cd_medico='"&cd_medico&"'"
			action = "Medico exluído com sucesso"
			response.write("apagado!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=medico&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_medicos_inativar @cd_medico='"&cd_medico&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Medico está inativo"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=medico&status=1")
				end if
		end if

	'************
	'*** Rack ***
	'************
	Elseif acao = "inserir" AND modo = "rack" then
	'Insere dados
		xsql = "up_rack_insert @nm_rack='"&nm_rack&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
		action = "Rack inserido com sucesso"
		dbconn.execute(xsql)
			if janela = "" Then
			response.redirect("../ferramentas.asp?tipo=rack&status=1")
			end if
		
	Elseif acao = "editar" AND modo = "rack" then
	'Altera dados
		If status = 1 then
			xsql = "up_rack_update @cd_rack='"&cd_rack&"', @nm_rack='"&nm_rack&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Rack inserido com sucesso"
			dbconn.execute(xsql)
			response.write("editado!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=rack&status=1")
				end if
		
		elseif status = 0 then
			xsql = "up_rack_inativar @cd_rack='"&cd_rack&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O cadastro do rack foi reativado"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=rack&status=1")
				end if
		end if
		
	Elseif acao = "excluir" AND modo = "rack" then
	'Exclui dados
		if status = 2 then
			xsql = "up_rack_delete @cd_rack='"&cd_rack&"'"
			action = "Rack exluído com sucesso"
			response.write("apagado!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=rack&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_rack_inativar @cd_rack='"&cd_rack&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Rack está inativo"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=rack&status=1")
				end if
		end if

	'**********************************
	'*** Especialidade - Protocolos ***
	'**********************************
	Elseif acao = "inserir" AND modo = "espec" then
	'Insere dados
	xsql = "up_espec_insert @nm_especialidade='"&nm_especialidade&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "Especialidade inserida com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		'response.redirect("../especialidade.asp?action="&action&"")
		response.redirect("../ferramentas.asp?tipo=espec&status=1")
		end if
		
	Elseif acao = "editar" AND modo = "espec" then
	'Altera dados
		If status = 1 then
			xsql = "up_espec_update @cd_especialidade='"&cd_especialidade&"', @nm_especialidade='"&nm_especialidade&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Especialidade inserida com sucesso"
			dbconn.execute(xsql)
			response.write("editada!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=espec&status=1")
				end if
				
		elseif status = 0 then
			xsql = "up_espec_inativar @cd_especialidade='"&cd_especialidade&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "A especialidade foi reativada"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=espec&status=1")
				end if
		end if
		
	Elseif acao = "excluir" AND modo = "espec" then
	'Exclui dados
		if status = 2 then
			xsql = "up_espec_delete @cd_especialidade='"&cd_especialidade&"'"
			action = "Especialidade exluída com sucesso"
			response.write("apagada!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=espec&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_espec_inativar @cd_especialidade='"&cd_especialidade&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Especialidade está inativa"
			response.write("Inativa!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=espec&status=1")
				end if
		end if

	'**********************************
	'*** Especialidade - Manutenção ***
	'**********************************
	Elseif acao = "inserir" AND modo = "especialidade" then
	'Insere dados
	xsql = "up_especialidade_insert @nm_especialidade='"&nm_especialidade&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "Especialidade inserida com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		'response.redirect("../especialidade.asp?action="&action&"")
		response.redirect("../ferramentas.asp?tipo=especialidade&status=1")
		end if
		
	Elseif acao = "editar" AND modo = "especialidade" then
	'Altera dados
		If status = 1 then
			xsql = "up_especialidade_update @cd_especialidade='"&cd_especialidade&"', @nm_especialidade='"&nm_especialidade&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Especialidade inserida com sucesso"
			dbconn.execute(xsql)
			response.write("editada!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=especialidade&status=1")
				end if
				
		elseif status = 0 then
			xsql = "up_especialidade_inativar @cd_especialidade='"&cd_especialidade&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "A especialidade foi reativada"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=especialidade&status=1")
				end if
		end if
		
	Elseif acao = "excluir" AND modo = "especialidade" then
	'Exclui dados
		if status = 2 then
			xsql = "up_especialidade_delete @cd_especialidade='"&cd_especialidade&"'"
			action = "Especialidade exluída com sucesso"
			response.write("apagada!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=especialidade&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_especialidade_inativar @cd_especialidade='"&cd_especialidade&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Especialidade está inativa"
			response.write("Inativa!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=especialidade&status=1")
				end if
		end if

	'********************
	'*** Procedimento ***
	'********************
	Elseif acao = "inserir" AND modo = "procedimento" then
	'Insere dados
	xsql = "up_procedimento_insert @cod_amb='"&cod_amb&"', @nm_procedimento='"&nm_procedimento&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "Procedimento inserida com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		'response.redirect("../procedimento.asp?action="&action&"")
		response.redirect("../ferramentas.asp?tipo=procedimento&status=1")
		end if
		
	Elseif acao = "editar" AND modo = "procedimento" then
	'Altera dados
		If status = 1 then
			xsql = "up_procedimento_update @cd_procedimento = '"&cd_procedimento&"', @cod_amb='"&cod_amb&"', @nm_procedimento='"&nm_procedimento&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Procedimento inserido com sucesso"
			dbconn.execute(xsql)
			response.write("editado!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=procedimento&status=1")
				end if
		
		elseif status = 0 then
			xsql = "up_procedimento_inativar @cod_amb='"&cod_amb&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O procedimento foi reativado"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=procedimento&status=1")
				end if
		end if
	
	Elseif acao = "excluir" AND modo = "procedimento" then
	'Exclui dados
		if status = 2 then
			xsql = "up_procedimento_delete @cod_amb='"&cod_amb&"'"
			action = "Procedimento exluído com sucesso"
			response.write("apagado!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=procedimento&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_procedimento_inativar @cod_amb='"&cod_amb&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O procedimento está inativo"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=procedimento&status=1")
				end if
		end if
		
		
	
	'****************
	'*** Material ***
	'****************
	Elseif acao = "inserir" AND modo = "material" then
	'Insere dados
	xsql = "up_material_insert @nm_material='"&nm_material&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "Material inserida com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		'response.redirect("../material.asp?action="&action&"")
		response.redirect("../ferramentas.asp?tipo=material&status=1")
		end if
		
	Elseif acao = "editar" AND modo = "material" then
	'Altera dados
		If status = 1 then
			xsql = "up_materiais_update @cd_material='"&cd_material&"', @nm_material='"&nm_material&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Material inserido com sucesso"
			dbconn.execute(xsql)
			response.write("editado!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=material&status=1")
				end if
				
		elseif status = 0 then
			xsql = "up_materiais_inativar @cd_material='"&cd_material&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O material foi reativado"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=material&status=1")
				end if
		end if
		
	Elseif acao = "excluir" AND modo = "material" then
	'Exclui dados
		if status = 2 then
			xsql = "up_materiais_delete @cd_material='"&cd_material&"'"
			action = "Material exluído com sucesso"
			response.write("apagado!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=material&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_materiais_inativar @cd_material='"&cd_material&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O material está inativo"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=material&status=1")
				end if
		end if
	
	'*******************
	'*** Equipamento ***
	'*******************
	Elseif acao = "inserir" AND modo = "equipamento" then
	'Insere dados
	xsql = "up_equipamento_insert @nm_equipamento='"&nm_equipamento&"', @cd_pn='"&cd_pn&"',@cd_tipo= '"&cd_tipo&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "Equipamento inserido com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		response.redirect("../ferramentas.asp?tipo=equipamento&status=1")
		end if
		
	Elseif acao = "editar" AND modo = "equipamento" then
	'Altera dados
		If status = 1 then
			xsql = "up_equipamento_update @cd_equipamento='"&cd_equipamento&"', @nm_equipamento='"&nm_equipamento&"', @cd_pn='"&cd_pn&"', @cd_tipo='"&cd_tipo&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Equipamento inserido com sucesso"
			dbconn.execute(xsql)
			response.write("editado!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=equipamento&status=1")
				end if
				
		elseif status = 0 then
			xsql = "up_equipamento_inativar @cd_equipamento='"&cd_equipamento&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O equipamento foi reativado"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
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
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=equipamento&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_equipamento_inativar @cd_equipamento='"&cd_equipamento&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O equipamento está inativo"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=equipamento&status=1")
				end if
		end if
	
	
	
	'*****************
	'*** Aparelhos ***
	'*****************
	Elseif acao = "inserir" AND modo = "aparelho" then
	'Insere dados
	xsql = "up_aparelho_insert @nr_numero='"&nr_numero&"', @data_atual='"&data_atual&"', @cd_user='"&cd_user&"'"
	action = "Aparelho inserido com sucesso"
	dbconn.execute(xsql)
		if janela = "" Then
		'response.redirect("../material.asp?action="&action&"")
		response.redirect("../ferramentas.asp?tipo=aparelho&status=1")
		end if
		
	Elseif acao = "editar" AND modo = "aparelho" then
	'Altera dados
		If status = 1 then
			xsql = "up_aparelho_update @cd_aparelho='"&cd_aparelho&"', @nr_numero='"&nr_numero&"', @nr_imei_1='"&nr_imei_1&"', @nr_imei_2='"&nr_imei_2&"', @cd_modelo='"&cd_modelo&"', @cd_status='"&cd_status&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "Aparelho inserido com sucesso"
			dbconn.execute(xsql)
			response.write("editado!!!")
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=aparelho&status=1")
				end if
				
		elseif status = 0 then
			xsql = "up_aparelhos_inativar @cd_aparelho='"&cd_aparelho&"', @cd_status=1, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O aparelho foi reativado"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=material&status=1")
				end if
		end if
		
	Elseif acao = "excluir" AND modo = "aparelho" then
	'Exclui dados
		if status = 2 then
			xsql = "up_aparelhos_delete @cd_aparelho='"&cd_aparelho&"'"
			action = "Material exluído com sucesso"
			response.write("apagado!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=aparelho&status=0")
				end if
			
		Elseif status = 0 then
			xsql = "up_aparelhos_inativar @cd_aparelho='"&cd_aparelho&"', @cd_status=0, @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
			action = "O aparelho está inativo"
			response.write("Inativo!!!")
			dbconn.execute(xsql)
				if janela = "" Then
				response.redirect("../ferramentas.asp?tipo=aparelho&status=1")
				end if
		end if
	
	
	
	
	
	
	Else 
	action = "erro"
	End If
	
	if janela = "" Then
	'response.redirect(redirecionamento)
	end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>cadastros</title>
</head><%'metodo = ""%>

<%if metodo = "fecha_atualiza" Then%>
<body onLoad="fecha_atualiza()">
<%elseif metodo = "fecha" Then%>
<body onLoad="fecha()">
<%else%>
<body>
<%end if%>
<br>
Mensagem = <%=action%><br>
Ação: <%=acao%><br>
modo: <%=modo%><br>
janela: <%=janela%><br>
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
fornecedor = <%=nm_fornecedor%><br>
nm_contato = <%=nm_contato%><br>
nm_contato2 = <%=nm_contato2%><br>
nm_endereco = <%=nm_endereco%><br>
nm_numero = <%=nm_numero%><br>
nm_bairro = <%=nm_bairro%><br>
nm_cidade = <%=nm_cidade%><br>
nm_estado = <%=nm_estado%><br><br>

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
User = <%=cd_user%><br>




</body>
</html>
