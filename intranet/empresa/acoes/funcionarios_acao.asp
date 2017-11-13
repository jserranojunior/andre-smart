
<%@ Language=VBScript %>
<%

'Set Upload = Server.CreateObject("Persits.Upload.1")

'Count = Upload.Save("E:\users\vdlap.com.br\website\foto\temp\")
%>

<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%cd_user = session("cd_codigo")'request("cd_user")
strcod    = request("cod")
jan = request("jan")
func_ativos = request("func_ativos")
cod_cont  = request("cod_cont")
cod_dep   = request("cod_dep")
cod_esc   = request("cod_esc")
cod_emp   = request("cod_emp")
cod_cargo = request("cod_cargo")
cod_funcao = request("cod_funcao")
cod_contrato = request("cod_contrato")
cod_unidade = request("cod_unidade")
	if cod_unidade = "" Then
		cod_unidade = "27"
	end if
cod_horario = request("cod_horario")
cod_valor = request("cod_valor")

session_cd_usuario = session("cd_codigo")

stracao   = request("acao")
protege_exclusao = request("protege_exclusao")

strcd_funcionario = request("cd_funcionario")
strcd_sexo = request("cd_sexo")
strregime_trabalho = request("cd_regime_trabalho")
cd_funcionario_contrato = request("cd_funcionario_contrato")
strcd_matricula = request("cd_matricula")
strnome =  request("nm_nome")
strnm_sobrenome =  request("nm_sobrenome")
strfoto = request("foto_h")
'*************************************************
dia_nascimento = request("dia_nascimento")
mes_nascimento = request("mes_nascimento")
ano_nascimento = request("ano_nascimento")
strdt_nascimento = mes_nascimento&"/"&dia_nascimento&"/"&ano_nascimento
	if dia_nascimento <> "" AND mes_nascimento <> "" AND ano_nascimento <> "" Then
		strdt_nascimento = "'"&month(strdt_nascimento)&"/"&day(strdt_nascimento)&"/"&year(strdt_nascimento)&"'"
	else
		strdt_nascimento = "NULL"
	end if
'*************************************************

'Strdt_demissao =  request("dt_demissao")
strnm_email =  request("nm_email")
strnm_rg =  request("nm_rg")
	nm_rg_expedidor = request("nm_rg_expedidor")
	if nm_rg_expedidor = "" Then 
		nm_rg_expedidor = "NULL"
	else
		nm_rg_expedidor = "'"&nm_rg_expedidor&"'"
	end if


strnm_cpf =  request("nm_cpf")
strmae = request("nm_mae")
strpai = request("nm_pai")
strnm_ddd = request("nm_ddd")
strnm_fone =  request("nm_fone")
strnm_ddd_cel = request("nm_ddd_cel")
strnm_cel =  request("nm_cel")
strocontato = request("nm_o_contato")
strnm_endereco = request("nm_endereco")
strnr_numero =  request("nr_numero")
strnm_complemento = request("nm_complemento")
strnm_cidade =  request("nm_cidade")
strnm_estado =  request("nm_estado")
strnm_cep = request("nm_cep")
strbairro = request("nm_bairro")
str_nacionalidade = request("nm_nacionalidade")
str_escolaridade = request("nm_escolaridade")
	str_escolaridade = replace(str_escolaridade,"'","&quot;")
str_pis = request("cd_pis")
str_ctps = request("cd_ctps")
str_ctps_serie = request("cd_ctps_serie")
strqtd_dependentes = request("cd_qtd_dependentes")
strqtd_filhos = request("cd_qtd_filhos")

strcd_ferias = request("cd_ferias")

'str_estado_civil 	= request("cd_estado_civil")
'str_cd_estado_civil = request("cd_estado_civil")
cd_estado_civil 	= request("cd_estado_civil")
nm_conjuge = request("nm_conjuge")
'dt_conjuge_nasc = request("dt_conjuge_nasc")
dt_conjuge_nasc_dia = request("dt_conjuge_nasc_dia")
dt_conjuge_nasc_mes = request("dt_conjuge_nasc_mes")
dt_conjuge_nasc_ano = request("dt_conjuge_nasc_ano")

	if dt_conjuge_nasc_dia <> "" AND dt_conjuge_nasc_mes <> "" AND dt_conjuge_nasc_ano <> "" Then
		'dt_conjuge_nasc = "'"&month(dt_conjuge_nasc)&"/"&day(dt_conjuge_nasc)&"/"&year(dt_conjuge_nasc)&"'"
		dt_conjuge_nasc = "'"&dt_conjuge_nasc_mes&"/"&dt_conjuge_nasc_dia&"/"&dt_conjuge_nasc_ano&"'"
	else
		dt_conjuge_nasc = "NULL"
	end if
cd_plano_saude_conj = request("cd_plano_saude_conj")
	if cd_plano_saude_conj = "" Then
		cd_plano_saude_conj = null
	end if

str_carga_horaria = request("cd_carga_horaria")
str_carga_diaria = request("cd_carga_diaria")
str_assistencia_medica = request("cd_assistencia_medica")
str_assistencia_medica_matricula = request("cd_assistencia_medica_matricula")
str_assistencia_medica_coparticipacao = request("nr_assistencia_medica_coparticipacao")
if str_assistencia_medica_coparticipacao = "" Then
		str_assistencia_medica_coparticipacao = "NULL"
	end if

str_assistencia_odonto = request("cd_assistencia_odonto")
str_assistencia_odonto_matricula = request("cd_assistencia_odonto_matricula")
str_assistencia_odonto_coparticipacao = request("nr_assistencia_odonto_coparticipacao")
	if str_assistencia_odonto_coparticipacao = "" Then
		str_assistencia_odonto_coparticipacao = "NULL"
	end if
str_vr = request("cd_vr")
str_vale_alimentacao = request("cd_vale_alimentacao")
str_cmt_bom = request("cd_cmt_bom")
str_sptrans = request("cd_sptrans")

str_cargo = request("cd_cargo")
str_funcao = request("cd_funcao")
str_status = request("cd_status")
strstatus =  request("cd_status")
str_aparelhos = request("cd_aparelhos")

strcd_unidade = request("cd_unidade")
strcd_empresa = request("cd_empresa")

strnm_intervalo = request("nm_intervalo")
strcd_suspenso = request("cd_suspenso")
str_tbl = request("tbl")

'*** Admissão ***
cd_dia = request("cd_dia")
cd_mes = request("cd_mes")
cd_ano = request("cd_ano")
	cd_hora = hour(now)
	cd_minuto = minute(now)
	cd_segundo = second(now)

'*** Demissão ***
cd_dia_d = request("cd_dia_d")
cd_mes_d = request("cd_mes_d")
cd_ano_d = request("cd_ano_d")
	cd_hora_d = hour(now)
	cd_minuto_d = minute(now)
	cd_segundo_d = second(now)
		if cd_dia_d <> "" and cd_mes_d <> "" and cd_ano_d <> "" Then
				Strdt_desligado = "'"&cd_mes_d&"/"&cd_dia_d&"/"&cd_ano_d&"'"
			else
				Strdt_desligado = "NULL"
			end if


		if stracao = "recisao" then ' *** RECISÃO ***
			if cd_dia <> "" and cd_mes <> "" and cd_ano <> "" Then
				Strdt_recisao = "'"&cd_mes&"/"&cd_dia&"/"&cd_ano&"'"
			else
				Strdt_recisao = "NULL"
			end if
			
		else ' *** ADMISSÃO ***
			if cd_dia <> "" and cd_mes <> "" and cd_ano <> "" Then
				'Strdt_contratado = "'"&cd_mes&"/"&cd_dia&"/"&cd_ano&" 00:00'"
				Strdt_contratado = "'"&cd_mes&"/"&cd_dia&"/"&cd_ano&"'"
			else
				Strdt_contratado = "NULL"
			end if
		end if
	


dt_atualizacao = request("dt_atualizacao")
	if dt_atualizacao = "" then
		dt_atualizacao = month(now)&"/"&day(now)&"/"&year(now)
	end if
	
'dt_atualizacao_manual = request("cd_dia")&"/"&request("cd_mes")&"/"&request("cd_ano")
dt_atualizacao_manual = request("cd_ano")&"-"&request("cd_mes")&"-"&request("cd_dia")


'strhr_entrada = dt_atualizacao&" "&request("hr_entrada")&":"&request("min_entrada")
strhr_entrada = request("hr_entrada")&":"&request("min_entrada")
strhr_saida = request("hr_saida")&":"&request("min_saida")

if len(strhr_entrada) > 3 OR len(strhr_saida) > 3 Then
     
     strhr_entrada = dt_atualizacao_manual&" "&request("hr_entrada")&":"&request("min_entrada")

	'Verifica se o horário é noturno e altera o dia
	if int(hour(strhr_entrada)) > int(hour(strhr_saida)) then
		dia_atualizacao = day(dt_atualizacao_manual) + 1
		mes_atualizacao = month(dt_atualizacao_manual)
		ano_atualizacao = year(dt_atualizacao_manual)
			if int(dia_atualizacao) > int(ultimodiames(mes_atualizacao,ano_atualizacao)) then
				dia_atualizacao = 1
				mes_atualizacao = mes_atualizacao + 1
				ano_atualizacao = ano_atualizacao
					if mes_atualizacao > 12 then
						mes_atualizacao = 1
						ano_atualizacao = ano_atualizacao + 1
					end if
			end if
			'strhr_saida = dia_atualizacao&"/"&mes_atualizacao&"/"&ano_atualizacao&" "&strhr_saida
			'strhr_saida = dia_atualizacao&"/"&mes_atualizacao&"/"&ano_atualizacao&" "&strhr_saida
			strhr_saida = ano_atualizacao&"-"&mes_atualizacao&"-"&dia_atualizacao&" "&strhr_saida
		
	else
		strhr_saida = dt_atualizacao_manual&" "&request("hr_saida")&":"&request("min_saida")
	end if

end if


str_banco = request("nm_banco")
str_banco_ag = request("cd_banco_ag")
str_banco_conta = request("cd_banco_conta")
str_banco_conta_tipo = request("cd_banco_conta_tipo")
nm_tit_eleitor = request("nm_tit_eleitor")
nm_reservista = request("nm_reservista")
nm_creche = request("nm_creche")

str_contatos_total = request("contatos_total")
str_dependentes_total = request("dependentes_total")
str_escolaridade_total = request("escolaridade_total")
str_emprego_total = request("emprego_total")
str_valores_total = request("valores_total")

strcd_valor = request("cd_valor")
str_valor_tipo = request("cd_padrao_valor_tipo")
str_valor = request("nr_padrao_valor")
str_valor_atualizacao = request("dt_padrao_atualizacao")
str_valor_obs = request("nm_padrao_obs")


nm_naturalidade = request("nm_naturalidade")
dt_rg_emissao_dia = request("dt_rg_emissao_dia")
dt_rg_emissao_mes = request("dt_rg_emissao_mes")
dt_rg_emissao_ano = request("dt_rg_emissao_ano")
	
	if dt_rg_emissao_dia <> "" AND dt_rg_emissao_mes <> "" AND dt_rg_emissao_ano <> "" Then
		dt_rg_emissao = "'"&dt_rg_emissao_mes&"/"&dt_rg_emissao_dia&"/"&dt_rg_emissao_ano&"'"
	else
		dt_rg_emissao = "NULL"
	end if

nr_tit_eleitor_zona = request("nr_tit_eleitor_zona")
nr_tit_eleitor_seccao = request("nr_tit_eleitor_seccao")
nm_reservista_cat = request("nm_reservista_cat")
nr_hepatite_b1 = request("nr_hepatite_b1")
nr_hepatite_b2 = request("nr_hepatite_b2")
nr_hepatite_b3 = request("nr_hepatite_b3")
	if nr_hepatite_b3 <> "" Then 
		nr_hepatite_b = nr_hepatite_b3
	elseif	nr_hepatite_b2 <> "" then
		nr_hepatite_b = nr_hepatite_b2
	elseif nr_hepatite_b1 <> "" then
		nr_hepatite_b = nr_hepatite_b1
	end if
nr_dupla_adulto1 = request("nr_dupla_adulto1")
nr_dupla_adulto2 = request("nr_dupla_adulto2")
nr_dupla_adulto3 = request("nr_dupla_adulto3")
	if nr_dupla_adulto3 <> "" Then 
		nr_dupla_adulto = nr_dupla_adulto3
	elseif	nr_dupla_adulto2 <> "" then
		nr_dupla_adulto = nr_dupla_adulto2
	elseif nr_dupla_adulto1 <> "" then
		nr_dupla_adulto = nr_dupla_adulto1
	end if
	
dt_dupla_adulto_validade_dia = request("dt_dupla_adulto_validade_dia")
dt_dupla_adulto_validade_mes = request("dt_dupla_adulto_validade_mes")
dt_dupla_adulto_validade_ano = request("dt_dupla_adulto_validade_ano")
	if dt_dupla_adulto_validade_dia <> "" AND dt_dupla_adulto_validade_mes <> "" AND dt_dupla_adulto_validade_ano <> "" Then
		dt_dupla_adulto_validade = "'"&dt_dupla_adulto_validade_mes&"/"&dt_dupla_adulto_validade_dia&"/"&dt_dupla_adulto_validade_ano&"'"		
	else
		dt_dupla_adulto_validade = "NULL"
	end if
nr_scr = request("nr_scr")
dt_exame_medico_dia = request("dt_exame_medico_dia")
dt_exame_medico_mes = request("dt_exame_medico_mes")
dt_exame_medico_ano = request("dt_exame_medico_ano")
	if dt_exame_medico_dia <> "" AND dt_exame_medico_mes <> "" AND  dt_exame_medico_ano <> "" Then
		dt_exame_medico = "'"&dt_exame_medico_mes&"/"&dt_exame_medico_dia&"/"&dt_exame_medico_ano&"'"		
	else
		dt_exame_medico = "NULL"
	end if
nr_caract_altura = request("nr_caract_altura")
nr_caract_peso = request("nr_caract_peso")
nr_caract_manequim = request("nr_caract_manequim")
nr_caract_calcado = request("nr_caract_calcado")
nm_parente = request("nm_parente")

cd_parente_parentesco = request("cd_parente_parentesco")
cd_residencia_tipo = request("cd_residencia_tipo")
	'if 
cd_residencia_financ = request("cd_residencia_financ")
	if cd_residencia_tipo = 2 then
		cd_residencia_financ = 0
	end if
cd_residencia_reside = request("cd_residencia_reside")
cd_veiculo_tipo = request("cd_veiculo_tipo")
cd_veiculo_financ = request("cd_veiculo_financ")

nm_obs = request("nm_obs")

str_conspro = request("cd_conspro")
str_rgpro_cargo = request("cd_rgpro_cargo")
str_numreg = request("cd_numreg")
str_rgpro_concessao = request("rgpro_concessao")
str_rgpro_status = request("rgpro_status")
str_rgpro = request("obs_rgpro")

str_dt_rgproinscr_dia = request("dt_rgproinscr_dia")
str_dt_rgproinscr_mes = request("dt_rgproinscr_mes")
str_dt_rgproinscr_ano = request("dt_rgproinscr_ano")
	if str_dt_rgproinscr_dia <> "" AND str_dt_rgproinscr_mes <> "" AND str_dt_rgproinscr_ano <> "" Then
		str_dt_rgproinscr = "'"&str_dt_rgproinscr_mes&"/"&str_dt_rgproinscr_dia&"/"&str_dt_rgproinscr_ano&"'"		
	else
		str_dt_rgproinscr = "NULL"
	end if

str_dt_rgpropag_dia = request("dt_rgpropag_dia")
str_dt_rgpropag_mes = request("dt_rgpropag_mes")
str_dt_rgpropag_ano = request("dt_rgpropag_ano")
	if str_dt_rgpropag_dia <> "" AND  str_dt_rgpropag_mes <> "" AND  str_dt_rgpropag_ano <> "" Then
		str_dt_rgpropag = "'"&str_dt_rgpropag_mes&"/"&str_dt_rgpropag_dia&"/"&str_dt_rgpropag_ano&"'"
	else
		str_dt_rgpropag = "NULL"
	end if

str_dt_rgproval_dia = request("dt_rgproval_dia")
str_dt_rgproval_mes = request("dt_rgproval_mes")
str_dt_rgproval_ano = request("dt_rgproval_ano")
	if str_dt_rgproval_dia <> "" AND  str_dt_rgproval_mes <> "" AND  str_dt_rgproval_ano <> "" Then
		str_dt_rgproval = "'"&str_dt_rgproval_mes&"/"&str_dt_rgproval_dia&"/"&str_dt_rgproval_ano&"'"		
	else
		str_dt_rgproval = "NULL"
	end if

strper_dia_i = request("per_dia_i")
strper_mes_i = request("per_mes_i")
strper_ano_i = request("per_ano_i")
	if strper_dia_i <> "" AND  strper_mes_i <> "" AND  strper_ano_i <> "" Then
		str_periodo_i = "'"&strper_mes_i&"/"&strper_dia_i&"/"&strper_ano_i&"'"
	else
		str_periodo_i = "NULL"
	end if

strper_dia_f = request("per_dia_f")
strper_mes_f = request("per_mes_f")
strper_ano_f = request("per_ano_f")
	if strper_dia_f <> "" AND  strper_mes_f <> "" AND  strper_ano_f <> "" Then
		str_periodo_f = "'"&strper_mes_f&"/"&strper_dia_f&"/"&strper_ano_f&"'"
	else
		str_periodo_f = "NULL"
	end if

dt_dia_pgto = request("dt_dia_pgto")
dt_mes_pgto = request("dt_mes_pgto")
dt_ano_pgto = request("dt_ano_pgto")
	if dt_dia_pgto <> "" AND  dt_mes_pgto <> "" AND  dt_ano_pgto <> "" Then
		dt_pagamento_ferias = "'"&dt_mes_pgto&"/"&dt_dia_pgto&"/"&dt_ano_pgto&"'"
	else
		dt_pagamento_ferias = "NULL"
	end if

valor_pgto_ferias = request("valor_pgto_ferias")

str_endereco_alt = request("nm_endereco_alt")

cd_transporte = request("cd_transporte")
'mes_transp_atualizacao = request("mes_transp_atualizacao")
'ano_transp_atualizacao = request("ano_transp_atualizacao")
mes_transp_atualizacao = 1
ano_transp_atualizacao = 2012

cd_refeicao = request("cd_refeicao")
'mes_refeic_atualizacao = request("mes_refeic_atualizacao")
'ano_refeic_atualizacao = request("ano_refeic_atualizacao")
mes_refeic_atualizacao = 1
ano_refeic_atualizacao = 2012

cd_contato_origem = request("cd_contato_origem")
id_contato_origem = request("id_contato_origem")


dt_ano = YEAR(Now)
dt_mes = Month(Now)
dt_dia = Day(now)
dt_hora = hour(now)
dt_minuto = minute(now)

fecha_janela = request("fecha_janela")

'*****************************
'*** INSERE AS OCORRENCIAS ***
'*****************************%>
<!--#include file="../../ocorrencia/acoes/ocorrencias_acao.asp"-->
<br>


		

<%'***** MOSTRA OS CAMPOS RECEBIDOS PARA INSERÇÃO NO BANCO DE DADOS *******


response.write("up_funcionario_insert<br> @cd_matricula='"&strcd_matricula&"',<br>@nm_nome='"&strnome&"', <br>@cd_sexo='"&strcd_sexo&"',<br> @nm_foto='"&strfoto&"', @dt_nascimento="&strdt_nascimento&", @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"',<br> @nm_pai='"&strpai&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_complemento='"&strnm_complemento&"',<br> @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"',@nm_cep='"&strnm_cep&"', @nm_bairro = '"&strbairro&"',<br> @nm_nacionalidade='"&str_nacionalidade&"', @cd_pis='"&str_pis&"',@cd_ctps='"&str_ctps&"',@cd_ctps_serie='"&str_ctps_serie&"',<br> @cd_assistencia_medica_matricula='"&str_assistencia_medica_matricula&"',<br> @nr_assistencia_medica_coparticipacao="&str_assistencia_medica_coparticipacao&",<br> @cd_assistencia_odonto_matricula='"&str_assistencia_odonto_matricula&"',<br> @nr_assistencia_odonto_coparticipacao="&str_assistencia_odonto_coparticipacao&",<br> @cd_vr='"&str_vr&"',<br>@cd_vale_alimentacao='"&str_vale_alimentacao&"',@cd_sptrans='"&str_sptrans&"',<br>@cd_cmt_bom='"&str_cmt_bom&"',<br> @nm_banco='"&str_banco&"', <br>@cd_banco_ag='"&str_banco_ag&"', @cd_banco_conta='"&str_banco_conta&"',<br> @nm_tit_eleitor='"&nm_tit_eleitor&"',<br>@nm_reservista='"&nm_reservista&"',<br> @cd_conspro='"&str_conspro&"',<br>@cd_numreg='"&str_numreg&"',<br>@rgpro_concessao='"&str_rgpro_concessao&"',<br>@cd_rgpro_status='"&str_rgpro_status&"', @obs_rgpro='"&str_rgpro&"',<br>@dt_rgproinscr="&str_dt_rgproinscr&",@dt_rgpropag="&str_dt_rgpropag&",<br>@dt_rgproval="&str_dt_rgproval&",<br> @nm_endereco_alt='"&str_endereco_alt&"', <br>@cd_estado_civil='"&cd_estado_civil&"',@nm_conjuge='"&nm_conjuge&"',<br> @dt_conjuge_nasc="&dt_conjuge_nasc&",<br>@cd_plano_saude_conj='"&cd_plano_saude_conj&"',@nm_naturalidade='"&nm_naturalidade&"',<br> @dt_rg_emissao="&dt_rg_emissao&", <br>@nm_rg_expedidor="&nm_rg_expedidor&", <br>@nr_tit_eleitor_zona='"&nr_tit_eleitor_zona&"',<br> @nr_tit_eleitor_seccao='"&nr_tit_eleitor_seccao&"', @nm_reservista_cat='"&nm_reservista_cat&"', <br>@nr_hepatite_b='"&nr_hepatite_b&"',<br> @nr_dupla_adulto='"&nr_dupla_adulto&"', <br>@dt_dupla_adulto_validade="&dt_dupla_adulto_validade&", <br>@nr_scr='"&nr_scr&"',<br> @dt_exame_medico="&dt_exame_medico&", <br>@nr_caract_altura='"&nr_caract_altura&"',<br> @nr_caract_peso='"&nr_caract_peso&"',<br>@nr_caract_manequim='"&nr_caract_manequim&"',<br>@nr_caract_calcado='"&nr_caract_calcado&"',<br>@nm_parente='"&nm_parente&"',<br>@cd_parente_parentesco='"&cd_parente_parentesco&"', <br>@cd_residencia_tipo='"&cd_residencia_tipo&"',<br> @cd_residencia_financ='"&cd_residencia_financ&"', <br>@cd_residencia_reside='"&cd_residencia_reside&"',<br> @cd_veiculo_tipo='"&cd_veiculo_tipo&"',<br>@cd_veiculo_financ='"&cd_veiculo_financ&"',<br>@dt_admissao_geral="&Strdt_contratado&"<br>*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-")



'**********************************************************************
'*** 					Insere funcionario						    ***
'**********************************************************************
if session_cd_usuario = 46 then
	x = ""
end if


if x = "" then
	If stracao = "insert" Then
		'xsql = "up_funcionario_insert @cd_matricula='"&strcd_matricula&"',@nm_nome='"&strnome&"', @cd_sexo='"&strcd_sexo&"', @nm_foto='"&strfoto&"', @dt_nascimento="&strdt_nascimento&", @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"', @nm_pai='"&strpai&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_complemento='"&strnm_complemento&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"',@nm_cep='"&strnm_cep&"', @nm_bairro = '"&strbairro&"', @nm_nacionalidade='"&str_nacionalidade&"', @cd_pis='"&str_pis&"',@cd_ctps='"&str_ctps&"',@cd_ctps_serie='"&str_ctps_serie&"', @cd_assistencia_medica_matricula='"&str_assistencia_medica_matricula&"', @nr_assistencia_medica_coparticipacao="&str_assistencia_medica_coparticipacao&", @cd_assistencia_odonto_matricula='"&str_assistencia_odonto_matricula&"', @nr_assistencia_odonto_coparticipacao="&str_assistencia_odonto_coparticipacao&", @cd_vr='"&str_vr&"',@cd_vale_alimentacao='"&str_vale_alimentacao&"',@cd_sptrans='"&str_sptrans&"',@cd_cmt_bom='"&str_cmt_bom&"', @nm_banco='"&str_banco&"', @cd_banco_ag='"&str_banco_ag&"', @cd_banco_conta='"&str_banco_conta&"', @nm_tit_eleitor='"&nm_tit_eleitor&"',@nm_reservista='"&nm_reservista&"', @cd_conspro='"&str_conspro&"',@cd_rgpro_cargo='"&str_rgpro_cargo&"',@cd_numreg='"&str_numreg&"',@rgpro_concessao='"&str_rgpro_concessao&"',@cd_rgpro_status='"&str_rgpro_status&"', @obs_rgpro='"&str_rgpro&"',@dt_rgproinscr="&str_dt_rgproinscr&",@dt_rgpropag="&str_dt_rgpropag&",@dt_rgproval="&str_dt_rgproval&", @nm_endereco_alt='"&str_endereco_alt&"', @cd_estado_civil='"&cd_estado_civil&"',@nm_conjuge='"&nm_conjuge&"', @dt_conjuge_nasc="&dt_conjuge_nasc&",@cd_plano_saude_conj='"&cd_plano_saude_conj&"',@nm_naturalidade='"&nm_naturalidade&"', @dt_rg_emissao="&dt_rg_emissao&", @nm_rg_expedidor="&nm_rg_expedidor&", @nr_tit_eleitor_zona='"&nr_tit_eleitor_zona&"', @nr_tit_eleitor_seccao='"&nr_tit_eleitor_seccao&"', @nm_reservista_cat='"&nm_reservista_cat&"', @nr_hepatite_b='"&nr_hepatite_b&"', @nr_dupla_adulto='"&nr_dupla_adulto&"', @dt_dupla_adulto_validade="&dt_dupla_adulto_validade&", @nr_scr='"&nr_scr&"', @dt_exame_medico="&dt_exame_medico&", @nr_caract_altura='"&nr_caract_altura&"', @nr_caract_peso='"&nr_caract_peso&"',@nr_caract_manequim='"&nr_caract_manequim&"',@nr_caract_calcado='"&nr_caract_calcado&"',@nm_parente='"&nm_parente&"',@cd_parente_parentesco='"&cd_parente_parentesco&"', @cd_residencia_tipo='"&cd_residencia_tipo&"', @cd_residencia_financ='"&cd_residencia_financ&"', @cd_residencia_reside='"&cd_residencia_reside&"', @cd_veiculo_tipo='"&cd_veiculo_tipo&"',@cd_veiculo_financ='"&cd_veiculo_financ&"',@dt_admissao_geral="&Strdt_contratado&""
		'dbconn.execute(xsql)
		
		xsql = "up_funcionario_insert @cd_matricula='"&strcd_matricula&"', @nm_nome='"&strnome&"', @cd_sexo='"&strcd_sexo&"', @nm_foto='"&strfoto&"', @dt_nascimento="&strdt_nascimento&", @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"', @nm_pai='"&strpai&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_complemento='"&strnm_complemento&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"', @nm_cep='"&strnm_cep&"', @nm_bairro = '"&strbairro&"', @nm_nacionalidade='"&str_nacionalidade&"', @cd_pis='"&str_pis&"', @cd_ctps='"&str_ctps&"', @cd_ctps_serie='"&str_ctps_serie&"', @cd_assistencia_medica_matricula='"&str_assistencia_medica_matricula&"', @cd_assistencia_odonto_matricula='"&str_assistencia_odonto_matricula&"', @cd_vr='"&str_vr&"', @cd_vale_alimentacao='"&str_vale_alimentacao&"', @cd_sptrans='"&str_sptrans&"', @cd_cmt_bom='"&str_cmt_bom&"', @nm_banco='"&str_banco&"', @cd_banco_ag='"&str_banco_ag&"', @cd_banco_conta='"&str_banco_conta&"', @nm_tit_eleitor='"&nm_tit_eleitor&"', @nm_reservista='"&nm_reservista&"', @cd_conspro='"&str_conspro&"', @cd_rgpro_cargo='"&str_rgpro_cargo&"', @cd_numreg='"&str_numreg&"', @rgpro_concessao='"&str_rgpro_concessao&"', @cd_rgpro_status='"&str_rgpro_status&"', @obs_rgpro='"&str_rgpro&"', @dt_rgproinscr="&str_dt_rgproinscr&", @dt_rgpropag="&str_dt_rgpropag&", @dt_rgproval="&str_dt_rgproval&", @nm_endereco_alt='"&str_endereco_alt&"', @cd_estado_civil='"&cd_estado_civil&"', @nm_conjuge='"&nm_conjuge&"', @dt_conjuge_nasc="&dt_conjuge_nasc&", @cd_plano_saude_conj='"&cd_plano_saude_conj&"', @nm_naturalidade='"&nm_naturalidade&"', @dt_rg_emissao="&dt_rg_emissao&", @nm_rg_expedidor="&nm_rg_expedidor&", @nr_tit_eleitor_zona='"&nr_tit_eleitor_zona&"', @nr_tit_eleitor_seccao='"&nr_tit_eleitor_seccao&"', @nm_reservista_cat='"&nm_reservista_cat&"', @nr_hepatite_b='"&nr_hepatite_b&"', @nr_dupla_adulto='"&nr_dupla_adulto&"', @dt_dupla_adulto_validade="&dt_dupla_adulto_validade&", @nr_scr='"&nr_scr&"', @dt_exame_medico="&dt_exame_medico&", @nm_parente='"&nm_parente&"', @cd_parente_parentesco='"&cd_parente_parentesco&"', @dt_admissao_geral="&Strdt_contratado&""
		dbconn.execute(xsql)
		
			'***************************************************************************
			'Cargo: Verifica o nome do funcionario recém registrado e resgata seu código
			'***************************************************************************
			'strsql = "Select * From TBL_funcionario Where nm_nome='"&strnome&"' AND nm_sobrenome='"&strnm_sobrenome&"'"
			strsql = "Select * From TBL_funcionario Where nm_nome='"&strnome&"' AND dt_nascimento="&strdt_nascimento&""
			Set rs = dbconn.execute(strsql)
				if Not rs.EOF Then
				cd_cod_funcionario = rs("cd_codigo")
				response.write("<***"&cd_cod_funcionario&"***>")
				end if
				
				
				
			'*********************************************************
			'*** Insere a data do primeiro contrato do funcionario ***
			'*********************************************************
			xsql = "up_funcionario_contrato_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_contrato='"&strregime_trabalho&"', @dt_admissao="&Strdt_contratado&", @dt_atualizacao='"&dt_atualizacao&"'"
			dbconn.execute(xsql)
			
			'Insere o cargo do funcionario recem cadastrado
				xsql = "up_funcionario_cargo_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_cargo='"&str_cargo&"', @dt_atualizacao='"&dt_atualizacao&"'"
				dbconn.execute(xsql)
				
			'Insere o status do funcionário
				xsql = "up_funcionario_status_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_status='"&str_status&"', @dt_atualizacao="&Strdt_contratado&""
				dbconn.execute(xsql)
				
			'Insere a unidade do funcionário
				xsql = "up_funcionario_unidade_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_unidade='"&strcd_unidade&"', @dt_atualizacao="&Strdt_contratado&""
				dbconn.execute(xsql)
				
			'Insere a função do funcionario recem cadastrado
				xsql = "up_funcionario_funcao_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_funcao='"&str_funcao&"', @dt_atualizacao="&Strdt_contratado&""
				dbconn.execute(xsql)
			
			'Insere o horário do funcionario recem cadastrado
				xsql = "up_funcionario_horario_insert @cd_funcionario='"&cd_cod_funcionario&"', @hr_entrada='"&strhr_entrada&"', @nm_intervalo='"&strnm_intervalo&"', @hr_saida='"&strhr_saida&"', @dt_atualizacao="&Strdt_contratado&", @cd_user='"&cd_user&"' "
				dbconn.execute(xsql)
				
				'Registra o mes posterior à sua entrada para pagamento do sindicato
				'dt_ano_post = year(Strdt_contratado)
				'dt_mes_post = month(Strdt_contratado)+1
				'	if dt_mes_post = 13 Then
				'	dt_mes_post = 1
				'	dt_ano_post = dt_ano_post + 1
				'	end if
				
			'********************************************************************
			'******** Insere Contatos *******************************************
			'********************************************************************
				cd_origem = 4 '- 
				id_contato_origem = cd_cod_funcionario
				response.write("+++"&str_contatos_total&"+++")
			
				If str_contatos_total <> "" Then
					str_contatos_total = replace(str_contatos_total,"$$","$")
					contatos_array = split(str_contatos_total,"$")
						
						For Each contatos_item_junto In contatos_array
								
							if contatos_item_junto <> "" Then
								contatos_item_junto = replace(contatos_item_junto,"%20"," ")
								contatos_item_junto = replace(contatos_item_junto,"Ã§","ç")
							
								separador1 = int(instr(contatos_item_junto,"!")) 'diz onde está o sinal, para separação de cada item
								separador2 = instr(contatos_item_junto,"{")
								separador3 = instr(contatos_item_junto,"}")
								separador4 = instr(contatos_item_junto,"[")
								separador5 = instr(contatos_item_junto,"]")
								separador6 = instr(contatos_item_junto,"_")
								
								contatos_tipo =  mid(contatos_item_junto,1,separador1)
									contatos_tipo = replace(contatos_tipo,"!","")
								contatos_nome =  mid(contatos_item_junto,separador1,separador2-separador1)
									contatos_nome = replace(contatos_nome,"!","")
								contatos_cargo = mid(contatos_item_junto,separador2,separador3-separador2)
									contatos_cargo = replace(contatos_cargo,"{","")
								contatos_ddd =   mid(contatos_item_junto,separador3,separador4-separador3)
									contatos_ddd = replace(contatos_ddd,"}","")
								contatos_numero = mid(contatos_item_junto,separador4,separador5-separador4)
									contatos_numero = replace(contatos_numero,"[","")
								contatos_obs =   mid(contatos_item_junto,separador5,separador6-separador5)
									contatos_obs = replace(contatos_obs,"]","")
							end if
							
							xsql = "up_contato_insert @nm_nome='"&contatos_nome&"', @nm_cargo='"&contatos_cargo&"', @cd_categoria='"&contatos_tipo&"',@cd_ddd='"&contatos_ddd&"',@nm_numero='"&contatos_numero&"', @nm_obs='"&contatos_obs&"', @cd_origem='"&cd_contato_origem&"', @id_origem='"&id_contato_origem&"'"
							dbconn.execute(xsql)
						next
				end if
			
			'***********************************************************************************
			'******** Insere Dependentes *******************************************************
			'***********************************************************************************
				If str_dependentes_total <> "" Then
					str_dependentes_total = replace(str_dependentes_total,"$$","$")
					dependentes_array = split(str_dependentes_total,"$")
						
						For Each dependentes_item_junto In dependentes_array
							dependentes_item_junto = replace(dependentes_item_junto,"%20"," ")
							dependentes_item_junto = replace(dependentes_item_junto,"Ã§","ç")
							
							separador1 = int(instr(dependentes_item_junto,"!")) 'diz onde está o sinal, para separação de cada item
							separador2 = instr(dependentes_item_junto,"{")
							separador3 = instr(dependentes_item_junto,"}")
							separador4 = instr(dependentes_item_junto,"[")
							separador5 = instr(dependentes_item_junto,"]")
							separador6 = instr(dependentes_item_junto,"_")
							
							dependentes_nome =  mid(dependentes_item_junto,1,separador1) 'Mostra o primeiro compo
								dependentes_nome = replace(dependentes_nome,"!","")
							dependentes_parentesco =  mid(dependentes_item_junto,separador1,separador2-separador1)'Mostra o segundo campo
								dependentes_parentesco = replace(dependentes_parentesco,"!","")
							dependentes_nascimento = mid(dependentes_item_junto,separador2,separador3-separador2)'Mostra o terceiro campo
								dependentes_nascimento = replace(dependentes_nascimento,"{","")
								if dependentes_nascimento <> "" Then								
										dependentes_nascimento = "'"&month(dependentes_nascimento)&"/"&day(dependentes_nascimento)&"/"&year(dependentes_nascimento)&"'"
									else
										dependentes_nascimento = "Null"
									end if
							dependentes_obs =   mid(dependentes_item_junto,separador3,separador4-separador3)
								dependentes_obs = replace(dependentes_obs,"}","")
							dependentes_sexo = mid(dependentes_item_junto,separador4,separador5-separador4)
								dependentes_sexo = replace(dependentes_sexo,"[","")
						
							'xsql = "up_dependente_insert @cd_funcionario='"&cd_cod_funcionario&"', @nm_nome='"&dependentes_nome&"', @cd_parentesco='"&dependentes_parentesco&"', @cd_sexo='"&dependentes_sexo&"', @dt_nascimento='"&month(dependentes_nascimento)&"/"&day(dependentes_nascimento)&"/"&year(dependentes_nascimento)&"', @nm_obs='"&dependentes_obs&"'"
							xsql = "up_dependente_insert @cd_funcionario='"&cd_cod_funcionario&"', @nm_nome='"&dependentes_nome&"', @cd_parentesco='"&dependentes_parentesco&"', @cd_sexo='"&dependentes_sexo&"', @dt_nascimento="&dependentes_nascimento&", @nm_obs='"&dependentes_obs&"'"
							dbconn.execute(xsql)
						next
				end if
			
			'*****************************************************************************
			'******** escolaridade *******************************************************
			'*****************************************************************************
				If str_escolaridade_total <> "" Then
					str_escolaridade_total = replace(str_escolaridade_total,"$$","$")
					escolaridade_array = split(str_escolaridade_total,"$")
						
						For Each escolaridade_item_junto In escolaridade_array
								
							if escolaridade_item_junto <> "" Then
								escolaridade_item_junto = replace(escolaridade_item_junto,"%20"," ")
								escolaridade_item_junto = replace(escolaridade_item_junto,"Ã§","ç")
								separador1 = int(instr(escolaridade_item_junto,"!"))
								separador2 = instr(escolaridade_item_junto,"{")
								separador3 = instr(escolaridade_item_junto,"}")
								separador4 = instr(escolaridade_item_junto,"[")
								separador5 = instr(escolaridade_item_junto,"]")
								separador6 = instr(escolaridade_item_junto,"_")
								
								escolaridade_ensino_grau = 	mid(escolaridade_item_junto,1,separador1) 
									escolaridade_ensino_grau = replace(escolaridade_ensino_grau,"!","")
								escolaridade_instituicao = 	mid(escolaridade_item_junto,separador1,separador2-separador1)
									escolaridade_instituicao = replace(escolaridade_instituicao,"!","")
								escolaridade_curso = 		mid(escolaridade_item_junto,separador2,separador3-separador2)
									escolaridade_curso = replace(escolaridade_curso,"{","")
								escolaridade_andamento = 	mid(escolaridade_item_junto,separador3,separador4-separador3)
									escolaridade_andamento = replace(escolaridade_andamento,"}","")
								escolaridade_termino = 		mid(escolaridade_item_junto,separador4,separador5-separador4)
									escolaridade_termino = replace(escolaridade_termino,"[","")
									if str_escolaridade_termino <> "" Then
										escolaridade_termino = "'"&month(str_escolaridade_termino)&"/"&day(str_escolaridade_termino)&"/"&year(str_escolaridade_termino)&"'"
									else
										escolaridade_termino = "Null"
									end if
								escolaridade_obs = 	mid(escolaridade_item_junto,separador5,separador6-separador5)
									escolaridade_obs = replace(escolaridade_obs,"]","")	
									'ver_ter_esc = separador1+separador2+separador3+separador4+1
							end if
							
							xsql = "up_escolaridade_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_ensino_grau='"&escolaridade_ensino_grau&"', @cd_ensino_andamento='"&escolaridade_andamento&"', @nm_instituicao='"&escolaridade_instituicao&"', @nm_curso='"&escolaridade_curso&"', @dt_termino="&escolaridade_termino&", @nm_obs='"&escolaridade_obs&"'"
							dbconn.execute(xsql)
						next
				end if
				
			'********************************************************************************
			'******** Emprego anterior*******************************************************
			'********************************************************************************
				
				If str_emprego_total <> "" Then
					str_emprego_total = replace(str_emprego_total,"$$","$")
					emprego_array = split(str_emprego_total,"$")
						
						For Each emprego_item_junto In emprego_array
							emprego_item_junto = replace(emprego_item_junto,"%20"," ")
							emprego_item_junto = replace(emprego_item_junto,"Ã§","ç")
								
							if len(emprego_item_junto) > 3 then
								separador1 = int(instr(emprego_item_junto,"!"))
								separador2 = instr(emprego_item_junto,"{")
								separador3 = instr(emprego_item_junto,"}")
								separador4 = instr(emprego_item_junto,"[")
								separador5 = instr(emprego_item_junto,"]")
								'separador6 = instr(emprego_item_junto,"_")
								
								emprego_empresa = mid(emprego_item_junto,1,separador1)
									emprego_empresa = replace(emprego_empresa,"!","")
								emprego_funcao = mid(emprego_item_junto,separador1,separador2-separador1)
									emprego_funcao = replace(emprego_funcao,"!","")
								emprego_entrada = mid(emprego_item_junto,separador2,separador3-separador2)
									emprego_entrada = replace(emprego_entrada,"{","")
									if emprego_entrada <> "" Then
										emprego_entrada = "'"&month(emprego_entrada)&"/"&day(emprego_entrada)&"/"&year(emprego_entrada)&"'"
									else
										emprego_entrada = "Null"
									end if
								emprego_saida = mid(emprego_item_junto,separador3,separador4-separador3)
									emprego_saida = replace(emprego_saida,"}","")
										if emprego_saida <> "" Then
											emprego_saida = "'"&month(emprego_saida)&"/"&day(emprego_saida)&"/"&year(emprego_saida)&"'"
										else
											emprego_saida = "Null"
										end if
								emprego_obs = mid(emprego_item_junto,separador4,separador5-separador4)
									emprego_obs = replace(emprego_obs,"["," ")
									
								xsql = "up_emprego_anterior_insert @cd_funcionario='"&cd_cod_funcionario&"', @emprego_empresa='"&emprego_empresa&"', @emprego_funcao='"&emprego_funcao&"', @emprego_entrada="&emprego_entrada&", @emprego_saida="&emprego_saida&", @emprego_obs='"&emprego_obs&"'"
								dbconn.execute(xsql)
							end if
						next
						
				end if
				
				'******** Valores *******************************************************
				
				'If str_valores_total <> "" Then
				'	str_valores_total = replace(str_valores_total,"$$","$")
				'	valores_array = split(str_valores_total,"$")
				'		
				'		
				'		For Each valores_item_junto In valores_array
				'			emprego_item_len = len(valores_item_junto) 'tamanho da sentença toda
				'			
				'		separador1 = int(instr(valores_item_junto,"!")) 'diz onde está o sinal, para separação de cada item
				'		separador2 = int(instr(mid(valores_item_junto,(separador1)+1,emprego_item_len),"!")) 
				'		separador3 = int(instr(mid(valores_item_junto,(separador1+separador2)+1,emprego_item_len),"!")) 
				'		separador4 = int(instr(mid(valores_item_junto,(separador1+separador2+separador3)+1,emprego_item_len),"!"))
				'		
				'		cd_valor_tipo = mid(valores_item_junto,1,separador1-1) 'Mostra o primeiro compo
				'		nr_valor = formatnumber(mid(valores_item_junto,separador1+1,separador2-1),2)'Mostra o segundo campo
				'		dt_valor_atualizacao = mid(valores_item_junto,separador1+separador2+1,separador3-1)'Mostra o terceiro campo
				'		nm_valor_obs = mid(valores_item_junto,separador1+separador2+separador3+separador4+1,separador5-1)
				'			nm_valor_obs = replace(nm_valor_obs,"%20"," ")
				'			
				'			xsql = "up_funcionario_valor_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_valor_tipo='"&cd_valor_tipo&"', @nr_valor='"&nr_valor&"', @dt_valor_atualizacao='"&month(dt_valor_atualizacao)&"/"&day(dt_valor_atualizacao)&"/"&year(dt_valor_atualizacao)&"',  @nm_valor_obs='"&nm_valor_obs&"'"
				'			dbconn.execute(xsql)
				'		next
						
				'end if
				
				%>
				<body onLoad="window.close(); window.opener.document.getElementById('form').submit(true);">
				<%'response.redirect "../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod="&cd_cod_funcionario&"&msg="&strmsg&"&func_ativos="&func_ativos&""
	
	
'**********************************************************************
'*** 					Altera funcionario						    ***
'**********************************************************************
	
	ElseIf stracao = "altera" Then
		if strstatus <> 4 Then ' Altera
		xsql = "up_Funcionario_update @cd_matricula='"&strcd_matricula&"',@nm_nome='"&strnome&"', @cd_sexo='"&strcd_sexo&"', @nm_foto='"&strfoto&"', @dt_nascimento="&strdt_nascimento&", @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"', @nm_pai='"&strpai&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_bairro='"&strbairro&"', @nm_complemento='"&strnm_complemento&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"',@nm_cep='"&strnm_cep&"', @cd_codigo='"&strcod&"', @nm_nacionalidade='"&str_nacionalidade&"', @cd_pis='"&str_pis&"',@cd_ctps='"&str_ctps&"',@cd_ctps_serie='"&str_ctps_serie&"', @cd_assistencia_medica_matricula='"&str_assistencia_medica_matricula&"', @nr_assistencia_medica_coparticipacao="&str_assistencia_medica_coparticipacao&", @cd_assistencia_odonto_matricula='"&str_assistencia_odonto_matricula&"', @nr_assistencia_odonto_coparticipacao="&str_assistencia_odonto_coparticipacao&", @cd_vr='"&str_vr&"',@cd_vale_alimentacao='"&str_vale_alimentacao&"',@cd_sptrans='"&str_sptrans&"',@cd_cmt_bom='"&str_cmt_bom&"', @nm_banco='"&str_banco&"', @cd_banco_ag='"&str_banco_ag&"', @cd_banco_conta='"&str_banco_conta&"', @nm_tit_eleitor='"&nm_tit_eleitor&"',@nm_reservista='"&nm_reservista&"', @cd_conspro='"&str_conspro&"',@cd_rgpro_cargo='"&str_rgpro_cargo&"',@cd_numreg='"&str_numreg&"',@rgpro_concessao='"&str_rgpro_concessao&"',@cd_rgpro_status='"&str_rgpro_status&"', @obs_rgpro='"&str_rgpro&"',@dt_rgproinscr="&str_dt_rgproinscr&",@dt_rgpropag="&str_dt_rgpropag&",@dt_rgproval="&str_dt_rgproval&", @nm_endereco_alt='"&str_endereco_alt&"', @cd_estado_civil='"&cd_estado_civil&"', @nm_conjuge='"&nm_conjuge&"',@dt_conjuge_nasc="&dt_conjuge_nasc&", @cd_plano_saude_conj='"&cd_plano_saude_conj&"', @nm_naturalidade='"&nm_naturalidade&"', @dt_rg_emissao="&dt_rg_emissao&", @nm_rg_expedidor="&nm_rg_expedidor&", @nr_tit_eleitor_zona='"&nr_tit_eleitor_zona&"', @nr_tit_eleitor_seccao='"&nr_tit_eleitor_seccao&"',@nm_reservista_cat='"&nm_reservista_cat&"',@nr_hepatite_b='"&nr_hepatite_b&"', @nr_dupla_adulto='"&nr_dupla_adulto&"', @dt_dupla_adulto_validade="&dt_dupla_adulto_validade&",@nr_scr='"&nr_scr&"', @dt_exame_medico="&dt_exame_medico&", @nr_caract_altura='"&nr_caract_altura&"', @nr_caract_peso='"&nr_caract_peso&"',@nr_caract_manequim='"&nr_caract_manequim&"',@nr_caract_calcado='"&nr_caract_calcado&"',@nm_parente='"&nm_parente&"',@cd_parente_parentesco='"&cd_parente_parentesco&"', @cd_residencia_tipo='"&cd_residencia_tipo&"', @cd_residencia_financ='"&cd_residencia_financ&"', @cd_residencia_reside='"&cd_residencia_reside&"', @cd_veiculo_tipo='"&cd_veiculo_tipo&"',@cd_veiculo_financ='"&cd_veiculo_financ&"'"
		dbconn.execute(xsql)
			
			'-------------------------------------------------------------------------
			'******** Contatos *******************************************************
			'-------------------------------------------------------------------------
				cd_origem = 4 '- 
				id_contato_origem = strcod' cd_cod_funcionario
					
					If str_contatos_total <> "" Then
					str_contatos_total = replace(str_contatos_total,"$$","$")
					contatos_array = split(str_contatos_total,"$")
					
					response.write(str_contatos_total)
						
						For Each contatos_item_junto In contatos_array
								
							if contatos_item_junto <> "" Then
								contatos_item_junto = replace(contatos_item_junto,"%20"," ")
								contatos_item_junto = replace(contatos_item_junto,"Ã§","ç")
							
								separador1 = int(instr(contatos_item_junto,"!")) 'diz onde está o sinal, para separação de cada item
								separador2 = instr(contatos_item_junto,"{")
								separador3 = instr(contatos_item_junto,"}")
								separador4 = instr(contatos_item_junto,"[")
								separador5 = instr(contatos_item_junto,"]")
								separador6 = instr(contatos_item_junto,"_")
								
								contatos_tipo =  mid(contatos_item_junto,1,separador1)
									contatos_tipo = replace(contatos_tipo,"!","")
								contatos_nome =  mid(contatos_item_junto,separador1,separador2-separador1)
									contatos_nome = replace(contatos_nome,"!","")
								contatos_cargo = mid(contatos_item_junto,separador2,separador3-separador2)
									contatos_cargo = replace(contatos_cargo,"{","")
								contatos_ddd =   mid(contatos_item_junto,separador3,separador4-separador3)
									contatos_ddd = replace(contatos_ddd,"}","")
								contatos_numero = mid(contatos_item_junto,separador4,separador5-separador4)
									contatos_numero = replace(contatos_numero,"[","")
								contatos_obs =   mid(contatos_item_junto,separador5,separador6-separador5)
									contatos_obs = replace(contatos_obs,"]","")
									
									
							end if
							
							xsql = "up_contato_insert @nm_nome='"&contatos_nome&"', @nm_cargo='"&contatos_cargo&"', @cd_categoria='"&contatos_tipo&"',@cd_ddd='"&contatos_ddd&"',@nm_numero='"&contatos_numero&"', @nm_obs='"&contatos_obs&"', @cd_origem='"&cd_contato_origem&"', @id_origem='"&id_contato_origem&"'"
							'response.write xsql
							dbconn.execute(xsql)
						next
				end if
				
				'----------------------------------------------------------------------------
				'******** Dependentes *******************************************************
				'----------------------------------------------------------------------------
				If str_dependentes_total <> "" Then
					str_dependentes_total = replace(str_dependentes_total,"$$","$")
					dependentes_array = split(str_dependentes_total,"$")
												
						For Each dependentes_item_junto In dependentes_array
							dependentes_item_junto = replace(dependentes_item_junto,"%20"," ")
							dependentes_item_junto = replace(dependentes_item_junto,"Ã§","ç")
							
							separador1 = int(instr(dependentes_item_junto,"!")) 'diz onde está o sinal, para separação de cada item
							separador2 = instr(dependentes_item_junto,"{")
							separador3 = instr(dependentes_item_junto,"}")
							separador4 = instr(dependentes_item_junto,"[")
							separador5 = instr(dependentes_item_junto,"]")
							separador6 = instr(dependentes_item_junto,"_")
							
							dependentes_nome =  mid(dependentes_item_junto,1,separador1) 'Mostra o primeiro compo
								dependentes_nome = replace(dependentes_nome,"!","")
							dependentes_parentesco =  mid(dependentes_item_junto,separador1,separador2-separador1)'Mostra o segundo campo
								dependentes_parentesco = replace(dependentes_parentesco,"!","")
							dependentes_nascimento = mid(dependentes_item_junto,separador2,separador3-separador2)'Mostra o terceiro campo
								dependentes_nascimento = replace(dependentes_nascimento,"{","")
								if dependentes_nascimento <> "" Then								
										dependentes_nascimento = "'"&month(dependentes_nascimento)&"/"&day(dependentes_nascimento)&"/"&year(dependentes_nascimento)&"'"
									else
										dependentes_nascimento = "Null"
									end if
							dependentes_obs =   mid(dependentes_item_junto,separador3,separador4-separador3)
								dependentes_obs = replace(dependentes_obs,"}","")
							dependentes_sexo = mid(dependentes_item_junto,separador4,separador5-separador4)
								dependentes_sexo = replace(dependentes_sexo,"[","")
							
								'*** Verifica se o dependente já consta no BD. ***
								xsql_verif_dep = "SELECT * FROM TBL_dependente WHERE cd_funcionario='"&strcod&"' AND (nm_nome LIKE '%"&left(dependentes_nome,7)&"%')"
								Set rs_verif_dep = dbconn.execute(xsql_verif_dep)
									while not rs_verif_dep.EOF
										existe = rs_verif_dep("cd_codigo")
									rs_verif_dep.movenext
									wend
									
									if existe = "" Then									
										xsql = "up_dependente_insert @nm_nome='"&dependentes_nome&"', @cd_parentesco='"&dependentes_parentesco&"', @cd_sexo='"&dependentes_sexo&"', @dt_nascimento="&dependentes_nascimento&", @nm_obs='"&dependentes_obs&"', @cd_funcionario='"&strcod&"'"
										dbconn.execute(xsql)
									end if
						next
						
				end if
				
				'----------------------------------------------------------------------------
				'******** escolaridade ******************************************************
				'----------------------------------------------------------------------------
				If str_escolaridade_total <> "" Then
					str_escolaridade_total = replace(str_escolaridade_total,"$$","$")
					escolaridade_array = split(str_escolaridade_total,"$")
						
						For Each escolaridade_item_junto In escolaridade_array
								
							if escolaridade_item_junto <> "" Then
								escolaridade_item_junto = replace(escolaridade_item_junto,"%20"," ")
								escolaridade_item_junto = replace(escolaridade_item_junto,"Ã§","ç")
								separador1 = int(instr(escolaridade_item_junto,"!"))
								separador2 = instr(escolaridade_item_junto,"{")
								separador3 = instr(escolaridade_item_junto,"}")
								separador4 = instr(escolaridade_item_junto,"[")
								separador5 = instr(escolaridade_item_junto,"]")
								separador6 = instr(escolaridade_item_junto,"_")
								
								escolaridade_ensino_grau = 	mid(escolaridade_item_junto,1,separador1) 
									escolaridade_ensino_grau = replace(escolaridade_ensino_grau,"!","")
								escolaridade_instituicao = 	mid(escolaridade_item_junto,separador1,separador2-separador1)
									escolaridade_instituicao = replace(escolaridade_instituicao,"!","")
								escolaridade_curso = 		mid(escolaridade_item_junto,separador2,separador3-separador2)
									escolaridade_curso = replace(escolaridade_curso,"{","")
								escolaridade_andamento = 	mid(escolaridade_item_junto,separador3,separador4-separador3)
									escolaridade_andamento = replace(escolaridade_andamento,"}","")
								escolaridade_termino = 		mid(escolaridade_item_junto,separador4,separador5-separador4)
									escolaridade_termino = replace(escolaridade_termino,"[","")
									if escolaridade_termino <> "" Then
										escolaridade_termino = "'"&month(escolaridade_termino)&"/"&day(escolaridade_termino)&"/"&year(escolaridade_termino)&"'"
									else
										escolaridade_termino = "Null"
									end if
								escolaridade_obs = 	mid(escolaridade_item_junto,separador5,separador6-separador5)
									escolaridade_obs = replace(escolaridade_obs,"]","")	
									'ver_ter_esc = separador1+separador2+separador3+separador4+1
							end if
							
							xsql = "up_escolaridade_insert @cd_funcionario='"&strcod&"', @cd_ensino_grau='"&escolaridade_ensino_grau&"', @cd_ensino_andamento='"&escolaridade_andamento&"', @nm_instituicao='"&escolaridade_instituicao&"', @nm_curso='"&escolaridade_curso&"', @dt_termino="&escolaridade_termino&", @nm_obs='"&escolaridade_obs&"'"
							dbconn.execute(xsql)
						next
				end if
				
				'----------------------------------------------------------------------------
				'******** Emprego anterior***************************************************
				'----------------------------------------------------------------------------
				
				If str_emprego_total <> "" Then
					str_emprego_total = replace(str_emprego_total,"$$","$")
					emprego_array = split(str_emprego_total,"$")
						
						For Each emprego_item_junto In emprego_array
							emprego_item_junto = replace(emprego_item_junto,"%20"," ")
							emprego_item_junto = replace(emprego_item_junto,"Ã§","ç")
								
							separador1 = int(instr(emprego_item_junto,"!"))
							separador2 = instr(emprego_item_junto,"{")
							separador3 = instr(emprego_item_junto,"}")
							separador4 = instr(emprego_item_junto,"[")
							separador5 = instr(emprego_item_junto,"]")
							separador6 = instr(emprego_item_junto,"_")
					
							emprego_empresa = mid(emprego_item_junto,1,separador1)
								emprego_empresa = replace(emprego_empresa,"!","")
							emprego_funcao = mid(emprego_item_junto,separador1,separador2-separador1)
								emprego_funcao = replace(emprego_funcao,"!","")
							emprego_entrada = mid(emprego_item_junto,separador2,separador3-separador2)
								emprego_entrada = replace(emprego_entrada,"{","")
								if emprego_entrada <> "" Then
									emprego_entrada = "'"&month(emprego_entrada)&"/"&day(emprego_entrada)&"/"&year(emprego_entrada)&"'"
								else
									emprego_entrada = "Null"
								end if
							emprego_saida = mid(emprego_item_junto,separador3,separador4-separador3)
								emprego_saida = replace(emprego_saida,"}","")
									if emprego_saida <> "" Then
										emprego_saida = "'"&month(emprego_saida)&"/"&day(emprego_saida)&"/"&year(emprego_saida)&"'"
									else
										emprego_saida = "Null"
									end if
							emprego_obs = mid(emprego_item_junto,separador4,separador5-separador4)
								emprego_obs = replace(emprego_obs,"["," ")
							
							xsql = "up_emprego_anterior_insert @emprego_empresa='"&emprego_empresa&"', @emprego_funcao='"&emprego_funcao&"', @emprego_entrada="&emprego_entrada&", @emprego_saida="&emprego_saida&", @emprego_obs='"&emprego_obs&"', @cd_funcionario='"&strcod&"'"
							dbconn.execute(xsql)
						next
						
				end if
				
				'----------------------------------------------------------------------------
				'******** Valores ***********************************************************
				'----------------------------------------------------------------------------
				
				If str_valores_total <> "" Then
					str_valores_total = replace(str_valores_total,"$$","$")
					valores_array = split(str_valores_total,"$")
						
						For Each valores_item_junto In valores_array
						valores_item_len = len(valores_item_junto) 'tamanho da sentença toda
							
						separador1 = int(instr(valores_item_junto,"!")) 'diz onde está o sinal, para separação de cada item
						separador2 = int(instr(mid(valores_item_junto,(separador1)+1,valores_item_len),"!")) 
						separador3 = int(instr(mid(valores_item_junto,(separador1+separador2)+1,valores_item_len),"!")) 
						separador4 = int(instr(mid(valores_item_junto,(separador1+separador2+separador3)+1,valores_item_len),"!"))
						
						cd_valor_tipo 			= mid(valores_item_junto,1,separador1-1) 'Mostra o primeiro compo
						nr_valor   = formatnumber(mid(valores_item_junto,separador1+1,separador2-1),2)'Mostra o segundo campo
						dt_valor_atualizacao 	= mid(valores_item_junto,separador1+separador2+1,separador3-1)'Mostra o terceiro campo
						nm_valor_obs 			= mid(valores_item_junto,separador1+separador2+separador3+1,separador4-1)
							nm_valor_obs = replace(nm_valor_obs,"%20"," ")
							
							xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcod&"', @cd_valor_tipo='"&cd_valor_tipo&"', @nr_valor='"&nr_valor&"', @dt_valor_atualizacao='"&month(dt_valor_atualizacao)&"/"&day(dt_valor_atualizacao)&"/"&year(dt_valor_atualizacao)&"',  @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_sessao&"',@dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
							dbconn.execute(xsql)
						next
				end if
				
				'----------------------------------------------------------------------------
				'******** Tipo Transporte ***************************************************
				'----------------------------------------------------------------------------
				
				If cd_transporte <> "" AND mes_transp_atualizacao <> "" AND ano_transp_atualizacao <> "" Then
				'if cd_transporte <> "" Then
					cd_beneficio_tipo = 2
					cd_beneficio = cd_transporte
					dt_atualizacao = mes_transp_atualizacao&"/1/"&ano_transp_atualizacao
					
					'response.write("cd_beneficio_tipo"&cd_beneficio_tipo&" - cd_beneficio:"&cd_beneficio&" - dt_transp_atualizacao:"&dt_atualizacao)
					
					xsql = "up_funcionario_beneficio_insert @cd_funcionario='"&strcod&"', @cd_beneficio_tipo='"&cd_beneficio_tipo&"', @cd_beneficio='"&cd_beneficio&"', @dt_transp_atualizacao='"&dt_atualizacao&"',  @cd_user='"&cd_sessao&"',@dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
				
				'----------------------------------------------------------------------------
				'******** Tipo Refeição	*****************************************************
				'----------------------------------------------------------------------------
				
				If cd_refeicao <> "" AND mes_refeic_atualizacao <> "" AND ano_refeic_atualizacao <> "" Then
				'if cd_transporte <> "" Then
					cd_beneficio_tipo = 1
					cd_beneficio = cd_refeicao
					dt_atualizacao = mes_refeic_atualizacao&"/1/"&ano_refeic_atualizacao
					
					'response.write("cd_beneficio_tipo"&cd_beneficio_tipo&" - cd_beneficio:"&cd_beneficio&" - dt_transp_atualizacao:"&dt_atualizacao)
					
					xsql = "up_funcionario_beneficio_insert @cd_funcionario='"&strcod&"', @cd_beneficio_tipo='"&cd_beneficio_tipo&"', @cd_beneficio='"&cd_beneficio&"', @dt_transp_atualizacao='"&dt_atualizacao&"',  @cd_user='"&cd_sessao&"',@dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
				
		strmsg = "Colaborador Alterado com sucesso!"	
		'response.redirect "../../empresa.asp?tipo=lista&msg="&strmsg&""
		'response.redirect "../funcionario/funcionarios_cadastro.asp?tipo=cadastro&cod="&strcod&"&msg="&strmsg&""
		'response.redirect "../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod=524&busca="
		response.redirect "../funcionario/funcionarios_cadastro.asp?tipo=cadastro&cod="&strcod&"&func_ativos="&func_ativos&""
		response.write("TESTE MANO!!!!")
		end if
	
	
	'**********************************************************************
	'*** 					Exclui funcionario						    ***
	'**********************************************************************
	
	Elseif stracao = "excluir" Then
	    if protege_exclusao = 1 then
		xsql= "up_Funcionario_delete @cd_codigo='"&strcod&"'"
		dbconn.execute(xsql)
			'Apaga o histórico de cargos do funcionário
			xsql= "up_Funcionario_cargo_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga o histórico de unidades do funcionário
			xsql= "up_Funcionario_unidade_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga o histórico de status do funcionário
			xsql= "up_Funcionario_status_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga a lista de horarios do funcionário
			xsql= "up_Funcionario_horario_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga a lista de ausencias do funcionário
			xsql= "up_Funcionario_ausencias_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga a lista de contratos do funcionário
			xsql= "up_Funcionario_contratos_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga a lista de contatos do funcionário
			xsql= "up_contato_delete @cd_origem=4,@id_origem='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga a lista de dependentes do funcionário
			xsql= "up_dependentes_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga a escolaridade do funcionário
			xsql= "up_eescolaridade_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga a lista de empregos anteriores do funcionário
			xsql= "up_emprego_anterior_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga a função do funcionário
			xsql= "up_funcionario_funcao_delete @cd_funcionario='"&strcod&"'"
			dbconn.execute(xsql)
			'Apaga as ocorrencias do funcionário
			xsql= "up_ocorrencias_grupo_delete @cd_origem=4 AND id_origem='"&strcod&"'"
			dbconn.execute(xsql)
			
			strmsg = "funcionario excluido com sucesso!!!"
		end if
		
		'response.write("excluiiu")
		'response.redirect "../../empresa.asp?tipo=lista&msg="&strmsg&""%>
			<body onload="window.close(); window.opener.document.getElementById('form').submit(true);" >
		<%
	
	
	'******************************************************************
	'*** 					Cargo - tratamento					    ***
	'******************************************************************
	
	Elseif stracao = "cargos" Then
		'Insere cargo
		xsql = "up_funcionario_cargo_insert @cd_funcionario='"&strcod&"', @cd_cargo='"&str_cargo&"', @dt_atualizacao='"&cd_mes&"/"&cd_dia&"/"&cd_ano&"'"
		dbconn.execute(xsql)
		fecha_janela = 1
							
	Elseif stracao = "cargo_delete" then
		xsql= "up_funcionario_cargo_especifico_delete @cod_cargo="&cod_cargo&""
		dbconn.execute(xsql)
		fecha_janela = 1
	
	'******************************************************************
	'*** 				Aparelhos - tratamento					    ***
	'******************************************************************
	
	Elseif stracao = "aparelhos"  AND cd_aparelhos <> "" Then
		'Insere aparelho
		xsql = "up_funcionario_aparelhos_insert @cd_funcionario='"&strcod&"', @cd_aparelhos='"&str_aparelhos&"', @dt_atualizacao='"&cd_mes&"/"&cd_dia&"/"&cd_ano&"'"
		dbconn.execute(xsql)
		fecha_janela = 1
							
	Elseif stracao = "aparelhos_delete" then
		xsql= "up_funcionario_aparelhos_especifico_delete @cod_aparelhos="&cod_aparelhos&""
		dbconn.execute(xsql)
		fecha_janela = 1
	
	'******************************************************************
	'*** 					Funções - tratamento					***
	'******************************************************************
	
	Elseif stracao = "funcoes" Then
		'Insere função
		xsql = "up_funcionario_funcao_insert @cd_funcionario='"&strcod&"', @cd_funcao='"&str_funcao&"', @dt_atualizacao='"&cd_mes&"/"&cd_dia&"/"&cd_ano&"'"
		dbconn.execute(xsql)
		fecha_janela = 1
							
	Elseif stracao = "funcao_delete" then
		xsql= "up_funcionario_funcao_especifico_delete @cod_funcao="&cod_funcao&""
		dbconn.execute(xsql)
		'fecha_janela = 1
	
	'******************************************************************
	'*** 					Unidades - tratamento				    ***
	'******************************************************************
	
	ElseIf stracao = "unidade" Then
		'Insere dados
		xsql = "up_funcionario_unidade_insert @cd_funcionario='"&strcod&"', @cd_unidade='"&strcd_unidade&"', @dt_atualizacao='"&cd_mes&"/"&cd_dia&"/"&cd_ano&"'"
		dbconn.execute(xsql)
		fecha_janela = 1
							
	ElseIf stracao = "unidade_alterar" Then
		'altera dados
		xsql = "up_funcionario_unidade_update @cod_unidade='"&cod_unidade&"', @cd_unidade='"&strcd_unidade&"', @dt_atualizacao='"&cd_mes&"/"&cd_dia&"/"&cd_ano&"'"
		dbconn.execute(xsql)
		fecha_janela = 1
							
	Elseif stracao = "unidade_delete" then
		'apaga dados
		xsql= "up_funcionario_unidade_especifico_delete @cod_unidade="&cod_unidade&""
		dbconn.execute(xsql)
		fecha_janela = 1
	
	'**************************************************************
	'*** 					Novo Staus						    ***
	'**************************************************************
	
	ElseIf stracao = "status" Then
		'Insere dados
		'xsql = "up_funcionario_status_insert @cd_funcionario='"&strcod&"', @cd_status='"&str_status&"', @dt_atualizacao='"&month(dt_atualizacao)&"/"&day(dt_atualizacao)&"/"&year(dt_atualizacao)&"'"
		xsql = "up_funcionario_status_insert @cd_funcionario='"&strcod&"', @cd_status='"&str_status&"', @dt_atualizacao='"&cd_mes&"/"&cd_dia&"/"&cd_ano&"'"
		dbconn.execute(xsql)
		fecha_janela = 1
		'response.redirect("assistencia.asp?action="&action&"")
		response.redirect("../janelas/funcionarios_status.asp?cd_funcionario="&strcod&"&func_ativos="&func_ativos&"")
	
	'**************************************************************
	'*** 						Contrato					    ***
	'**************************************************************
	
	ElseIf stracao = "contrato" Then
	'	if cd_funcionario_contrato = "" Then
			'Insere dados
			response.write(dt_atualizacao)
			'xsql = "up_funcionario_contrato_insert @cd_funcionario='"&strcd_funcionario&"', @cd_contrato="&strregime_trabalho&", @dt_admissao="&Strdt_contratado&", @dt_atualizacao='"&dt_atualizacao&"'"
			xsql = "up_funcionario_contrato_insert @cd_funcionario='"&strcd_funcionario&"', @cd_contrato="&strregime_trabalho&", @dt_admissao="&Strdt_contratado&", @dt_atualizacao='"&dt_atualizacao&"'"
			dbconn.execute(xsql)
			fecha_janela = 1
				
				'Altera o campo dt_demissao_geral
				xsql = "UPDATE TBL_Funcionario SET dt_demissao_geral=null WHERE cd_codigo='"&strcd_funcionario&"'"
				dbconn.execute(xsql)
	'	else
	ElseIf stracao = "contrato_altera" Then
			'Altera dados
			xsql = "up_funcionario_contrato_altera @cd_codigo='"&cod_contrato&"', @cd_contrato='"&strregime_trabalho&"', @dt_admissao="&Strdt_contratado&", @dt_desligado="&Strdt_desligado&", @dt_atualizacao='"&dt_atualizacao&"'"
			dbconn.execute(xsql)
			fecha_janela = 1
			jan = 1
	'	end if
	
	Elseif stracao = "contrato_delete" Then
		xsql= "up_funcionario_contrato_especifico_delete @cod_contrato="&cod_contrato&""
		dbconn.execute(xsql)
		fecha_janela = 1
	
	'**************************************************************
	'*** 					Recisão	de contrato					***
	'**************************************************************
	
	ElseIf stracao = "recisao" Then
		nao_fecha = 1
		'Insere dados
		'response.write("@cd_contrato='"&strcod&"', @dt_recisao="&Strdt_recisao&", @dt_atualizacao='"&dt_atualizacao&"'")
		'xsql = "up_funcionario_contrato_recisao @cd_contrato='"&strcod&"', @dt_recisao="&Strdt_recisao&", @dt_atualizacao='"&dt_atualizacao&"'"
		xsql = "up_funcionario_contrato_recisao @cd_contrato='"&strcod&"', @dt_recisao="&Strdt_recisao&", @dt_atualizacao='"&dt_atualizacao&"'"
		dbconn.execute(xsql)
		
			'Altera o campo dt_demissao_geral
			xsql = "UPDATE TBL_Funcionario SET dt_demissao_geral="&Strdt_recisao&" WHERE cd_codigo='"&strcd_funcionario&"'"
			dbconn.execute(xsql)
			
			'Altera o status do Funcionario
			xsql = "up_funcionario_status_insert @cd_funcionario='"&strcd_funcionario&"', @cd_status=0, @dt_atualizacao='"&dt_atualizacao&"'"
			dbconn.execute(xsql)
			
			response.redirect("../janelas/funcionarios_contrato.asp?cd_funcionario="&strcod&"&func_ativos="&func_ativos&"")
				
		'response.redirect("assistencia.asp?action="&action&"")
	
	'**************************************************************
	'*** 					Novo Horário					    ***
	'**************************************************************
	
	ElseIf stracao = "horario" Then
		'Insere dados
		xsql = "up_funcionario_horario_insert @cd_funcionario='"&strcod&"', @hr_entrada='"&strhr_entrada&"', @nm_intervalo='"&strnm_intervalo&"', @hr_saida='"&strhr_saida&"', @dt_atualizacao='"&cd_mes&"/"&cd_dia&"/"&cd_ano&"', @cd_user="&cd_user&""
		dbconn.execute(xsql)
		fecha_janela = 1
		response.redirect("../funcionario/funcionarios_cadastro.asp?tipo=cadastro&cod="&strcod&"&func_ativos="&func_ativos)
	
	Elseif stracao = "horario_delete" Then
		xsql= "up_funcionario_horario_especifico_delete @cod_horario="&cod_horario&""
		dbconn.execute(xsql)
		fecha_janela = 1
		%>
		<body onload="window.close();">
		<%
	'cod_horario
	
	'**************************************************************
	'*** 					Novo Valor (padrão)				    ***
	'**************************************************************
	
	Elseif stracao = "valor_padrao" then
		'Insert
		xsql = "up_funcionario_valor_padrao_insert @cd_valor_tipo='"&str_valor_tipo&"', @nr_valor='"&str_valor&"', @dt_valor_atualizacao='"&month(str_valor_atualizacao)&"/"&day(str_valor_atualizacao)&"/"&year(str_valor_atualizacao)&"',  @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
		dbconn.execute(xsql)
		response.redirect("../../empresa.asp?tipo=valor_padrao&func_ativos="&func_ativos&"")
	
	Elseif stracao = "valor_padrao_altera" then
		'Update
		xsql = "up_funcionario_valor_padrao_update @cd_valor_tipo='"&str_valor_tipo&"', @cd_valor="&strcd_valor&", @nr_valor='"&str_valor&"', @dt_valor_atualizacao='"&month(str_valor_atualizacao)&"/"&day(str_valor_atualizacao)&"/"&year(str_valor_atualizacao)&"',  @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
		dbconn.execute(xsql)
		response.redirect("../../empresa.asp?tipo=valor_padrao&func_ativos="&func_ativos&"")
	
	Elseif stracao = "valor_padrao_especifico_delete"  Then
		'Delete
		xsql= "up_funcionario_valor_padrao_especifico_delete @cd_valor="&strcd_valor&""
		dbconn.execute(xsql)
		response.redirect("../../empresa.asp?tipo=valor_padrao&func_ativos="&func_ativos&"")
		
	'*******************************************************************
	'*** 					Férias					 			     ***
	'*******************************************************************
	Elseif stracao = "insert_ferias" then
		response.write("<br><br>@cd_funcionario='"&strcd_funcionario&"', @cd_empresa='"&strcd_empresa&"', @cd_unidade='"&strcd_unidade&"', @nm_cid='', @cd_motivo='6', @dt_afastamento_i="&str_periodo_i&", @dt_afastamento_f="&str_periodo_f&", @dt_pagamento_ferias="&dt_pagamento_ferias&", @nm_obs='"&nm_obs&"', @valor_pgto_ferias='"&valor_pgto_ferias&"', @cd_user='"&cd_user&"',@dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'")
		xsql = "sp_funcionario_ausencia_CRUD @funcao='IF', @cd_funcionario='"&strcd_funcionario&"', @cd_empresa='"&strcd_empresa&"', @cd_unidade='"&strcd_unidade&"', @nm_cid='', @cd_motivo='6', @dt_afastamento_i="&str_periodo_i&", @dt_afastamento_f="&str_periodo_f&", @dt_pagamento_ferias="&dt_pagamento_ferias&", @valor_pgto_ferias='"&valor_pgto_ferias&"', @nm_obs='"&nm_obs&"', @cd_user='"&cd_user&"',@dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
		dbconn.execute(xsql)
		nao_fecha = 1
		
		response.redirect("../../empresa/funcionario/ausencias/funcionarios_cad_ausencias.asp?cod='"&strcd_funcionario&"'&acao=insert_ferias")
	
	Elseif stracao = "update_ferias" then
	nao_fecha = 1
	response.write("@cd_ferias='"&strcd_ferias&"', @cd_funcionario='"&strcd_funcionario&"', @cd_empresa='"&strcd_empresa&"', @cd_unidade='"&strcd_unidade&"', @nm_cid='', @cd_motivo='6', @dt_afastamento_i="&str_periodo_i&", @dt_afastamento_f="&str_periodo_f&", @dt_pagamento_ferias="&dt_pagamento_ferias&", @valor_pgto_ferias='"&valor_pgto_ferias&"', @nm_obs='"&nm_obs&"', @cd_user='"&cd_user&"',@dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'")
	xsql = "sp_funcionario_ausencia_CRUD @funcao='UF', @cd_ferias='"&strcd_ferias&"', @cd_funcionario='"&strcd_funcionario&"', @cd_empresa='"&strcd_empresa&"', @cd_unidade='"&strcd_unidade&"', @nm_cid='', @cd_motivo='6', @dt_afastamento_i="&str_periodo_i&", @dt_afastamento_f="&str_periodo_f&", @dt_pagamento_ferias="&dt_pagamento_ferias&", @valor_pgto_ferias='"&valor_pgto_ferias&"', @nm_obs='"&nm_obs&"', @cd_user='"&cd_user&"',@dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
		dbconn.execute(xsql)
		response.redirect("../../empresa/funcionario/ausencias/funcionarios_cad_ausencias.asp?cod='"&strcd_funcionario&"'&acao=insert_ferias")
	
	Elseif stracao = "delete_ferias" then
	nao_fecha = 1
	response.write("@funcao='DF', @cd_ferias='"&strcd_ferias&"', cd_funcionarios='"&strcd_funcionario&"'")
	xsql = "sp_funcionario_ausencia_CRUD @funcao='DF', @cd_ferias='"&strcd_ferias&"'"
		dbconn.execute(xsql)
		response.redirect("../../empresa/funcionario/ausencias/funcionarios_cad_ausencias.asp?cod='"&strcd_funcionario&"'&acao=insert_ferias")
		
	'*******************************************************************
	'*** 					Suspende registros		 			     ***
	'*******************************************************************
	Elseif stracao = "suspende" then
		' suspende
		
		'xsql = "UPDATE TBL_funcionario_"&str_tbl&" SET cd_suspenso="&strcd_suspenso&", cd_cad_upd="&cd_user&", dt_cad_upd='"&month(now)&"/"&day(now)&"/"&year(now)&"' where cd_codigo="&strcod&"" '
		xsql = "UPDATE TBL_funcionario_"&str_tbl&" SET cd_suspenso="&strcd_suspenso&" where cd_codigo="&strcod&"" '
		dbconn.execute(xsql)
		response.redirect("../janelas/funcionarios_"&str_tbl&".asp?cd_funcionario="&strcd_funcionario&"&func_ativos="&func_ativos&"")
	
	'***************************************************************
	'*** 					Apaga Contato		 			     ***
	'***************************************************************
	Elseif stracao = "apaga_contato" Then
		xsql= "up_contato_especifico_delete @cd_contato="&cod_cont&""
		dbconn.execute(xsql)
		response.redirect("../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod="&strcod&"&func_ativos="&func_ativos&"")
	
	'*******************************************************************
	'*** 					Apaga Dependente		 			     ***
	'*******************************************************************
	Elseif stracao = "apaga_dependente" Then
		xsql= "up_dependente_especifico_delete @cd_dependente="&cod_dep&""
		dbconn.execute(xsql)
		response.redirect("../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod="&strcod&"&func_ativos="&func_ativos&"")
	
	'*******************************************************************
	'*** 					Apaga Escolaridade		 			     ***
	'*******************************************************************
	Elseif stracao = "apaga_escolaridade" Then
		xsql= "up_escolaridade_especifico_delete @cd_escolaridade="&cod_esc&""
		dbconn.execute(xsql)
		response.redirect("../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod="&strcod&"&func_ativos="&func_ativos&"")
	
	'*******************************************************************
	'*** 					Apaga Emprego			 			     ***
	'*******************************************************************
	Elseif stracao = "apaga_emprego" Then
		xsql= "up_emprego_especifico_delete @cd_emprego="&cod_emp&""
		dbconn.execute(xsql)
		response.redirect("../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod="&strcod&"&func_ativos="&func_ativos&"")
	
	'*******************************************************************
	'*** 					Apaga Valores			 			     ***
	'*******************************************************************
	Elseif stracao = "apaga_valor" Then
		xsql= "up_funcionario_valor_especifico_delete @cd_valor="&cod_valor&""
		dbconn.execute(xsql)
		response.redirect("../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod="&strcod&"&func_ativos="&func_ativos&"")
	
	'*******************************************************************
	else
	response.write("erro")
	End if
	
	
	'Set Upload = Nothing
%>
   
<%if nao_fecha <> 1 then%>
	<%if jan = 1 OR fecha_janela = 1 then%>   
		<body onload="window.close();">
	<%else
		response.redirect "../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod="&strcod&"&func_ativos="&func_ativos&""%>
		<body>
	<%end if%>
<%end if%>	
	<br><br>
	<%'response.write("up_funcionario_insert @cd_regime_trabalho='"&strregime_trabalho&"',@cd_matricula='"&strcd_matricula&"',@nm_nome='"&strnome&"',@nm_sobrenome='"&strnm_sobrenome&"', @nm_foto='"&strfoto&"', @dt_contratado="&Strdt_contratado&", @dt_nascimento="&strdt_nascimento&", @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"', @nm_pai='"&strpai&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_complemento='"&strnm_complemento&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"',@nm_cep='"&strnm_cep&"', @nm_bairro = '"&strbairro&"', @nm_nacionalidade='"&str_nacionalidade&"', @cd_pis='"&str_pis&"',@cd_ctps='"&str_ctps&"',@cd_ctps_serie='"&str_ctps_serie&"', @cd_assistencia_medica_matricula='"&str_assistencia_medica_matricula&"', @cd_assistencia_odonto_matricula='"&str_assistencia_odonto_matricula&"', @cd_vr='"&str_vr&"',@cd_vale_alimentacao='"&str_vale_alimentacao&"',@cd_sptrans='"&str_sptrans&"',@cd_cmt_bom='"&str_cmt_bom&"', @nm_banco='"&str_banco&"', @cd_banco_ag='"&str_banco_ag&"', @cd_banco_conta='"&str_banco_conta&"',@nm_tit_eleitor='"&nm_tit_eleitor&"',@nm_reservista='"&nm_reservista&"', @cd_conspro='"&str_conspro&"',@cd_numreg='"&str_numreg&"',@rgpro_concessao='"&str_rgpro_concessao&"',@cd_rgpro_status='"&str_rgpro_status&"', @obs_rgpro='"&str_rgpro&"',@dt_rgproinscr="&str_dt_rgproinscr&",@dt_rgpropag="&str_dt_rgpropag&",@dt_rgproval="&str_dt_rgproval&", @nm_endereco_alt='"&str_endereco_alt&"'")%>
<%end if%>
<br><br>
session_cd_usuario = <%=session_cd_usuario%><br>
<br>
jan = <%=jan%><br>
código: <%=strcod%> ou <%=cd_cod_funcionario%> ou <%=cd_funcionario_contrato%><br>
ação: <%=stracao%><br>
cod_contrato <%=cod_contrato%><br>
cargo: <%=cod_cargo%> - <%=str_cargo%><br>
atual: <%=dt_atualizacao%><br>
Status: <%=str_status%><br>
strcd_matricula: <%=strcd_matricula%><br>
strnome: <%=strnome%><br>
strnm_sobrenome: <%=strnm_sobrenome%><br>
Strdt_contratado: <%=Strdt_contratado%> - <%=Strdt_desligado%><br>
strregime_trabalho: <%=strregime_trabalho%><br>
strmsg: <%=strmsg%><br>
str_assistencia_medica_tipo: <%=strfoto%><br>
str_assistencia_medica_tipo: <%=str_assistencia_medica_tipo%><br>
str_banco: <%=str_banco%><br>
str_cd_banco_ag: <%=str_cd_banco_ag%><br>
str_cd_banco_conta: <%=str_cd_banco_conta%><br>
str_banco_conta_tipo: <%=str_banco_conta_tipo%><br>
status: <%=strstatus%><br>
Demissao: <%=Strdt_recisao%><br>
Hora: <%=dt_hora%><br>
Minuto: <%=dt_minuto%><br>
<br>
cod_unidade: <%=cod_unidade%><br>
unidade: <%=strcd_unidade%><br>
sigla: <%=nm_sigla%><br>
entrada: <%=strhr_entrada%><br>
Intervalo: <%=strnm_intervalo%><br>
Saida: <%=strhr_saida%><br>
TBL: <%=str_tbl%><br>
Suspende: <%=strcd_suspenso%><br>
ocorrencia: <%=ocorrencia%><br>
<br>
Endereço alt: <%=str_endereco_alt%><br>
<br>
:: Contato ::<br>
cd funcionario.: <%=id_contato_origem%><br>
Origem: <%=id_contato_origem%><br>
Nome: <%=nm_contato_nome%><br>
Cargo: <%=nm_contato_cargo%><br>
Categoria: <%=cd_contato_categoria%><br>
DDD: <%=cd_contato_ddd%><br>
Contato: <%=nm_contato_numero%><br>
Obs.: <%=nm_contato_obs%><br>
<br>
nm_dependente_nome: <%=nm_dependente_nome%><br>
cd_dependente_parentesco:<%=cd_dependente_parentesco%><br>
dt_dependente_nascimento: <%=dt_dependente_nascimento%><br>
nm_dependente_obs: <%=nm_dependente_obs%><br>
dt_atualizacao = <%=dt_atualizacao%><br><br>
str_escolaridade_total = <%=str_escolaridade_total%><br>
escolaridade_termino = <%=escolaridade_termino%><br>
cd_emprego = <%=cod_emp%><br><br>
--------------------------------------------------------------<br>
emprego_empresa = <%=emprego_empresa%><br>
emprego_funcao = <%=emprego_funcao%><br>
emprego_entrada = <%=emprego_entrada%><br>
emprego_saida = <%=emprego_saida%><br>
emprego_saida = <%=emprego_saida%><br>
emprego_obs<%=emprego_obs%><br><br>


						
str_valores_total = <%=str_valores_total%><br><br><br>


cd_valor_tipo=<%=str_valor_tipo%><br>
cd_valor=<%=strcd_valor%><br>
nr_valor=<%=str_valor%>
dt_valor_atualizacao=<%=month(str_valor_atualizacao)&"/"&day(str_valor_atualizacao)&"/"&year(str_valor_atualizacao)%><br>
nm_valor_obs=<%=str_valor_obs%><br>
cd_sessao=<%=session_cd_usuario%><br>
dt_atualizador=<%=month(now)&"/"&day(now)&"/"&year(now)%><br>
<br>
cd_transporte = <%=cd_transporte%><br>
<br>
cd_rgpro_cargo = <%=str_rgpro_cargo%>
<br><br>

cod_funcao = <%=cod_funcao%><br>
<br>
hr entrada - <%=strhr_entrada%><br>
hr saida - <%=strhr_saida%><br>
intervalo - <%=strnm_intervalo%><br>
<br>
----------------------------------<br>
contatos_total = <%=str_contatos_total%><br>
----------------------------------<br>
cd_aparelhos = <%=str_aparelhos%><br>
----------------------------------<br>



</body>