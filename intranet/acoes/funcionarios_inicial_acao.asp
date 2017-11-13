<%@ Language=VBScript %>
<!--#include file="../includes/inc_open_connection.asp"-->

<%
stracao = request("acao")

strregime_trabalho = request("cd_regime_trabalho")
strcd_matricula = request("cd_matricula")

strnome =  request("nm_nome")
str_sexo = request("cd_sexo")
strnm_endereco = request("nm_endereco")
strnr_numero =  request("nr_numero")
strnm_complemento = request("nm_complemento")
strnm_bairro = request("nm_bairro")
strnm_cep = request("nm_cep")
strnm_cidade =  request("nm_cidade")
strnm_estado =  request("nm_estado")
strdia_nascimento = request("dia_nascimento")
strmes_nascimento = request("mes_nascimento")
strano_nascimento = request("ano_nascimento")
	strdt_nascimento = strmes_nascimento&"/"&strdia_nascimento&"/"&strano_nascimento
strnm_nacionalidade = request("nm_nacionalidade")
strnm_naturalidade = request("nm_naturalidade")
strcd_estado_civil = request("cd_estado_civil")
strnm_conjuge = request("nm_conjuge")
strdt_conjuge_nasc = request("dt_conjuge_nasc")
	if strdt_conjuge_nasc <> "" Then
		strdt_conjuge_nasc = "'"&month(strdt_conjuge_nasc)&"/"&day(strdt_conjuge_nasc)&"/"&year(strdt_conjuge_nasc)&"'"
	else
		strdt_conjuge_nasc = "null"
	end if
strnm_endereco_alt = request("nm_endereco_alt")
strnm_mae = request("nm_mae")
strnm_pai = request("nm_pai")

strcd_contato_categoria1 = request("cd_contato_categoria1")
strnm_contato_ddd1 = request("cd_contato_ddd1")
strnm_contato_numero1 = request("nm_contato_numero1")
strnm_contato_obs1 = request("nm_contato_obs1")
	nm_contato1 = strcd_contato_categoria1&"!"&strnm_contato_ddd1&"#"&strnm_contato_numero1&"$"&strnm_contato_obs1&"*"

strcd_contato_categoria2 = request("cd_contato_categoria2")
strnm_contato_ddd2 = request("cd_contato_ddd2")
strnm_contato_numero2 = request("nm_contato_numero2")
strnm_contato_obs2 = request("nm_contato_obs2")
	nm_contato2 = strcd_contato_categoria2&"!"&strnm_contato_ddd2&"#"&strnm_contato_numero2&"$"&strnm_contato_obs2&"*"
	
strcd_contato_categoria3 = request("cd_contato_categoria3")
strnm_contato_email3 = request("cd_contato_email3")
strnm_contato_obs3 = request("nm_contato_obs3")
	nm_contato3 = strcd_contato_categoria3&"!"&strnm_contato_email3&"#"&strnm_contato_obs3&"$"

strcd_contato_categoria4 = request("cd_contato_categoria4")
strnm_contato_ddd4 = request("cd_contato_ddd4")
strnm_contato_numero4 = request("nm_contato_numero4")
strnm_contato_obs4 = request("nm_contato_obs4")
strnm_contato_nome4 = request("nm_contato_nome4")
strnm_contato_cargo4 = request("nm_contato_cargo4")
	nm_contato4 = strcd_contato_categoria4&"!"&strnm_contato_ddd4&"#"&strnm_contato_numero4&"$"&strnm_contato_obs4&"*"&strnm_contato_nome4&"%"&strnm_contato_cargo4&"_"

strcd_pis = request("cd_pis")
strcd_ctps = request("cd_ctps")
strcd_ctps_serie = request("cd_ctps_serie")

strnm_rg =  request("nm_rg")
strdt_rg_emissao =  request("dt_rg_emissao")
	barra_ver_rg = instr(1,strdt_rg_emissao,"/",1)
	if barra_ver_rg > 0 then 
		strdt_rg_emissao = "'"&month(strdt_rg_emissao)&"/"&day(strdt_rg_emissao)&"/"&year(strdt_rg_emissao)&"'"
	else
		strdt_rg_emissao = "null"
	end if
strnm_rg_expedidor =  request("nm_rg_expedidor")
strnm_cpf =  request("nm_cpf")

strnm_tit_eleitor = request("nm_tit_eleitor")
strnr_tit_eleitor_zona = request("nr_tit_eleitor_zona")
strnr_tit_eleitor_seccao = request("nr_tit_eleitor_seccao")

strnm_reservista = request("nm_reservista")
strnm_reservista_cat = request("nm_reservista_cat")

strcd_conspro = request("cd_conspro")
strcd_numreg = request("cd_numreg")
str_rgpro_concessao = request("rgpro_concessao")
strrgpro_status = request("rgpro_status")
strobs_rgpro = request("obs_rgpro")
strdt_rgproinscr = request("dt_rgproinscr")
	if strdt_rgproinscr <> "" Then
		strdt_rgproinscr = "'"&month(strdt_rgproinscr)&"/"&day(strdt_rgproinscr)&"/"&year(strdt_rgproinscr)&"'"
	else
		strdt_rgproinscr = "null"
	end if
strdt_rgpropag = request("dt_rgpropag")
	if strdt_rgpropag <> "" then
	strdt_rgpropag = "'"&month(strdt_rgpropag)&"/"&day(strdt_rgpropag)&"/"&year(strdt_rgpropag)&"'"
	else
		strdt_rgpropag = "null"
	end if
strdt_rgproval = request("dt_rgproval")
	if strdt_rgproval <> "" Then
		strdt_rgproval = "'"&month(strdt_rgproval)&"/"&day(strdt_rgproval)&"/"&year(strdt_rgproval)&"'"
	else
		strdt_rgproval = "null"
	end if

strnr_hepatite_b = request("nr_hepatite_b")
strnr_dupla_adulto = request("nr_dupla_adulto")
strdt_dupla_adulto_validade = request("dt_dupla_adulto_validade")
	if strdt_dupla_adulto_validade <> "" Then
		strdt_dupla_adulto_validade = "'"&month(strdt_dupla_adulto_validade)&"/"&day(strdt_dupla_adulto_validade)&"/"&year(strdt_dupla_adulto_validade)&"'"
	else
		strdt_dupla_adulto_validade = "null"
	end if 
strnr_scr = request("nr_scr")
strdt_exame_medico = request("dt_exame_medico")
	if strdt_exame_medico <> "" Then
		strdt_exame_medico = "'"&month(strdt_exame_medico)&"/"&day(strdt_exame_medico)&"/"&year(strdt_exame_medico)&"'"
	else
		strdt_exame_medico = "null"
	end if

strnr_caract_altura = request("nr_caract_altura")
strnr_caract_peso = request("nr_caract_peso")
strnr_caract_manequim = request("nr_caract_manequim")
strnr_caract_calcado = request("nr_caract_calcado")

strnm_dependente_nome1 = request("nm_dependente_nome1")
strcd_dependente_sexo1 = request("cd_dependente_sexo1")
strdt_dependente_nascimento1 = request("dt_dependente_nascimento1")
	if strdt_dependente_nascimento1 <> "" Then
		strdt_dependente_nascimento1 = month(strdt_dependente_nascimento1)&"/"&day(strdt_dependente_nascimento1)&"/"&year(strdt_dependente_nascimento1)
	'else
		'strdt_dependente_nascimento1 = "null"
	end if
strcd_dependente_obs1 = request("cd_dependente_obs1")
	strnm_dependente1 = strnm_dependente_nome1&"!"&strcd_dependente_sexo1&"#"&strdt_dependente_nascimento1&"$"&strcd_dependente_obs1

strnm_dependente_nome2 = request("nm_dependente_nome2")
strcd_dependente_sexo2 = request("cd_dependente_sexo2")
strdt_dependente_nascimento2 = request("dt_dependente_nascimento2")
	if strdt_dependente_nascimento2 <> "" Then
		strdt_dependente_nascimento2 = month(strdt_dependente_nascimento2)&"/"&day(strdt_dependente_nascimento2)&"/"&year(strdt_dependente_nascimento2)
	'else
		'strdt_dependente_nascimento2 = "null"
	end if
strcd_dependente_obs2 = request("cd_dependente_obs2")
	strnm_dependente2 = strnm_dependente_nome2&"!"&strcd_dependente_sexo2&"#"&strdt_dependente_nascimento2&"$"&strcd_dependente_obs2

strnm_dependente_nome3 = request("nm_dependente_nome3")
strcd_dependente_sexo3 = request("cd_dependente_sexo3")
strdt_dependente_nascimento3 = request("dt_dependente_nascimento3")
	if strdt_dependente_nascimento3 <> "" Then
		strdt_dependente_nascimento3 = month(strdt_dependente_nascimento3)&"/"&day(strdt_dependente_nascimento3)&"/"&year(strdt_dependente_nascimento3)
	'else
		'strdt_dependente_nascimento3 = "null"
	end if
strcd_dependente_obs3 = request("cd_dependente_obs3")
	strnm_dependente3 = strnm_dependente_nome3&"!"&strcd_dependente_sexo3&"#"&strdt_dependente_nascimento3&"$"&strcd_dependente_obs3

strnm_dependente_nome4 = request("nm_dependente_nome4")
strcd_dependente_sexo4 = request("cd_dependente_sexo4")
strdt_dependente_nascimento4 = request("dt_dependente_nascimento4")
	if strdt_dependente_nascimento4 <> "" Then
		strdt_dependente_nascimento4 = month(strdt_dependente_nascimento4)&"/"&day(strdt_dependente_nascimento4)&"/"&year(strdt_dependente_nascimento4)
	'else
		'strdt_dependente_nascimento4 = "null"
	end if
strcd_dependente_obs4 = request("cd_dependente_obs4")
	strnm_dependente4 = strnm_dependente_nome4&"!"&strcd_dependente_sexo4&"#"&strdt_dependente_nascimento4&"$"&strcd_dependente_obs4

strcd_ensino_grau1 = request("cd_ensino_grau1")
strnm_instituicao1 = request("nm_instituicao1")
strnm_curso1 = request("nm_curso1")
strcd_ensino_andamento1 = request("cd_ensino_andamento1")
strdt_termino1 = request("dt_termino1")
	if strdt_termino1 <> "" Then
		strdt_termino1 = month(strdt_termino1)&"/"&day(strdt_termino1)&"/"&year(strdt_termino1)
	end if
strnm_obs1 = request("nm_obs1")
	strnm_ensino1 = strcd_ensino_grau1&"!"&strnm_instituicao1&"#"&strnm_curso1&"$"&strcd_ensino_andamento1&"*"&strdt_termino1&"%"&strnm_obs1

strcd_ensino_grau2 = request("cd_ensino_grau2")
strnm_instituicao2 = request("nm_instituicao2")
strnm_curso2 = request("nm_curso2")
strcd_ensino_andamento2 = request("cd_ensino_andamento2")
strdt_termino2 = request("dt_termino2")
	if strdt_termino2 <> "" Then
		strdt_termino2 = month(strdt_termino2)&"/"&day(strdt_termino2)&"/"&year(strdt_termino2)
	end if
strnm_obs2 = request("nm_obs2")
	strnm_ensino2 = strcd_ensino_grau2&"!"&strnm_instituicao2&"#"&strnm_curso2&"$"&strcd_ensino_andamento2&"*"&strdt_termino2&"%"&strnm_obs2

strcd_ensino_grau3 = request("cd_ensino_grau3")
strnm_instituicao3 = request("nm_instituicao3")
strnm_curso3 = request("nm_curso3")
strcd_ensino_andamento3 = request("cd_ensino_andamento3")
strdt_termino3 = request("dt_termino3")
	if strdt_termino3 <> "" Then
		strdt_termino3 = month(strdt_termino3)&"/"&day(strdt_termino3)&"/"&year(strdt_termino3)
	end if
strnm_obs3 = request("nm_obs3")
	strnm_ensino3 = strcd_ensino_grau3&"!"&strnm_instituicao3&"#"&strnm_curso3&"$"&strcd_ensino_andamento3&"*"&strdt_termino3&"%"&strnm_obs3

strcd_ensino_grau4 = request("cd_ensino_grau4")
strnm_instituicao4 = request("nm_instituicao4")
strnm_curso4 = request("nm_curso4")
strcd_ensino_andamento4 = request("cd_ensino_andamento4")
strdt_termino4 = request("dt_termino4")
	if strdt_termino4 <> "" Then
		strdt_termino4 = month(strdt_termino4)&"/"&day(strdt_termino4)&"/"&year(strdt_termino4)
	end if
strnm_obs4 = request("nm_obs4")
	strnm_ensino4 = strcd_ensino_grau4&"!"&strnm_instituicao4&"#"&strnm_curso4&"$"&strcd_ensino_andamento4&"*"&strdt_termino4&"%"&strnm_obs4

'*** Empregos anteriores
nm_emprego_empresa1 = request("nm_emprego_empresa1")
strnm_emprego_cargo1 = request("nm_emprego_cargo1")
strdt_emprego_entrada1 = request("dt_emprego_entrada1")
	if strdt_emprego_entrada1 <> "" Then
		strdt_emprego_entrada1 = month(strdt_emprego_entrada1)&"/"&day(strdt_emprego_entrada1)&"/"&year(strdt_emprego_entrada1)
	end if
strdt_emprego_saida1 = request("dt_emprego_saida1")
	if strdt_emprego_saida1 <> "" Then
		strdt_emprego_saida1 = month(strdt_emprego_saida1)&"/"&day(strdt_emprego_saida1)&"/"&year(strdt_emprego_saida1)
	end if
strdt_emprego_obs1 = request("dt_emprego_obs1")
	strnm_emprego1 = nm_emprego_empresa1&"!"&strnm_emprego_cargo1&"#"&strdt_emprego_entrada1&"$"&strdt_emprego_saida1&"*"&strdt_emprego_obs1
	
nm_emprego_empresa2 = request("nm_emprego_empresa2")
strnm_emprego_cargo2 = request("nm_emprego_cargo2")
strdt_emprego_entrada2 = request("dt_emprego_entrada2")
	if strdt_emprego_entrada2 <> "" Then
		strdt_emprego_entrada2 = month(strdt_emprego_entrada2)&"/"&day(strdt_emprego_entrada2)&"/"&year(strdt_emprego_entrada2)
	end if
strdt_emprego_saida2 = request("dt_emprego_saida2")
	if strdt_emprego_saida2 <> "" Then
		strdt_emprego_saida2 = month(strdt_emprego_saida2)&"/"&day(strdt_emprego_saida2)&"/"&year(strdt_emprego_saida2)
	end if
strdt_emprego_obs2 = request("dt_emprego_obs2")
	strnm_emprego2 = nm_emprego_empresa2&"!"&strnm_emprego_cargo2&"#"&strdt_emprego_entrada2&"$"&strdt_emprego_saida2&"*"&strdt_emprego_obs2

nm_emprego_empresa3 = request("nm_emprego_empresa3")
strnm_emprego_cargo3 = request("nm_emprego_cargo3")
strdt_emprego_entrada3 = request("dt_emprego_entrada3")
	if strdt_emprego_entrada3 <> "" Then
		strdt_emprego_entrada3 = month(strdt_emprego_entrada3)&"/"&day(strdt_emprego_entrada3)&"/"&year(strdt_emprego_entrada3)
	end if
strdt_emprego_saida3 = request("dt_emprego_saida3")
	if strdt_emprego_saida3 <> "" Then
		strdt_emprego_saida3 = month(strdt_emprego_saida3)&"/"&day(strdt_emprego_saida3)&"/"&year(strdt_emprego_saida3)
	end if
strdt_emprego_obs3 = request("dt_emprego_obs3")
	strnm_emprego3 = nm_emprego_empresa3&"!"&strnm_emprego_cargo3&"#"&strdt_emprego_entrada3&"$"&strdt_emprego_saida3&"*"&strdt_emprego_obs3

nm_emprego_empresa4 = request("nm_emprego_empresa4")
strnm_emprego_cargo4 = request("nm_emprego_cargo4")
strdt_emprego_entrada4 = request("dt_emprego_entrada4")
	if strdt_emprego_entrada4 <> "" Then
		strdt_emprego_entrada4 = month(strdt_emprego_entrada4)&"/"&day(strdt_emprego_entrada3)&"/"&year(strdt_emprego_entrada3)
	end if
strdt_emprego_saida4 = request("dt_emprego_saida4")
	if strdt_emprego_saida4 <> "" Then
		strdt_emprego_saida4 = month(strdt_emprego_saida4)&"/"&day(strdt_emprego_saida4)&"/"&year(strdt_emprego_saida4)
	end if
strdt_emprego_obs4 = request("dt_emprego_obs4")
	strnm_emprego4 = nm_emprego_empresa4&"!"&strnm_emprego_cargo4&"#"&strdt_emprego_entrada4&"$"&strdt_emprego_saida4&"*"&strdt_emprego_obs4
	
strnm_parente = request("nm_parente")
strcd_parente_parentesco = request("cd_parente_parentesco")
	if strcd_parente_parentesco = "" Then strcd_parente_parentesco = "null"

strcd_residencia_tipo = request("cd_residencia_tipo")
	if strcd_residencia_tipo =  "" then strcd_residencia_tipo = "null"
strcd_residencia_financ = request("cd_residencia_financ")
	if strcd_residencia_financ =  "" then strcd_residencia_financ = "null"
strcd_residencia_reside = request("cd_residencia_reside")
	if strcd_residencia_reside =  "" then strcd_residencia_reside = "null"
strcd_veiculo_tipo = request("cd_veiculo_tipo")
	if strcd_veiculo_tipo =  "" then strcd_veiculo_tipo = "null"
strcd_veiculo_financ = request("cd_veiculo_financ")
	if strcd_veiculo_financ =  "" then strcd_veiculo_financ = "null"

cd_dia = request("cd_dia")
cd_mes = request("cd_mes")
cd_ano = request("cd_ano")
	cd_hora = hour(now)
	cd_minuto = minute(now)
	cd_segundo = second(now)

	dt_desliga = cd_mes&"/"&cd_dia&"/"&cd_ano
	dt_atualizacao = cd_mes&"/"&cd_dia&"/"&cd_ano&" "&cd_hora&":"&cd_minuto&":"&cd_segundo
	
	'dt_desliga = cd_dia&"/"&cd_mes&"/"&cd_ano
	'dt_desliga = "12/09/2008"

dt_ano = YEAR(Now)
dt_mes = Month(Now)
dt_dia = Day(now)
dt_hora = hour(now)
dt_minuto = minute(now)

str_banco = request("nm_banco")
str_banco_ag = request("cd_banco_ag")
str_banco_conta = request("cd_banco_conta")
str_banco_conta_tipo = request("cd_banco_conta_tipo")





'**********************************************************************
'*** 					Insere funcionario						    ***
'**********************************************************************
If stracao = "insert" Then
		xsql = "up_Funcionario_inicial_insert @nm_nome='"&strnome&"', @cd_sexo='"&str_sexo&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"',@nm_complemento='"&strnm_complemento&"', @nm_bairro='"&strnm_bairro&"', @nm_cep='"&strnm_cep&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"', @dt_nascimento='"&strdt_nascimento&"',@nm_nacionalidade='"&strnm_nacionalidade&"',@nm_naturalidade='"&strnm_naturalidade&"',@cd_estado_civil='"&strcd_estado_civil&"',@nm_conjuge='"&strnm_conjuge&"', @dt_conjuge_nasc="&strdt_conjuge_nasc&",@nm_endereco_alt='"&strnm_endereco_alt&"',@nm_mae = '"&strnm_mae&"', @nm_pai='"&strnm_pai&"', @nm_contato1='"&nm_contato1&"', @nm_contato2='"&nm_contato2&"', @nm_contato3='"&nm_contato3&"', @nm_contato4='"&nm_contato4&"',	@cd_pis='"&strcd_pis&"', @cd_ctps='"&strcd_ctps&"', @cd_ctps_serie='"&strcd_ctps_serie&"',	@nm_rg='"&strnm_rg&"',@dt_rg_emissao="&strdt_rg_emissao&", @nm_rg_expedidor='"&strnm_rg_expedidor&"', @nm_cpf='"&strnm_cpf&"', @nm_tit_eleitor='"&strnm_tit_eleitor&"', @nm_tit_eleitor_zona='"&strnr_tit_eleitor_zona&"', @nm_tit_eleitor_seccao='"&strnr_tit_eleitor_seccao&"', @nm_reservista='"&strnm_reservista&"', @cd_conspro='"&strcd_conspro&"', @cd_numreg='"&strcd_conspro&"', @rgpro_concessao='"&str_rgpro_concessao&"', @cd_rgpro_status='"&strrgpro_status&"', @obs_rgpro='"&strobs_rgpro&"', @dt_rgpro_inscr="&strdt_rgproinscr&", @dt_rgpropag="&strdt_rgpropag&", @dt_rgproval="&strdt_rgproval&", @nr_hepatite_b='"&strnr_hepatite_b&"', @nr_dupla_adulto='"&strnr_dupla_adulto&"', @dt_dupla_adulto_validade="&strdt_dupla_adulto_validade&", @nr_scr='"&strnr_scr&"', @dt_exame_medico="&strdt_exame_medico&",@nr_caract_altura='"&strnr_caract_altura&"',@nr_caract_peso='"&strnr_caract_peso&"',@nr_caract_manequim='"&strnr_caract_manequim&"', @nr_caract_calcado='"&strnr_caract_calcado&"', @nm_dependente1='"&strnm_dependente1&"', @nm_dependente2='"&strnm_dependente2&"', @nm_dependente3='"&strnm_dependente3&"', @nm_dependente4='"&strnm_dependente4&"',@nm_ensino1='"&strnm_ensino1&"',@nm_ensino2='"&strnm_ensino2&"',@nm_ensino3='"&strnm_ensino3&"',@nm_ensino4='"&strnm_ensino4&"', @nm_emprego1='"&strnm_emprego1&"', @nm_emprego2='"&strnm_emprego2&"', @nm_emprego3='"&strnm_emprego3&"', @nm_emprego4='"&strnm_emprego4&"', @nm_parente='"&strnm_parente&"', @cd_parente_parentesco='"&strcd_parente_parentesco&"', @cd_residencia_tipo='"&strcd_residencia_tipo&"', @cd_residencia_financ='"&cd_residencia_financ&"', @cd_residencia_reside='"&strcd_residencia_reside&"', @cd_veiculo_tipo='"&strcd_veiculo_tipo&"', @cd_veiculo_financ='"&strcd_veiculo_financ&"'"

	dbconn.execute(xsql)

'**********************************************************************
'*** 					Altera funcionario						    ***
'**********************************************************************





Else
response.write("erro")
End if

'Set Upload = Nothing 
'response.redirect "../empresa.asp?tipo=lista&msg="&strmsg&""
%>

   

   
<body onload="window.closes()">
<!--body-->
<!--br-->

@nm_nome=<%=strnome%><br> 
@cd_sexo=<%=str_sexo%><br> 
@nm_endereco=<%=strnm_endereco%><br> 
@nr_numero=<%=strnr_numero%><br>
@nm_complemento=<%=strnm_complemento%><br> 
@nm_bairro=<%=strnm_bairro%><br>
@nm_cep=<%=strnm_cep%><br> 
@nm_cidade=<%=strnm_cidade%><br> 
@nm_estado=<%=strnm_estado%><br>
@dt_nascimento=<%=strdt_nascimento%><br>
@nm_nacionalidade=<%=strnm_nacionalidade%><br>
@nm_naturalidade=<%=strnm_naturalidade%><br>
@cd_estado_civil=<%=strcd_estado_civil%><br>
@nm_conjuge=<%=strnm_conjuge%><br>
@dt_conjuge_nasc=<%=strdt_conjuge_nasc%><br>
@nm_endereco_alt=<%=strnm_endereco_alt%><br>
@nm_mae = <%=strnm_mae%><br>
@nm_pai=<%=strnm_pai%><br> 
@nm_contato1=<%=nm_contato1%><br> 
@nm_contato2=<%=nm_contato2%><br> 
@nm_contato3=<%=nm_contato3%><br> 
@nm_contato4=<%=nm_contato4%><br>
@cd_pis=<%=strcd_pis%><br>
@cd_ctps=<%=strcd_ctps%><br>
@cd_ctps_serie=<%=strcd_ctps_serie%><br>
@nm_rg=<%=strnm_rg%><br>
@dt_rg_emissao=<%=strdt_rg_emissao%><br>
@nm_rg_expedidor=<%=strnm_rg_expedidor%><br>
@nm_cpf=<%=strnm_cpf%><br>
@nm_tit_eleitor=<%=strnm_tit_eleitor%><br> 
@nm_tit_eleitor_zona=<%=strnr_tit_eleitor_zona%><br> 
@nm_tit_eleitor_seccao=<%=strnr_tit_eleitor_seccao%><br>
@nm_reservista=<%=strnm_reservista%><br>
@cd_conspro=<%=strcd_conspro%><br>
@cd_numreg=<%=strcd_conspro%><br>
@rgpro_concessao=<%=str_rgpro_concessao%><br>
@cd_rgpro_status=<%=strrgpro_status%><br>
@obs_rgpro=<%=strobs_rgpro%>, <br>
@dt_rgpro_inscr=<%=strdt_rgproinscr%> <br>
@dt_rgpropag=<%=strdt_rgpropag%><br>
@dt_rgproval=<%=strdt_rgproval%><br>
@nr_hepatite_b=<%=strnr_hepatite_b%><br> 
@nr_dupla_adulto=<%=strnr_dupla_adulto%><br> 
@dt_dupla_adulto_validade=<%=strdt_dupla_adulto_validade%><br>
@nr_scr=<%=strnr_scr%><br>
@dt_exame_medico=<%=strdt_exame_medico%><br>
@nr_caract_altura=<%=strnr_caract_altura%><br>
@nr_caract_peso=<%=strnr_caract_peso%><br>
@nr_caract_manequim=<%=strnr_caract_manequim%><br>
@nr_caract_calcado=<%=strnr_caract_calcado%><br>
@nm_dependente1=<%=strnm_dependente1%><br>
@nm_dependente2=<%=strnm_dependente2%><br>
@nm_dependente3=<%=strnm_dependente3%><br>
@nm_dependente4=<%=strnm_dependente4%><br>
@nm_ensino1=<%=strnm_ensino1%>,<br>
@nm_ensino2=<%=strnm_ensino2%>,<br>
@nm_ensino3=<%=strnm_ensino3%>,<br>
@nm_ensino4=<%=strnm_ensino4%>, <br>
@nm_emprego1=<%=strnm_emprego1%>, <br>
@nm_emprego2=<%=strnm_emprego2%>, <br>
@nm_emprego3=<%=strnm_emprego3%>, <br>
@nm_emprego4=<%=strnm_emprego4%>, <br>
@nm_parente=<%=strnm_parente%>, <br>
@cd_parente_parentesco=<%=strcd_parente_parentesco%>, <br>
@cd_residencia_tipo=<%=strcd_residencia_tipo%>, <br>
@cd_residencia_financ=<%=cd_residencia_financ%>, <br>
@cd_residencia_reside=<%=strcd_residencia_reside%>, <br>
@cd_veiculo_tipo=<%=strcd_veiculo_tipo%>, <br>
@cd_veiculo_financ=<%=strcd_veiculo_financ%>
<br>
<br>

<!-- -->
Funcionário Cadastrado com sucesso!!
</body>