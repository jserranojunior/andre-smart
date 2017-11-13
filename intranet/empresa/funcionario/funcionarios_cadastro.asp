<!--script language="JavaScript" type="text/javascript" src="../js/richtext.js"></script-->
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->

<%
session_cd_usuario = session("cd_codigo")
cd_user = session_cd_usuario
usuario_restrito = int("65")

strcod = request("cod")
strcod_ini = request("cod_ini")

strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")


If strcod <> "" And strmsg = "" Then
	'**** Abre dados do cadastro oficial ***
	xsql ="up_Funcionario_Lista_por_cd_codigo @cd_codigo="&strcod&""
	Set rs = dbconn.execute(xsql)
		If NOT rs.EOF Then
		str_funcionario = rs("cd_codigo")
		strmatricula = rs("cd_matricula") 
		strnome = rs("nm_nome")
		strfoto = rs("nm_foto")
		strdt_nascimento = rs("dt_nascimento")
			strdia_nascimento = zero(day(strdt_nascimento))
			strmes_nascimento = zero(month(strdt_nascimento))
			strano_nascimento = year(strdt_nascimento)
		strnm_rg = rs("nm_rg")
		strdt_rg_emissao = rs("dt_rg_emissao")
			if strdt_rg_emissao <> "" Then
				strdt_rg_emissao_dia = zero(day(strdt_rg_emissao))
				strdt_rg_emissao_mes = zero(month(strdt_rg_emissao))
				strdt_rg_emissao_ano = year(strdt_rg_emissao)
			end if
		strnm_rg_expedidor = rs("nm_rg_expedidor")
		strnm_cpf = rs("nm_cpf")
		strnm_tit_eleitor = rs("nm_tit_eleitor")
		strnm_reservista = rs("nm_reservista")
		strmae = rs("nm_mae")
		strpai = rs("nm_pai")
		strnm_endereco = rs("nm_endereco")
		strnr_numero = rs("nr_numero")
		strnm_complemento = rs("nm_complemento")
		strnm_cidade = rs("nm_cidade")
		strnm_estado = rs("nm_estado")
		strnm_cep = rs("nm_cep")
		strbairro = rs("nm_bairro")
		str_nacionalidade = rs("nm_nacionalidade")
		str_pis = rs("cd_pis")
		str_ctps = rs("cd_ctps")
		str_ctps_serie = rs("cd_ctps_serie")
		str_estado_civil = rs("cd_estado_civil")
		
		str_assistencia_medica_matricula = rs("cd_assistencia_medica_matricula")
		str_assistencia_medica_coparticipacao = rs("nr_assistencia_medica_coparticipacao")
		str_assistencia_odonto_matricula = rs("cd_assistencia_odonto_matricula")
		str_assistencia_odonto_coparticipacao = rs("nr_assistencia_odonto_coparticipacao")
		str_vr = rs("cd_vr")
		str_vale_alimentacao = rs("cd_vale_alimentacao")
		str_cmt_bom = rs("cd_cmt_bom")
		str_sptrans = rs("cd_sptrans")
		
		str_banco = rs("nm_banco")
		str_banco_ag = rs("cd_banco_ag")
		str_banco_conta = rs("cd_banco_conta")
		
		cd_numreg = rs("cd_numreg")
		str_estado_civil = rs("cd_estado_civil")
		str_conjuge = rs("nm_conjuge")
		str_conjuge_nasc = rs("dt_conjuge_nasc")
			if str_conjuge_nasc <> "" Then
				str_conjuge_nasc_dia = zero(day(str_conjuge_nasc))
				str_conjuge_nasc_mes = zero(month(str_conjuge_nasc))
				str_conjuge_nasc_ano = year(str_conjuge_nasc)
			else
				str_conjuge_nasc_dia = ""
				str_conjuge_nasc_mes = ""
				str_conjuge_nasc_ano = ""
			end if
			
		str_cd_plano_saude_conj = rs("cd_plano_saude_conj")
		str_naturalidade = rs("nm_naturalidade")
		strnr_tit_eleitor_zona = rs("nr_tit_eleitor_zona")
		strnr_tit_eleitor_seccao = rs("nr_tit_eleitor_seccao")
		strnm_reservista_cat = rs("nm_reservista_cat")
		strnr_hepatite_b = rs("nr_hepatite_b")
		strnr_dupla_adulto = rs("nr_dupla_adulto")
		strdt_dupla_adulto_validade = rs("dt_dupla_adulto_validade")
			if strdt_dupla_adulto_validade <> "" Then
				strdt_dupla_adulto_validade_dia = zero(day(strdt_dupla_adulto_validade))
				strdt_dupla_adulto_validade_mes = zero(month(strdt_dupla_adulto_validade))
				strdt_dupla_adulto_validade_ano = year(strdt_dupla_adulto_validade)
			else
				strdt_dupla_adulto_validade_dia = ""
				strdt_dupla_adulto_validade_mes = ""
				strdt_dupla_adulto_validade_ano = ""
			end if
		strnr_scr = rs("nr_scr")
		strdt_exame_medico = rs("dt_exame_medico")
			if strdt_exame_medico <> "" Then
				strdt_exame_medico_dia = zero(day(strdt_exame_medico))
				strdt_exame_medico_mes = zero(month(strdt_exame_medico))
				strdt_exame_medico_ano = year(strdt_exame_medico)
			else
				strdt_exame_medico_dia = ""
				strdt_exame_medico_mes = ""
				strdt_exame_medico_ano = ""
			end if
		strnr_caract_altura = rs("nr_caract_altura")
		strnr_caract_peso = rs("nr_caract_peso")
		strnr_caract_manequim = rs("nr_caract_manequim")
		strnr_caract_calcado = rs("nr_caract_calcado")
		strnm_parente = rs("nm_parente")
		strcd_parente_parentesco = rs("cd_parente_parentesco")
		strcd_residencia_tipo = rs("cd_residencia_tipo")
		strcd_residencia_financ = rs("cd_residencia_financ")
		strcd_residencia_reside = rs("cd_residencia_reside")
		strcd_veiculo_tipo = rs("cd_veiculo_tipo")
		strcd_veiculo_financ = rs("cd_veiculo_financ")
		
		str_conspro = rs("cd_conspro")
		str_rgpro_cargo = rs("cd_rgpro_cargo")
		str_numreg = rs("cd_numreg")
		str_rgpro_concessao = rs("rgpro_concessao")
		str_rgpro_status = rs("cd_rgpro_status")
		str_rgpro = rs("obs_rgpro")
		str_dt_rgproinscr = rs("dt_rgproinscr")
			if str_dt_rgproinscr <> "" Then
				str_dt_rgproinscr_dia = zero(day(str_dt_rgproinscr))
				str_dt_rgproinscr_mes = zero(month(str_dt_rgproinscr))
				str_dt_rgproinscr_ano = year(str_dt_rgproinscr)
			else
				str_dt_rgproinscr_dia = ""
				str_dt_rgproinscr_mes = ""
				str_dt_rgproinscr_ano = ""
			end if
		str_dt_rgpropag = rs("dt_rgpropag")
			if str_dt_rgpropag <> "" Then
				str_dt_rgpropag_dia = zero(day(str_dt_rgpropag))
				str_dt_rgpropag_mes = zero(month(str_dt_rgpropag))
				str_dt_rgpropag_ano = year(str_dt_rgpropag)
			else
				str_dt_rgpropag_dia = ""
				str_dt_rgpropag_mes = ""
				str_dt_rgpropag_ano = ""
			end if
		str_dt_rgproval = rs("dt_rgproval")
			if str_dt_rgproval <> "" Then
				str_dt_rgproval_dia = zero(day(str_dt_rgproval))
				str_dt_rgproval_mes = zero(month(str_dt_rgproval))
				str_dt_rgproval_ano = year(str_dt_rgproval)
			else
				str_dt_rgproval_dia = ""
				str_dt_rgproval_mes = ""
				str_dt_rgproval_ano = ""
			end if
		str_endereco_alt = rs("nm_endereco_alt")
		
		str_sexo = rs("cd_sexo")		
		
		End If
		
Elseif strcod_ini <> "" Then
	xsql ="SELECT * FROM TBL_funcionario_inicial where cd_codigo="&strcod_ini&""
	Set rs = dbconn.execute(xsql)
		if not rs.EOF then
			strnome = rs("nm_nome")
			str_sexo = rs("cd_sexo")
			strnm_endereco = rs("nm_endereco")
			strnr_numero = rs("nr_numero")
			strbairro = rs("nm_bairro")
			strnm_complemento = rs("nm_complemento")
			strnm_cep = rs("nm_cep")
			strnm_cidade = rs("nm_cidade")
			strnm_estado = rs("nm_estado")
			strdt_nascimento = rs("dt_nascimento")
				strdia_nascimento = zero(day(strdt_nascimento))
				strmes_nascimento = zero(month(strdt_nascimento))
				strano_nascimento = year(strdt_nascimento)
			str_nacionalidade = rs("nm_nacionalidade")
			str_naturalidade = rs("nm_naturalidade")
			str_estado_civil = rs("cd_estado_civil")
			str_conjuge = rs("nm_conjuge")
		str_conjuge_nasc = rs("dt_conjuge_nasc")
			if str_conjuge_nasc <> "" Then str_conjuge_nasc = zero(day(str_conjuge_nasc))&"/"&zero(month(str_conjuge_nasc))&"/"&year(str_conjuge_nasc)
		strmae = rs("nm_mae")
		strpai = rs("nm_pai")
		str_pis = rs("cd_pis")
		str_ctps = rs("cd_ctps")
		str_ctps_serie = rs("cd_ctps_serie")
		strnm_rg = rs("nm_rg")
		strdt_rg_emissao = rs("dt_rg_emissao")
			if strdt_rg_emissao <> "" Then strdt_rg_emissao = zero(day(strdt_rg_emissao))&"/"&zero(month(strdt_rg_emissao))&"/"&year(strdt_rg_emissao)
		strnm_rg_expedidor = rs("nm_rg_expedidor")
		strnm_cpf = rs("nm_cpf")
		strnm_tit_eleitor = rs("nm_tit_eleitor")
		strnr_tit_eleitor_zona = rs("nm_tit_eleitor_zona")
		strnr_tit_eleitor_seccao = rs("nm_tit_eleitor_seccao")
		strnm_reservista = rs("nm_reservista")
		
		str_conspro = rs("cd_conspro")
		str_rgpro_cargo = rs("cd_rgpro_cargo")
		str_numreg = rs("cd_numreg")
		str_rgpro_concessao = rs("rgpro_concessao")
		str_rgpro_status = rs("cd_rgpro_status")
		str_rgpro = rs("obs_rgpro")
		str_dt_rgproinscr = rs("dt_rgpro_inscr")
			if str_dt_rgproinscr <> "" Then str_dt_rgproinscr = zero(day(str_dt_rgproinscr))&"/"&zero(month(str_dt_rgproinscr))&"/"&year(str_dt_rgproinscr)
		str_dt_rgpropag = rs("dt_rgpropag")
			if str_dt_rgpropag <> "" Then str_dt_rgpropag = zero(day(str_dt_rgpropag))&"/"&zero(month(str_dt_rgpropag))&"/"&year(str_dt_rgpropag)
		str_dt_rgproval = rs("dt_rgproval")	
			if str_dt_rgproval <> "" Then str_dt_rgproval = zero(day(str_dt_rgproval))&"/"&zero(month(str_dt_rgproval))&"/"&year(str_dt_rgproval)
		
		strnr_hepatite_b = rs("nr_hepatite_b")
		strnr_dupla_adulto = rs("nr_dupla_adulto")
		strdt_dupla_adulto_validade = rs("dt_dupla_adulto_validade")
			if strdt_dupla_adulto_validade <> "" Then strdt_dupla_adulto_validade = zero(day(strdt_dupla_adulto_validade))&"/"&zero(month(strdt_dupla_adulto_validade))&"/"&year(strdt_dupla_adulto_validade)
		strnr_scr = rs("nr_scr")
		
		strnr_caract_altura = rs("nr_caract_altura")
		strnr_caract_peso = rs("nr_caract_peso")
		strnr_caract_manequim = rs("nr_caract_manequim")
		strnr_caract_calcado = rs("nr_caract_calcado")
		
		strnm_parente = rs("nm_parente")
		strcd_parente_parentesco = rs("cd_parente_parentesco")
		
		strcd_residencia_tipo = rs("cd_residencia_tipo")
		strcd_residencia_financ = rs("cd_residencia_financ")
		strcd_residencia_reside = rs("cd_residencia_reside")
		strcd_veiculo_tipo = rs("cd_veiculo_tipo")
		strcd_veiculo_financ = rs("cd_veiculo_financ")
		
		'**** Contatos **** 
		nm_contato1 = rs("nm_contato1")
			if nm_contato1 <> "" Then
			pos_sep1 = instr(1,nm_contato1,"!",1)
			pos_sep2 = instr(1,nm_contato1,"#",1)
			pos_sep3 = instr(1,nm_contato1,"$",1)
			pos_sep4 = instr(1,nm_contato1,"*",1)
				tipo_contato1 = mid(nm_contato1,1,pos_sep1)
				ddd_contato1 = mid(nm_contato1,pos_sep1,pos_sep2-pos_sep1)
				tel_contato1 = mid(nm_contato1,pos_sep2,pos_sep3-pos_sep2)
				obs_contato1 = mid(nm_contato1,pos_sep3,pos_sep4-pos_sep3)
					tipo_contato1 = replace(tipo_contato1,"!","")
					ddd_contato1 = replace(ddd_contato1,"!","")
					tel_contato1 = replace(tel_contato1,"#","")
					obs_contato1 = replace(obs_contato1,"$","")
						contatos_total_inicial1 = tipo_contato1&"#"&nome_contato1&"{"&relac_contato1&"}"&ddd_contato1&"["&tel_contato1&"]"&obs_contato1&"_"
			end if
		
		nm_contato2 = rs("nm_contato2")
			if nm_contato2 <> "" Then
			pos_sep1 = instr(1,nm_contato2,"!",1)
			pos_sep2 = instr(1,nm_contato2,"#",1)
			pos_sep3 = instr(1,nm_contato2,"$",1)
			pos_sep4 = instr(1,nm_contato2,"*",1)
				tipo_contato2 = mid(nm_contato2,1,pos_sep1)
				ddd_contato2 = mid(nm_contato2,pos_sep1,pos_sep2-pos_sep1)
				tel_contato2 = mid(nm_contato2,pos_sep2,pos_sep3-pos_sep2)
				obs_contato2 = mid(nm_contato2,pos_sep3,pos_sep4-pos_sep3)
					tipo_contato2 = replace(tipo_contato1,"!","")
					ddd_contato2 = replace(ddd_contato2,"!","")
					tel_contato2 = replace(tel_contato2,"#","")
					obs_contato2 = replace(obs_contato2,"$","")
						contatos_total_inicial2 = tipo_contato2&"#"&nome_contato2&"{"&relac_contato2&"}"&ddd_contato2&"["&tel_contato2&"]"&obs_contato2&"_"
			end if
		
		nm_contato3 = rs("nm_contato3")
			if nm_contato3 <> "" Then
			pos_sep1 = instr(1,nm_contato3,"!",1)
			pos_sep2 = instr(1,nm_contato3,"#",1)
			pos_sep3 = instr(1,nm_contato3,"$",1)
			pos_sep4 = instr(1,nm_contato3,"*",1)
				tipo_contato3 = mid(nm_contato3,1,pos_sep1)
				ddd_contato3 = mid(nm_contato3,pos_sep1,pos_sep2-pos_sep1)
				tel_contato3 = mid(nm_contato3,pos_sep2,pos_sep3-pos_sep2)
					tipo_contato3 = replace(tipo_contato3,"!","")
					ddd_contato3 = replace(ddd_contato3,"!","")
					tel_contato3 = replace(tel_contato3,"#","")
					obs_contato3 = replace(obs_contato3,"$","")
						contatos_total_inicial3 = tipo_contato3&"#"&nome_contato3&"{"&relac_contato3&"}"&ddd_contato3x&"["&ddd_contato3&"]"&tel_contato3&"_"
			end if
		
		nm_contato4 = rs("nm_contato4")
			if nm_contato4 <> "" Then
			pos_sep1 = instr(1,nm_contato4,"!",1)
			pos_sep2 = instr(1,nm_contato4,"#",1)
			pos_sep3 = instr(1,nm_contato4,"$",1)
			pos_sep4 = instr(1,nm_contato4,"*",1)
			pos_sep5 = instr(1,nm_contato4,"%",1)
			pos_sep6 = instr(1,nm_contato4,"_",1)
				tipo_contato4 = mid(nm_contato4,1,pos_sep1)
				ddd_contato4 = mid(nm_contato4,pos_sep1,pos_sep2-pos_sep1)
				tel_contato4 = mid(nm_contato4,pos_sep2,pos_sep3-pos_sep2)
				obs_contato4 = mid(nm_contato4,pos_sep3,pos_sep4-pos_sep3)
				nome_contato4 = mid(nm_contato4,pos_sep4,pos_sep5-pos_sep4)
				relac_contato4 = mid(nm_contato4,pos_sep5,pos_sep6-pos_sep5)
					tipo_contato4 = replace(tipo_contato4,"!","")
					ddd_contato4 = replace(ddd_contato4,"!","")
					tel_contato4 = replace(tel_contato4,"#","")
					obs_contato4 = replace(obs_contato4,"$","")
					nome_contato4 = replace(nome_contato4,"*","")
					relac_contato4 = replace(relac_contato4,"%","")
						contatos_total_inicial4 = tipo_contato4&"#"&nome_contato4&"{"&relac_contato4&"}"&ddd_contato4&"["&tel_contato4&"]"&obs_contato4&"_"
			end if
			
			contatos_total_inicial = contatos_total_inicial1&"$"&contatos_total_inicial2&"$"&contatos_total_inicial3&"$"&contatos_total_inicial4
		
		'**** Dependentes ****
		'Filhote#1{12/01/2011}teste%20filhote[1]
		
		nm_dependente1 = rs("nm_dependente1")
			if nm_dependente1 <> "" Then
			pos_sep1 = instr(1,nm_dependente1,"!",1)
			pos_sep2 = instr(1,nm_dependente1,"#",1)
			pos_sep3 = instr(1,nm_dependente1,"$",1)
			pos_sep4 = instr(1,nm_dependente1,"*",1)
				nome_dependente1 = mid(nm_dependente1,1,pos_sep1)
				sexo_dependente1 = mid(nm_dependente1,pos_sep1,pos_sep2-pos_sep1)
				nascimento_dependente1 = mid(nm_dependente1,pos_sep2,pos_sep3-pos_sep2)
				obs_dependente1 = mid(nm_dependente1,pos_sep3,pos_sep4-pos_sep3)
					nome_dependente1 = replace(nome_dependente1,"!","")
					sexo_dependente1 = replace(sexo_dependente1,"!","")
					nascimento_dependente1 = replace(nascimento_dependente1,"#","")
					obs_dependente1 = replace(obs_dependente1,"$","")
					parentesco_dependente1 = "1"
							dependente_total_inicial1 = nome_dependente1&"#"&parentesco_dependente1&"{"&nascimento_dependente1&"}"&obs_dependente1&"["&sexo_dependente1&"]_"
			end if
			
		nm_dependente2 = rs("nm_dependente2")
			if nm_dependente2 <> "" Then
			pos_sep1 = instr(1,nm_dependente2,"!",1)
			pos_sep2 = instr(1,nm_dependente2,"#",1)
			pos_sep3 = instr(1,nm_dependente2,"$",1)
			pos_sep4 = instr(1,nm_dependente2,"*",1)
				nome_dependente2 = mid(nm_dependente2,1,pos_sep1)
				sexo_dependente2 = mid(nm_dependente2,pos_sep1,pos_sep2-pos_sep1)
				nascimento_dependente2 = mid(nm_dependente2,pos_sep2,pos_sep3-pos_sep2)
				obs_dependente2 = mid(nm_dependente2,pos_sep3,pos_sep4-pos_sep3)
					nome_dependente2 = replace(nome_dependente2,"!","")
					sexo_dependente2 = replace(sexo_dependente2,"!","")
					nascimento_dependente2 = replace(nascimento_dependente2,"#","")
					obs_dependente2 = replace(obs_dependente2,"$","")
					parentesco_dependente2 = "1"
							dependente_total_inicial2 = nome_dependente2&"#"&parentesco_dependente2&"{"&nascimento_dependente2&"}"&obs_dependente2&"["&sexo_dependente2&"]_"
			end if
			
		nm_dependente3 = rs("nm_dependente3")
			if nm_dependente3 <> "" Then
			pos_sep1 = instr(1,nm_dependente3,"!",1)
			pos_sep2 = instr(1,nm_dependente3,"#",1)
			pos_sep3 = instr(1,nm_dependente3,"$",1)
			pos_sep4 = instr(1,nm_dependente3,"*",1)
				nome_dependente3 = mid(nm_dependente3,1,pos_sep1)
				sexo_dependente3 = mid(nm_dependente3,pos_sep1,pos_sep2-pos_sep1)
				nascimento_dependente3 = mid(nm_dependente3,pos_sep2,pos_sep3-pos_sep2)
				obs_dependente3 = mid(nm_dependente3,pos_sep3,pos_sep4-pos_sep3)
					nome_dependente3 = replace(nome_dependente3,"!","")
					sexo_dependente3 = replace(sexo_dependente3,"!","")
					nascimento_dependente3 = replace(nascimento_dependente3,"#","")
					obs_dependente3 = replace(obs_dependente3,"$","")
					parentesco_dependente3 = "1"
							dependente_total_inicial3 = nome_dependente3&"#"&parentesco_dependente3&"{"&nascimento_dependente3&"}"&obs_dependente3&"["&sexo_dependente3&"]_"
			end if
			
			dependentes_total_inicial = dependente_total_inicial1&"$"&dependente_total_inicial2&"$"&dependente_total_inicial3&"$"&dependente_total_inicial4
		
		
		end if

End if
%>


<!--SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT-->
<script language="JavaScript">
function validar_cad(shipinfo){
	if (shipinfo.nm_nome.value.length==""){window.alert ("O nome deve ser preenchido.");shipinfo.nm_nome.focus();return (false);}
	if (shipinfo.cd_sexo.value.length==""){window.alert ("O sexo deve ser preenchido.");shipinfo.cd_sexo.focus();return (false);}
	if (shipinfo.dia_nascimento.value.length==""){window.alert ("A data de nascimento deve ser preenchida.");shipinfo.dia_nascimento.focus();return (false);}
	if (shipinfo.mes_nascimento.value.length==""){window.alert ("A data de nascimento deve ser preenchida.");shipinfo.mes_nascimento.focus();return (false);}
	if (shipinfo.ano_nascimento.value.length==""){window.alert ("A data de nascimento deve ser preenchida.");shipinfo.ano_nascimento.focus();return (false);}
	if (shipinfo.cd_regime_trabalho.value.length==""){window.alert ("O item Contrato deve ser preenchido.");shipinfo.cd_regime_trabalho.focus();return (false);}
	if (shipinfo.cd_dia.value.length==""){window.alert ("O dia de admissão deve ser preenchido.");shipinfo.cd_dia.focus();return (false);}
	if (shipinfo.cd_mes.value.length==""){window.alert ("O mês de admissão deve ser preenchido.");shipinfo.cd_mes.focus();return (false);}
	if (shipinfo.cd_ano.value.length==""){window.alert ("O ano de admissão deve ser preenchido.");shipinfo.cd_ano.focus();return (false);}
	if (shipinfo.cd_cargo.value.length==""){window.alert ("O cargo deve ser preenchido.");shipinfo.cd_cargo.focus();return (false);}
	if (shipinfo.cd_status.value.length==""){window.alert ("Informe a situação do funcionário.");shipinfo.cd_status.focus();return (false);}
	if (shipinfo.cd_unidade.value.length==""){window.alert ("Informe a unidade do funcionário.");shipinfo.cd_unidade.focus();return (false);}
	if (shipinfo.hr_entrada.value.length==""){window.alert ("Informe o horário de entrada do funcionário.");shipinfo.hr_entrada.focus();return (false);}
	if (shipinfo.nm_intervalo.value.length==""){window.alert ("Informe o tipo de intervalo do funcionário.");shipinfo.nm_intervalo.focus();return (false);}
	if (shipinfo.hr_saida.value.length==""){window.alert ("Informe o horário de saída do funcionário.");shipinfo.hr_saida.focus();return (false);}	
		
	return (true);}
	
</script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../../js/ajax2.js"></SCRIPT>
<!--include file="../../js/ajax2.asp"-->

<script language="Javascript">
<!-- 
//-------------------------------------------------------------------
function rep_space(string) {
	while (string.indexOf('a') != -1) {
 		string = string.replace('a','%20');
	}
	return string;
}
//--------------------------------------------------------------------

	
function adiciona_dependente(a,b,c,d,e,f,g,h,dependentes_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	f=f;
	g=g;
	h=h;
	dependentes_total=dependentes_total;
	
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}
			while (b.indexOf(' ') != -1) {
	 		b = b.replace(' ','%20');}
			while (c.indexOf(' ') != -1) {
	 		c = c.replace(' ','%20');}
			while (d.indexOf(' ') != -1) {
	 		d = d.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		e = e.replace(' ','%20');}
			while (f.indexOf(' ') != -1) {
	 		f = f.replace(' ','%20');}
			while (f.indexOf(' ') != -1) {
	 		g = g.replace(' ','%20');}
			while (f.indexOf(' ') != -1) {
	 		h = h.replace(' ','%20');}
			
			//contatos_total = contatos_total.replace(' ','%20');}
			
			while (dependentes_total.indexOf('$$') != -1) {
	 		dependentes_total = dependentes_total.replace('$$','$');}
						
		if (dependentes_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	dependentes_total = dependentes_total + separador
		if (a != ""){ //adiciona	
				dependentes_total = dependentes_total + a + '!' + c + '{' + d+'/'+e+'/'+f + '}' + g + '[' + h + ']';
				document.form.dependentes_result.value=a;
			}			
						
		if (b != ""){ //subtrai			
				dependentes_total = dependentes_total.replace(b,'');
				document.form.dependentes_total.value=dependentes_total;
			}

	document.form.dependentes_total.value=(dependentes_total.replace("$$", "$"));	
	document.form.dependentes_total.value=(dependentes_total); //c	
	document.form.nm_dependente_nome.value="";
	document.form.cd_dependente_parentesco.value="";
	document.form.dt_dependente_nascimento_dia.value="";
	document.form.dt_dependente_nascimento_mes.value="";
	document.form.dt_dependente_nascimento_ano.value="";
	document.form.cd_dependente_sexo.value="";
	document.form.cd_dependente_obs.value="";
	document.form.nm_dependente_nome.focus();
	
	//chama Ajax	
		loadXMLDoc("dependentes/ajax/ajax_dependentes_lista.asp?dependentes_total="+dependentes_total,function()
			{
		  		if (xmlhttp.readyState==4)//&& xmlhttp.status==200)
		    		{
		    			document.getElementById("divDependente_lista").innerHTML = xmlhttp.responseText;
		    		}
		  		}
			);
				
}

function adiciona_contato(a,b,c,d,e,f,g,contatos_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	f=f;
	g=g;
	contatos_total=contatos_total;
	//nextfield ="cd_material_1";
		
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}
			while (b.indexOf(' ') != -1) {
	 		b = b.replace(' ','%20');}
			while (c.indexOf(' ') != -1) {
	 		c = c.replace(' ','%20');}
			while (d.indexOf(' ') != -1) {
	 		d = d.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		e = e.replace(' ','%20');}
			while (f.indexOf(' ') != -1) {
	 		f = f.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		g = g.replace(' ','%20');}			
			
			while (contatos_total.indexOf('$$') != -1) {
	 		contatos_total = contatos_total.replace('$$','$');}
						
		if (contatos_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	contatos_total = contatos_total + separador
		if (a != ""){
			contatos_total = contatos_total + a + '!' + c + '{' + d + '}' + e + '[' + f + ']' + g + '_';
			document.form.contatos_result.value=a;} //adiciona				
						
		if (b != ""){
			contatos_total = contatos_total.replace(b,'');
			document.form.contatos_total.value=contatos_total;} //subtrai			

	document.form.contatos_total.value=(contatos_total.replace("$$", "$"));	
	document.form.contatos_total.value=(contatos_total); //c	
	document.form.nm_contato_nome.value="";
	document.form.nm_contato_cargo.value="";
	document.form.cd_contato_categoria.value="";
	document.form.cd_contato_ddd.value="";
	document.form.nm_contato_numero.value="";
	document.form.nm_contato_obs.value="";
	document.form.cd_contato_categoria.focus();
	
	//function myFunction2()
		//{
			loadXMLDoc("../../contatos/ajax/ajax_contatos_lista.asp?contatos_total="+contatos_total,function()
		  		{
		  			if (xmlhttp.readyState==4 && xmlhttp.status==200)
		    			{
		    				document.getElementById("divContatos_lista").innerHTML = xmlhttp.responseText;
		    			}
		  			}
				);
		//}
		
	
}


function adiciona_escolaridade(a,b,c,d,e,f,g,h,i,escolaridade_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	f=f;
	g=g;
	h=h;
	i=i;
	escolaridade_total=escolaridade_total;
			
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}
			while (b.indexOf(' ') != -1) {
	 		b = b.replace(' ','%20');}
			while (c.indexOf(' ') != -1) {
	 		c = c.replace(' ','%20');}
			while (d.indexOf(' ') != -1) {
	 		d = d.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		e = e.replace(' ','%20');}
			while (f.indexOf(' ') != -1) {
	 		f = f.replace(' ','%20');}
			while (g.indexOf(' ') != -1) {
	 		g = g.replace(' ','%20');}
			while (g.indexOf(' ') != -1) {
	 		h = h.replace(' ','%20');}
			while (g.indexOf(' ') != -1) {
	 		i = i.replace(' ','%20');}			
			
			while (escolaridade_total.indexOf('$$') != -1) {
	 		escolaridade_total = escolaridade_total.replace('$$','$');}
						
		if (escolaridade_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	escolaridade_total = escolaridade_total + separador
		if (a != ""){
			escolaridade_total = escolaridade_total + a + '!' + c + '{' + d + '}' + e + '[' + f+'/'+g+'/'+h + ']' + i + '_';
			document.form.escolaridade_result.value=a;} //adiciona				
						
		if (b != ""){
			escolaridade_total = escolaridade_total.replace(b,'');
			document.form.escolaridade_total.value=escolaridade_total;} //subtrai			

	document.form.escolaridade_total.value=(escolaridade_total.replace("$$", "$"));	
	document.form.escolaridade_total.value=(escolaridade_total); //c	
	document.form.cd_ensino_grau.value="";
	document.form.nm_instituicao.value="";
	document.form.nm_curso.value="";
	document.form.cd_ensino_andamento.value="";
	document.form.dt_termino_dia.value="";
	document.form.dt_termino_mes.value="";
	document.form.dt_termino_ano.value="";
	document.form.nm_obs.value="";
	document.form.cd_ensino_grau.focus();
	//chama Ajax	
	loadXMLDoc("escolaridade/ajax/ajax_escolaridade_lista.asp?escolaridade_total="+escolaridade_total,function()
			{
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200)
		    		{
		    			document.getElementById("divEscolaridade_lista").innerHTML = xmlhttp.responseText;
		    		}
		  		}
			);		
}


function adiciona_emprego(a,b,c,d,e,f,g,h,i,j,emprego_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	f=f;
	g=g;
	h=h;
	i=i;
	j=j;
	emprego_total=emprego_total;
	
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}
			while (b.indexOf(' ') != -1) {
	 		b = b.replace(' ','%20');}
			while (c.indexOf(' ') != -1) {
	 		c = c.replace(' ','%20');}
			while (d.indexOf(' ') != -1) {
	 		d = d.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		e = e.replace(' ','%20');}
			while (f.indexOf(' ') != -1) {
	 		f = f.replace(' ','%20');}
			while (g.indexOf(' ') != -1) {
	 		g = g.replace(' ','%20');}
			while (h.indexOf(' ') != -1) {
	 		h = h.replace(' ','%20');}
			while (i.indexOf(' ') != -1) {
	 		i = i.replace(' ','%20');}
			while (j.indexOf(' ') != -1) {
	 		j = j.replace(' ','%20');}
			
			while (emprego_total.indexOf('$$') != -1) {
	 		emprego_total = emprego_total.replace('$$','$');}
						
		if (emprego_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	emprego_total = emprego_total + separador
		if (a != ""){
			emprego_total = emprego_total + a + '!' + c + '{' + d+'/'+e+'/'+f + '}' + g+'/'+h+'/'+i + '[' + j + ']';
			document.form.emprego_result.value=a;} //adiciona				
						
		if (b != ""){
			emprego_total = emprego_total.replace(b,'');
			document.form.emprego_total.value=emprego_total;} //subtrai			

	document.form.emprego_total.value=(emprego_total.replace("$$", "$"));	
	document.form.emprego_total.value=(emprego_total); //c
	
	document.form.nm_emprego_empresa.value="";
	document.form.nm_emprego_cargo.value="";
	document.form.dt_emprego_entrada_dia.value="";
	document.form.dt_emprego_entrada_mes.value="";
	document.form.dt_emprego_entrada_ano.value="";
	document.form.dt_emprego_saida_dia.value="";
	document.form.dt_emprego_saida_mes.value="";
	document.form.dt_emprego_saida_ano.value="";
	document.form.dt_emprego_obs.value="";
	
	document.form.nm_emprego_empresa.focus();
	//chama Ajax	
	
	
 		loadXMLDoc("emprego_anterior/ajax/ajax_emprego_lista.asp?emprego_total="+emprego_total,function()
			{
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200)
		    		{
		    			document.getElementById("divEmprego_lista").innerHTML = xmlhttp.responseText;
		    		}
		  		}
			);
}

function adiciona_valores(a,b,c,d,e,valores_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	valores_total=valores_total;
			
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}
			while (b.indexOf(' ') != -1) {
	 		b = b.replace(' ','%20');}
			while (c.indexOf(' ') != -1) {
	 		c = c.replace(' ','%20');}
			while (d.indexOf(' ') != -1) {
	 		d = d.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		e = e.replace(' ','%20');}
			
			while (valores_total.indexOf('$$') != -1) {
	 		valores_total = valores_total.replace('$$','$');}
						
		if (valores_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	valores_total = valores_total + separador
		if (a != ""){
			valores_total = valores_total + a + '#' + c + '#' + d + '#' + e + '#';
			document.form.valores_result.value=a;} //adiciona				
						
		if (b != ""){
			valores_total = valores_total.replace(b,'');
			document.form.valores_total.value=valores_total;} //subtrai			

	document.form.valores_total.value=(valores_total.replace("$$", "$"));	
	document.form.valores_total.value=(valores_total); //c
	document.form.cd_valor_tipo.value="";
	document.form.nr_valor.value="";
	document.form.dt_valor_atualizacao.value="";
	document.form.nm_valor_obs.value="";
	
	document.form.cd_valor_tipo.focus();
	//chama Ajax	
	{
		var oHTTPRequest_valores = createXMLHTTP(); 
		oHTTPRequest_valores.open("post", "valores/ajax/ajax_valores_lista.asp", true);
		oHTTPRequest_valores.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_valores.onreadystatechange=function(){
		if (oHTTPRequest_valores.readyState==4){
	    document.all.divValores_lista.innerHTML = oHTTPRequest_valores.responseText;}}
	    oHTTPRequest_valores.send("valores_total=" + form.valores_total.value);
	}
 
}


function mostra_urgencia(){
	var objeto1='hidden_mostra1';
	var objeto2='hidden_mostra2';	
		
	if (document.form.cd_contato_categoria.value==5){
			var el1=document.getElementById(objeto1);
				el1.style.display='inline';		
			var el2=document.getElementById(objeto2);
				el2.style.display='inline';
		}
	if (document.form.cd_contato_categoria.value!=5){
			var el1=document.getElementById(objeto1);
				el1.style.display='none';
			var el2=document.getElementById(objeto2);
				el2.style.display='none';
		}
		
	}
//---------------------------------------------------------------------->
</script>
<style type="text/css" media="screen">
<!--
.ok_print{ visibility:hidden; display: none;}
.no_print{ visibility:visible; display:block;}
#frame{border:0px inset;
	position:absolute;
	height:93%;
	width:79%;
	top:20px;
	left:210px;
	margin: 0px;
	padding: 6px;
	text-align:left;
	overflow:scroll;
	text-decoration:none;}
table{background-color: #ffffff; 
    border: 1px solid #cccccc;
	font-family: verdana;
	font-size: 9px;
	text-decoration:none;
	border-collapse: collapse;}

a:link {color:#000000;text-decoration:none;}
a:visited {color:#000000;text-decoration:none;}
a:hover {color:#FF0000;}
a:active {color:#000000;}

.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.cabecalho {color: #000000;font-family: verdana;font-size:14px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; } 

-->
</style>

<%cor_campo = "#ebebeb"%>
					
					<%If strmsg <> "" then%>
						<table  border="0" cellspacing="0" cellpadding="0" class="txt_cinza">
							<tr class="no_print">
							 <td>
								<b><%=strmsg%></b>
							 </td>
							</tr>
							<tr class="no_print">
							 <td align=center><a href="coop_cooperados_lista.asp"><img src="../../imagens/bt_lista.gif" alt="" width="119" height="22" border="1"></a></td>
							</tr>
						</table>
					<%else%>
					<table border="1" cellspacing="0" cellpadding="1" bordercolorlight="#c0c0c0"  align="center" class="no_print" style="collapse-border:collpase;">
						<form name="form" id="form" action="../acoes/funcionarios_acao.asp"  method="POST" onsubmit="return validar_cad(document.form);">
						<input type="hidden" name="cd_sessao" value="<%=session_cd_usuario%>">
						<input type="hidden" name="jan" value="1">
						<%If strcod = "" then%>
						<input type="hidden" name="acao" value="insert">
						<%ElseIf strcod <> ""  then %>
						<input type="hidden" name="acao" value="altera">
						<input type="hidden" name="cod" value="<%=strcod%>">
						<input type="hidden" name="cd_funcionario" value="<%=str_funcionario%>">
						<input name="foto_h" type="hidden" value="<%=strfoto%>">
						<%End if%>
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
							<tr>
								<td colspan="4" align="center" class="cabecalho" style="background-color:black; color:white;"><b>FICHA CADASTRAL DO COLABORADOR <%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/px.gif" width="712" height="1">&nbsp;</td></tr>
							<tr>
								<td rowspan="6" align="center" valign="<%if strfoto = "Vdlap.gif" then%>middle<%else%>top<%end if%>" style="color:#7d7d7d;">									
										<img src="../../foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"><br>
									<%if strcod <> "" then%>
										<a href="#" onClick="window.open('../../empresa/janelas/funcionarios_foto.asp?cd_codigo=<%=strcod%>','upload','width=500,height=550')" alt="Inserir/mudar foto">(Alterar foto)</a>
									<%end if%>
								</td>
								<%'if 
								'end if%>
								<td bgcolor="<%=cor_campo%>"><b>Matrícula</b></td>
								<td bgcolor="<%=cor_campo%>"><input type="text" name="cd_matricula" value="<%=strmatricula%>" size="10" maxlength="4" class="inputs"></td>
								<td bgcolor="<%=cor_campo%>">&nbsp;</td>								
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="3"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="2" style="color:red;"><b>Nome</b></td>
								<td style="color:red;"><b>Sexo</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="2"><input type="text" name="nm_nome" value="<%=strnome %>" size="50" maxlength="70" class="inputs"></td>
								<td>
									<select name="cd_sexo"  class="inputs">
											<option value=""></option>
											<option value="1" <%if str_sexo = 1 then response.write("SELECTED")%>>Masculino</option>
											<option value="2" <%if str_sexo = 2 then response.write("SELECTED")%>>Feminino</option>
									</select>								
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="2"><b>Endereço / Nº</b></td>
								<td><b>Complemento</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="2"><input type="text" name="nm_endereco" value="<%=strnm_endereco%>" size="40" maxlength="100" class="inputs">&nbsp;
								<input type="text" name="nr_numero" value="<%=strnr_numero%>" size="4" maxlength="8" class="inputs"></td>
								<td><input type="text" name="nm_complemento" value="<%=strnm_complemento%>" size="25" maxlength="30" class="inputs"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Bairro</b></td>
								<td><b>CEP</b></td>
								<td><b>Cidade</b></td>
								<td><b>UF</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><input type="text" name="nm_bairro" value="<%=strbairro%>" size="20" maxlength="100" class="inputs"></td>
								<td><input type="text" name="nm_cep" value="<%=strnm_cep%>" size="10" maxlength="9" class="inputs"></td>
								<td><input type="text" name="nm_cidade" value="<%=strnm_cidade%>" size="25" maxlength="30" class="inputs"></td>
								<td><input type="text" name="nm_estado" value="<%=strnm_estado%>" size="1" maxlength="2" class="inputs"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td style="color:red;"><b>Data de nascimento</b></td>
								<td><b>Nacionalidade</b></td>
								<td><b>Naturalidade (Cidade)</b></td>
								<td><b>Estado Civil</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><input type="text" name="dia_nascimento" value="<%=strdia_nascimento%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
									<input type="text" name="mes_nascimento" value="<%=strmes_nascimento%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
									<input type="text" name="ano_nascimento" value="<%=strano_nascimento%>" size="4" maxlength="4" class="inputs">
								<td><input type="text" name="nm_nacionalidade" value="<%=str_nacionalidade%>" size="12" maxlength="25" class="inputs"></td>
								<td><input type="text" name="nm_naturalidade" value="<%=str_naturalidade%>" size="20" maxlength="30" class="inputs"></td>
								<td>
									<select name="cd_estado_civil" class="inputs">
									<option value=""></option>
									<%strsql = "Select * From TBL_estado_civil order by cd_codigo"
									Set rs_est_civil = dbconn.execute(strsql)%>
									<%Do While Not rs_est_civil.eof
											cd_estado_civil = rs_est_civil("cd_codigo")
											nm_estado_civil = rs_est_civil("nm_estado_civil")%>	
											<option value="<%=cd_estado_civil%>" <%if int(str_estado_civil) = cd_estado_civil then%>Selected<%end if%>><%=nm_estado_civil%></option>
											<%rs_est_civil.movenext
										loop
											rs_est_civil.close
											Set rs_est_civil = nothing%>											
									</select>
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Nome do Cônjuge</b></td>
								<td><b>Nasc. cônjuge</b></td>
								<td><b>Opta plan. saúde p/ conj.?</b></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">								
								
								<td><input type="text" name="nm_conjuge" value="<%=str_conjuge%>" size="25" maxlength="100" class="inputs"></td>
								<td><input type="text" name="dt_conjuge_nasc_dia" value="<%=str_conjuge_nasc_dia%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_conjuge_nasc_mes" value="<%=str_conjuge_nasc_mes%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_conjuge_nasc_ano" value="<%=str_conjuge_nasc_ano%>" size="4" maxlength="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
								<td><input type="checkbox" name="cd_plano_saude_conj" value="1" <%if str_cd_plano_saude_conj = 1 then%>checked<%end if%>  class="inputs">Sim</td>
								<td>&nbsp;</td>
							</tr>
							<tr><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							
							<!--tr bgcolor="<%=cor_campo%>">
								<td><b>Endereço alternativo</b></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4"><input type="text" name="nm_endereco_alt" value="<%=str_endereco_alt%>" size="100 maxlength="150" class="inputs"></td>
							</tr>
							<tr><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
							
<%'cor_campo = "#ffffff"%>	
							<tr bgcolor="#a1a1a1"><td colspan="4" style="color:white;"><b>FILIAÇÃO</b></td></tr>
							<tr bgcolor="<%=cor_campo%>">
							<td colspan="2"><b>Mãe</b><img src="../../imagens/px.gif" width="40" height="1"><input type="text" name="nm_mae" value="<%=strmae%>" size="40" maxlength="100" class="inputs"></td>
							<td colspan="2"><b>Pai</b><img src="../../imagens/px.gif" width="40" height="1"><input type="text" name="nm_pai" value="<%=strpai%>" size="40" maxlength="100" class="inputs"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
								<!-- ****** CONTATOS!!!! ****** -->
									<%'id_origem = str_funcionario
									'cd_origem = 4
									'pag_retorno = ".."&mem_posicao
									'variaveis = "../../empresa.asp?tipo=cadastro$cod="&str_funcionario%>
									<!--include file="../../contatos/default.asp"-->
								<!-- ****** CONTATOS!!!! ****** -->							
								</td>
							</tr>
							<tr><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<%cor_campo = "#ffffff"%>	
							<tr bgcolor="#a1a1a1"><td colspan="4" style="color:white;"><b>CONTATOS</b></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<table border="0" style="border-collapse:collapse; border:solid white 0px; ">
										<tr bgcolor="<%=cor_campo%>">
											<!--td><b>Nome</b></td-->
											<td align="left"><b>Tipo</b><br><img src="../../imagens/px.gif" width="100" height="1"></td>
											<td align="left" id="hidden_mostra1" style="display:none;"><b>Nome<img src="../../imagens/px.gif" width="50" height="1"> Grau relac.</b><br><img src="../../imagens/px.gif" width="180" height="1"></td>
											<td align="left"><b>Tel./Cel/E-mail/Nextel</b><br><img src="../../imagens/px.gif" width="190" height="1"></td>
											<td align="left"><b>Observações</b><br><img src="../../imagens/px.gif" width="200" height="1"></td>
											<td>&nbsp;</td>
										</tr>
										<input type="hidden" name="cd_contato_origem" value="4">
										<input type="hidden" name="id_contato_origem" value="<%=str_funcionario%>">
										<tr bgcolor="<%=cor_campo%>">
											<td>								
											<select name="cd_contato_categoria" class="inputs" onchange="mostra_urgencia();">
												<option value=""></option>
												<%strsql = "Select * From TBL_contato_cat "'Where cd_codigo=1 or cd_codigo = 2"
												Set rs_contato = dbconn.execute(strsql)%>
												<%Do While Not rs_contato.eof
														categoria_contato = rs_contato("categoria_contato")
														cd_contato_cat = rs_contato("cd_codigo")%>	
														<option value="<%=cd_contato_cat%>" <%=ckcontr%>><%=categoria_contato%></option>
														<%rs_contato.movenext
													loop
														rs_contato.close
														Set rs_contato = nothing%>											
											</select></td>
											<td id="hidden_mostra2" style="display:none;">
												<input type="text" name="nm_contato_nome" value="<%=str_contato_nome%>" size="10" maxlength="50" class="inputs">&nbsp;
												<input type="text" name="nm_contato_cargo" value="<%=str_contato_cargo%>" size="10" maxlength="50" class="inputs">&nbsp;</td>
											<td><input type="text" name="cd_contato_ddd" value="<%=str_contato_ddd%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">&nbsp;
												<input type="text" name="nm_contato_numero" value="<%=str_contato%>" size="20" maxlength="100" class="inputs"></td>
											<td><input type="text" name="nm_contato_obs" value="<%=str_contato_obs%>" size="25" maxlength="100" class="inputs"></td>
											<td><input type="button" name="inclui_contato" value="+" onFocus="adiciona_contato(document.form.cd_contato_categoria.value,'',document.form.nm_contato_nome.value,document.form.nm_contato_cargo.value,document.form.cd_contato_ddd.value,document.form.nm_contato_numero.value,document.form.nm_contato_obs.value,document.form.contatos_total.value)"></td>
										</tr>
									</table>
								</td>
							</tr>
							<!--tr><td colspan="4"--><input type="hidden" name="contatos_total" value="<%=contatos_total_inicial%>" size="100">	
							<input type="hidden" name="contatos_result" size="70">
							<!--/td></tr-->
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><span id='divContatos_lista'></span><%'=contatos_total_inicial%></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_contato Where cd_origem=4 and id_origem='"&str_funcionario&"' order by cd_categoria"
											Set rs_contato = dbconn.execute(strsql)
											
											'if not rs_contato.EOF then
											while not rs_contato.EOF
												cd_contato = rs_contato("cd_codigo")
												nm_contato_nome = rs_contato("nm_nome")
												cd_contato_categoria = rs_contato("cd_categoria")
												nm_cargo = rs_contato("nm_cargo")
												cd_contato_ddd = rs_contato("cd_ddd")
												nm_contato_numero = rs_contato("nm_numero")
												nm_contato_obs = rs_contato("nm_obs")	
													if int(cd_contato_categoria) = 5 then
														cabeca_contato = 0
													end if%>
										
												<%if cabeca_contato = 0 Then%>
												<tr bgcolor="#808080" style="color:white;">													
													<td>&nbsp;<br><img src="../../imagens/px.gif" width="80" height="1"></td>
													<td><b>Tel/Cel/Email</b><br><img src="../../imagens/px.gif" width="220" height="1"></td>
													<td><b>Observações</b><br><img src="../../imagens/px.gif" width="100" height="1"></td>
													<%if cd_contato_categoria = 5 then%>
													<td>Nomex<br><img src="../../imagens/px.gif" width="150" height="1"></td>
													<td>Grau relac.<br><img src="../../imagens/px.gif" width="100" height="1"></td>
													<%end if%>
												</tr>
												<%end if%>
												
												<tr id="contatos_relat">
													<%strsql = "Select * From TBL_contato_cat Where cd_codigo='"&cd_contato_categoria&"'"'Where cd_codigo=1 or cd_codigo = 2"
													Set rs_contato_cat = dbconn.execute(strsql)%>
													<%Do While Not rs_contato_cat.eof
														categoria_contato = rs_contato_cat("categoria_contato")
														cd_contato_cat = rs_contato_cat("cd_codigo")
													rs_contato_cat.movenext
													loop
													rs_contato_cat.close
													Set rs_contato_cat = nothing%>													
													<td <%if cd_contato_categoria = 5 then%> style="color:red; font-weight:bold;"<%end if%>>&nbsp;<%=categoria_contato%></td>
													<td>&nbsp;<%if cd_contato_ddd <> "" Then response.write(cd_contato_ddd &" - ")%> <%=nm_contato_numero%></td>
													<td>&nbsp;<%=nm_contato_obs%></td>
													<%if cd_contato_categoria = 5 then%>
													<td>&nbsp;<%=nm_contato_nome%></td>
													<td>&nbsp;<%=nm_cargo%></td>
													<%end if%>
													<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsContatoDelete('<%=str_funcionario%>','<%=cd_contato%>')"></td>
												</tr>
										<%cabeca_contato = cabeca_contato + 1
										
										
										rs_contato.movenext
										wend%>									
									</table>
								</td>
							</tr>							
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<%cor_campo = "#ebebeb"%>
							<tr bgcolor="#a1a1a1">
								<td colspan="4" style="color:white;"><b>DOCUMENTAÇÃO</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">								
								<td><b>PIS</b></td>
								<td><b>Carteira Profissional</b></td>
								<td><b>Série</b></td>								
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">								
								<td><input type="text" name="cd_pis" value="<%=str_pis%>" size="25" maxlength="25" class="inputs"></td>
								<td><input type="text" name="cd_ctps" value="<%=str_ctps%>" size="23" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_ctps_serie" size="5" class="inputs" maxlength="5" value="<%=str_ctps_serie%>"></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>RG</b></td>
								<td><b>Data emissão</b></td>
								<td><b>Órgão exped.</b></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><input type="text" name="nm_rg" value="<%=strnm_rg%>" size="25" maxlength="15" class="inputs"></td>
								<td><input type="text" name="dt_rg_emissao_dia" value="<%=strdt_rg_emissao_dia%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_rg_emissao_mes" value="<%=strdt_rg_emissao_mes%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_rg_emissao_ano" value="<%=strdt_rg_emissao_ano%>" size="4" maxlength="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
								<td><input type="text" name="nm_rg_expedidor" value="<%=strnm_rg_expedidor%>" size="4" maxlength="50" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>CPF</b></td>
								<td><b>Tit. Eleitor</b></td>
								<td><b>Zona &nbsp; &nbsp; &nbsp; &nbsp; - &nbsp; Secção</b></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><input type="text" name="nm_cpf" value="<%=strnm_cpf%>" size="25" maxlength="14" class="inputs"></td>
								<td><input type="text" name="nm_tit_eleitor" value="<%=strnm_tit_eleitor%>" size="20" maxlength="14" class="inputs"></td>
								<td><input type="text" name="nr_tit_eleitor_zona" value="<%=strnr_tit_eleitor_zona%>" size="5" maxlength="14" class="inputs"> &nbsp;
									<input type="text" name="nr_tit_eleitor_seccao" value="<%=strnr_tit_eleitor_seccao%>" size="5" maxlength="14" class="inputs"></td>
								<td>&nbsp;</td>	
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Reservista</b></td>
								<td colspan="2"><b>Categoria</b></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><input type="text" name="nm_reservista" value="<%=strnm_reservista%>" size="25" maxlength="14" class="inputs"></td>
								<td colspan="2"><input type="text" name="nm_reservista_cat" value="<%=strnm_reservista_cat%>" size="50" maxlength="200" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<!--*****************-->
							<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><a name="coren"></a><b>Registro Profissional</b></td>
								<td><b>Número / Cargo</b></td>
								<td><b>Tipo de concessão / Situação</b></td>
								<td><b>Observações</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td>
									<select name="cd_conspro" class="inputs">
										<option value="1"></option>
										<%strsql = "Select * From TBL_cons_profissional "
					  					Set rs_conspro = dbconn.execute(strsql)
											While not rs_conspro.EOF
												cod_conspro = rs_conspro("cd_codigo")
												nm_conspro = rs_conspro("nm_conspro")%>																		
												<option value="<%=cod_conspro%>" <%if cod_conspro = int(str_conspro) then response.write("Selected")%>><%=nm_conspro%></option>
												<%rs_conspro.movenext
											wend%>
									</select></td>
								<td><input type="text" name="cd_numreg" value="<%=str_numreg%>" size="12" maxlength="25" class="inputs">
									<select name="cd_rgpro_cargo" class="inputs">
										<option value=""></option>
										<%strsql = "Select * From Tbl_rgpro_cargos "
					  					Set rs_rgpro = dbconn.execute(strsql)
											While not rs_rgpro.EOF
												cod_rgpro_cargo = rs_rgpro("cd_codigo")
												nm_rgpro_abrev = rs_rgpro("nm_rgpro_abrev")%>																		
												<option value="<%=cod_rgpro_cargo%>" <%if cod_rgpro_cargo = int(str_rgpro_cargo) then response.write("Selected")%>><%=nm_rgpro_abrev%></option>
												<%rs_rgpro.movenext
											wend%>
									</select></td>								
								<td valign="top">
									<select name="rgpro_concessao" class="inputs">
											<%if str_rgpro_concessao = "" then
												concessao_sel_0 =	"SELECTED"
											elseif str_rgpro_concessao = 1 then
												concessao_sel_1 =	"SELECTED"
											elseif str_rgpro_concessao = 2 then
												concessao_sel_2 =	"SELECTED"
											end if%>
											<option value="" <%=concessao_sel_0%>></option>
											<option value="1" <%=concessao_sel_1%>>Definitivo</option>
											<option value="2" <%=concessao_sel_2%>>Provisório</option>
									</select>&nbsp;									
									<select name="rgpro_status" class="inputs">
											<%if str_rgpro_status = "" then
												rgpro_status_sel_0 =	"SELECTED"
											elseif str_rgpro_status = 1 then
												rgpro_status_sel_1 =	"SELECTED"
											elseif str_rgpro_status = 2 then
												rgpro_status_sel_2 =	"SELECTED"
											end if%>
											<option value="" <%=rgpro_status_sel_0%>></option>
											<option value="1" <%=rgpro_status_sel_1%>>Ativo</option>
											<option value="2" <%=rgpro_status_sel_2%>>Inativo</option>
									</select></td>
								<td valign="top" rowspan="3"><textarea cols="25" rows="4" name="obs_rgpro" class="inputs"><%=str_rgpro%></textarea></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Data de inscrição</b></td>
								<td><b>Data de pagamento</b></td>
								<td><b>Validade</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><input type="text" name="dt_rgproinscr_dia" size="2" maxlength="2" class="inputs" value="<%=str_dt_rgproinscr_dia%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_rgproinscr_mes" size="2" maxlength="2" class="inputs" value="<%=str_dt_rgproinscr_mes%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_rgproinscr_ano" size="4" maxlength="4" class="inputs" value="<%=str_dt_rgproinscr_ano%>"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
								<td><input type="text" name="dt_rgpropag_dia" size="2" maxlength="2" class="inputs" value="<%=str_dt_rgpropag_dia%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_rgpropag_mes" size="2" maxlength="2" class="inputs" value="<%=str_dt_rgpropag_mes%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_rgpropag_ano" size="4" maxlength="4" class="inputs" value="<%=str_dt_rgpropag_ano%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
								<td><input type="text" name="dt_rgproval_dia" size="2" maxlength="2" class="inputs" value="<%=str_dt_rgproval_dia%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_rgproval_mes" size="2" maxlength="2" class="inputs" value="<%=str_dt_rgproval_mes%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_rgproval_ano" size="4" maxlength="4" class="inputs" value="<%=str_dt_rgproval_ano%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
								<!--td><input type="text" name="dia_valid" size="3" maxlength="2" class="inputs">/<input type="text" name="mes_valid" size="3" maxlength="2" class="inputs">/<input type="text" name="ano_valid" size="5" maxlength="4" class="inputs"></td-->
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ffffff"%>
							<tr bgcolor="#a1a1a1">
								<td colspan="4" style="color:white;"><b>CARTEIRA DE VACINAÇÃO</b><a name="vacinas"></a></td>
							</tr>

							<tr bgcolor="<%=cor_campo%>">
								<td><b>Hepatite B</b></td>
								<td colspan="3">
									<%if int(strnr_hepatite_b) = 3 then
										ckhepatite_b3 = "checked"
										ckhepatite_b2 = "checked"
										ckhepatite_b1 = "checked"
									elseif int(strnr_hepatite_b) = 2 then
										ckhepatite_b2 = "checked"
										ckhepatite_b1 = "checked"
									elseif int(strnr_hepatite_b) = 1 then
										ckhepatite_b1 = "checked"
									end if%>
									<input type="checkbox" name="nr_hepatite_b1" value="1" style="border:0;" class="inputs" <%=ckhepatite_b1%>>1ª dose &nbsp; 
									<input type="checkbox" name="nr_hepatite_b2" value="2" style="border:0;" class="inputs" <%=ckhepatite_b2%>>2ª dose &nbsp;
									<input type="checkbox" name="nr_hepatite_b3" value="3" style="border:0;" class="inputs" <%=ckhepatite_b3%>>3ª dose</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Dupla Adulto</b></td>
								<td colspan="3">
									<%if int(strnr_dupla_adulto) = 3 then
										ckdp_adulto_b3 = "checked"
										ckdp_adulto_b2 = "checked"
										ckdp_adulto_b1 = "checked"
									elseif int(strnr_dupla_adulto) = 2 then
										ckdp_adulto_b2 = "checked"
										ckdp_adulto_b1 = "checked"
									elseif int(strnr_dupla_adulto) = 1 then
										ckdp_adulto_b1 = "checked"
									end if%>
									<input type="checkbox" name="nr_dupla_adulto1" value="1" style="border:0;" class="inputs" <%=ckdp_adulto_b1%>>1ª dose &nbsp; 
									<input type="checkbox" name="nr_dupla_adulto2" value="2" style="border:0;" class="inputs" <%=ckdp_adulto_b2%>>2ª dose &nbsp;
									<input type="checkbox" name="nr_dupla_adulto3" value="3" style="border:0;" class="inputs" <%=ckdp_adulto_b3%>>3ª dose &nbsp; &nbsp; 
									<b>Val.:</b> <input type="text" name="dt_dupla_adulto_validade_dia" value="<%=strdt_dupla_adulto_validade_dia%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_dupla_adulto_validade_mes" value="<%=strdt_dupla_adulto_validade_mes%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_dupla_adulto_validade_ano" value="<%=strdt_dupla_adulto_validade_ano%>" size="4" maxlength="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>SCR (Sarampo, Caxumba, Rubeóla)</b></td>
								<td><input type="radio" name="nr_scr" value="1" <%if int(strnr_scr) = 1 then%>checked<%end if%> style="border:0;" class="inputs"> SIM &nbsp; &nbsp; &nbsp; 
									<input type="radio" name="nr_scr" value="" <%if int(strnr_scr) = "" OR int(strnr_scr) = 0 then%>checked<%end if%> style="border:0;" class="inputs"> Não</td>								
								<td colspan="2">&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><a name="exammed"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></a></td></tr><tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="2"><b>Exame médico realizado em: </b><input type="text" name="dt_exame_medico_dia" value="<%=strdt_exame_medico_dia%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_exame_medico_mes" value="<%=strdt_exame_medico_mes%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_exame_medico_ano" value="<%=strdt_exame_medico_ano%>" size="4" maxlength="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 42)" onFocus="PararTAB(this);"></td>
								<td colspan="2"><b>Validade: </b><%if strdt_exame_medico <> "" Then%><%=day(strdt_exame_medico)&"/"&month(strdt_exame_medico)&"/"&year(strdt_exame_medico)+1%><%end if%></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ebebeb"%>
							<tr bgcolor="#a1a1a1">
								<td style="color:white;" colspan="4"><b>CARACTERÍSTICAS FÍSICAS</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Altura</b></td>
								<td><b>Peso </b></td>
								<td><b>Manequim</b></td>
								<td><b>Calçado</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><input type="text" name="nr_caract_altura" value="<%=strnr_caract_altura%>" size="6" maxlength="6" class="inputs"></td>
								<td><input type="text" name="nr_caract_peso" value="<%=strnr_caract_peso%>" size="6" maxlength="6" class="inputs"></td>
								<td><input type="text" name="nr_caract_manequim" value="<%=strnr_caract_manequim%>" size="6" maxlength="6" class="inputs"></td>
								<td><input type="text" name="nr_caract_calcado" value="<%=strnr_caract_calcado%>" size="6" maxlength="6" class="inputs"></td>
							</tr>
							
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<%cor_campo = "#ffffff"%>
							<tr bgcolor="#a1a1a1">
								<td colspan="4" style="color:white;"><b>DEPENDENTES</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
								
									<table style="border:solid white 0px; border-collapse:collapse;">
										<tr bgcolor="<%=cor_campo%>">
											<td><b>Nome</b><br><img src="../../imagens/px.gif" alt="" width="160" height="1" border="0"></td>
											<td><b>Parentesco<br><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
											<td><b>Sexo</b><br><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
											<td><b>Data nasc.</b><br><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
											<td><b>Observações.</b><br><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
											<td>&nbsp;</td>
										</tr>
										<!--form name="formdependente"-->
										<tr bgcolor="<%=cor_campo%>">
											<td><input type="text" name="nm_dependente_nome" size="20" maxlength="30" class="inputs"></td>
											<td><select name="cd_dependente_parentesco" class="inputs">
													<option value=""></option>
													<%strsql = "Select * From TBL_parentesco "'Where cd_codigo=1 or cd_codigo = 2"
													Set rs_parentesco = dbconn.execute(strsql)%>
													<%Do While Not rs_parentesco.eof
															nm_parentesco = rs_parentesco("nm_parentesco")
															cd_dependente_parentesco = rs_parentesco("cd_codigo")
															'if cd_dependente_parentesco = cat_ then%><%'ckparent = "selected"%><%'Else%><%'ckparent=""%><%'end if%>	
															<option value="<%=cd_dependente_parentesco%>" <%=ckparent%>><%=nm_parentesco%></option>
															<%rs_parentesco.movenext
														loop
															rs_parentesco.close
															Set rs_parentesco = nothing%>						
															
												</select><img src="../../imagens/px.gif" alt="" width="12" height="1" border="0"></td>
											<td><select name="cd_dependente_sexo" class="inputs">
													<option value=""></option>
													<option value="1">Masc.</option>
													<option value="2">Fem.</option>
												</select>
												</td>
											<td><input type="text" name="dt_dependente_nascimento_dia" value="" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_dependente_nascimento_mes" value="" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_dependente_nascimento_ano" value="" size="4" maxlength="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);">&nbsp;&nbsp;</td>
											<td><input type="text" name="cd_dependente_obs" value="" size="15" class="inputs"></td>
											<td><input type="button" name="inclui_dependente" value="+" onFocus="adiciona_dependente(document.form.nm_dependente_nome.value,'',document.form.cd_dependente_parentesco.value,document.form.dt_dependente_nascimento_dia.value,document.form.dt_dependente_nascimento_mes.value,document.form.dt_dependente_nascimento_ano.value,document.form.cd_dependente_obs.value,document.form.cd_dependente_sexo.value,document.form.dependentes_total.value)"></td>
										</tr>
										<!--tr><td colspan="6"-->
											<input type="hidden" name="dependentes_total" value="<%=dependentes_total_inicial%>" size="100">
											<input type="hidden" name="dependentes_result" size="70">
										<!--/td></tr-->
										<tr bgcolor="<%=cor_campo%>"><td colspan="6"><span id='divDependente_lista'> &nbsp;</span></td></tr>
										<!--*****************-->
										<tr bgcolor="<%=cor_campo%>">
											<td colspan="6">
												<table border="1" style="border:1px solid gray; border-collapse:collapse;">
													<tr>
													<%strsql = "Select * From TBL_dependente Where cd_funcionario='"&str_funcionario&"' order by cd_parentesco,dt_nascimento,nm_nome"
														Set rs_dependente = dbconn.execute(strsql)
														
													if rs_dependente.EOF AND str_funcionario <> "" Then%>
														<td colspan="6" style="width:300px; font-weight:bold;" align="center">&nbsp;Não Possui dependentes</td>
													<%end if%>
													</tr>
													<%while not rs_dependente.EOF
														cd_dependente = rs_dependente("cd_codigo")
														nm_dependente_nome = rs_dependente("nm_nome")
														cd_dependente_parentesco = rs_dependente("cd_parentesco")
														cd_dependente_sexo = rs_dependente("cd_sexo")
														dt_dependente_nascimento = rs_dependente("dt_nascimento")
														nm_dependente_obs = rs_dependente("nm_obs")
													
													conta_dependentes = conta_dependentes + 1
													if cd_dependente_parentesco = 1 then
														conta_filhos = conta_filhos + 1
													end if%>
													<%if cabeca_dependente = "" Then%>
													
													<tr bgcolor="#808080" style="color:white;">
														<td>&nbsp;<b>&nbsp;</b><br><img src="../../imagens/px.gif" alt="" width="12" height="1" border="0"></td>
														<td>&nbsp;<b>Nome</b><br><img src="../../imagens/px.gif" alt="" width="220" height="1" border="0"></td>
														<td>&nbsp;<b>Parentesco</b><br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
														<td>&nbsp;<b>Sexo</b><br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
														<td>&nbsp;<b>Nascimento</b><br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
														<td>&nbsp;<b>Idade</b><br><img src="../../imagens/px.gif" alt="" width="65" height="1" border="0"></td>
														<td>&nbsp;<b>Observações</b><br><img src="../../imagens/px.gif" alt="" width="175" height="1" border="0"></td>											
													</tr>
													<%end if%>
													<tr>
														<td align="right"><b><%=conta_dependentes%></b></td>
														<td>&nbsp;<%=nm_dependente_nome%></td>											
														<%strsql = "Select * From TBL_parentesco Where cd_codigo="&cd_dependente_parentesco&""
														Set rs_dependente_parent = dbconn.execute(strsql)
														
														While Not rs_dependente_parent.eof
															nm_dependente_parentesco = rs_dependente_parent("nm_parentesco")
														rs_dependente_parent.movenext
														wend
														%>
														
														<td>&nbsp;<%=nm_dependente_parentesco%></td>
														<td>&nbsp;<%if cd_dependente_sexo = 1 then
																		response.write("Masc")
																	elseif cd_dependente_sexo = 2 then
																		response.write("Fem")
																	end if%></td>
														<td align="center">&nbsp;<%=dt_dependente_nascimento%></td>
													<%idade = datediff("m",dt_dependente_nascimento,now)
														if idade >= 12 then
															if idade = 12 then
																tempo = "ano"
															else
																tempo = "anos"
															end if
															idade = int(idade/12)&" "&tempo
														else 
															if idade <= 1 then
																tempo = "mes"
															else
																tempo = "meses"
															end if
															idade = idade&" "&tempo
														end if%>
														<td align="right">&nbsp;<%=idade%></td>
														<td>&nbsp;<%=nm_dependente_obs%></td>
														<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsDependenteDelete('<%=str_funcionario%>','<%=cd_dependente%>')"></td>
													</tr>
													<%cabeca_dependente = cabeca_dependente + 1
													rs_dependente.movenext
													wend
													if cd_dependente <> "" then%>
													<tr><td colspan="7" bgcolor="gray" style="color:white;"><b>Quantidade de filhos: <%=conta_filhos%></b></td></tr>
													<tr><td colspan="7" bgcolor="#c5c5c5" style="color:#f8f8f8;">Quantidade Total de dependentes: <b><%=conta_dependentes%></b></td></tr>
													<%end if%>
												</table>
											</td>
										</tr>
									</table>
									
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ebebeb"%>
							<tr bgcolor="#a1a1a1" >
								<td colspan="4" style="color:white;"><b>ESCOLARIDADE / FORMAÇÃO</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4"><b>Escolaridade</b><img src="../../imagens/px.gif" alt="" width="65" height="1" border="0">
								<b>Instituição</b><img src="../../imagens/px.gif" alt="" width="85" height="1" border="0">
								<b>Curso</b><img src="../../imagens/px.gif" alt="" width="105" height="1" border="0"> 
								<b>andamento</b><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"> 
								<b>Término</b><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0">
								<b>Observação</b></td>
							</tr>
							<!--tr><td--><input type="hidden" name="escolaridade_total" value="" size="100">
							<input type="hidden" name="escolaridade_result" size="70"><!--/td></tr-->
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">								
									<select name="cd_ensino_grau" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_escolaridade_grau order by cd_codigo"
										Set rs_grau = dbconn.execute(xsql)
										while not rs_grau.EOF
										cd_grau = rs_grau("cd_codigo")
										nm_grau = rs_grau("nm_grau")%>
											<option value="<%=cd_grau%>"><%=nm_grau%></option>
										<%rs_grau.movenext
										Wend%>
									</select> &nbsp; 
								<input type="text" name="nm_instituicao" class="inputs" size="17">
								<input type="text" name="nm_curso" class="inputs" size="17">
									<select name="cd_ensino_andamento" class="inputs">
										<option value="0"></option>
										<%xsql = "Select * From TBL_escolaridade_andamento order by cd_codigo DESC"
										Set rs_andamento = dbconn.execute(xsql)
										while not rs_andamento.EOF
										cd_andamento = rs_andamento("cd_codigo")
										nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
											<option value="<%=cd_andamento%>"><%=nm_ensino_andamento%></option>
										<%rs_andamento.movenext
										Wend%>
									</select>
									<input type="text" name="dt_termino_dia" maxlength="2" size="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_termino_mes" maxlength="2" size="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_termino_ano" maxlength="4" size="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);">
									<input type="text" name="nm_obs" size="20" maxlength="200" class="inputs">
									<input type="button" name="inclui_escolaridade" value="+" onFocus="adiciona_escolaridade(document.form.cd_ensino_grau.value,'',document.form.nm_instituicao.value, document.form.nm_curso.value, document.form.cd_ensino_andamento.value, document.form.dt_termino_dia.value,document.form.dt_termino_mes.value, document.form.dt_termino_ano.value, document.form.nm_obs.value, document.form.escolaridade_total.value)"><br>
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><span id='divEscolaridade_lista'> &nbsp;</span></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_escolaridade Where cd_funcionario='"&str_funcionario&"' order by cd_ensino_grau"
											Set rs_escolaridade = dbconn.execute(strsql)
											
										'if rs_escolaridade.EOF Then%>
											<!--td colspan="100%" style="width:300px; font-weight:bold;" align="center">&nbsp;</td-->
										<%'end if
										while not rs_escolaridade.EOF
											cd_escolaridade = rs_escolaridade("cd_codigo")
											cd_escolaridade_grau = rs_escolaridade("cd_ensino_grau")
											cd_escolaridade_andamento = rs_escolaridade("cd_ensino_andamento")
											nm_escolaridade_instituicao = rs_escolaridade("nm_instituicao")
											nm_escolaridade_curso = rs_escolaridade("nm_curso")
											nm_escolaridade_obs = rs_escolaridade("nm_obs")
											dt_escolaridade_termino = rs_escolaridade("dt_termino")
										
										'conta_dependentes = conta_dependentes + 1
										'if cd_dependente_parentesco = 1 then
										'	conta_filhos = conta_filhos + 1
										'end if%>
										<%if cabeca_escolaridade = "" Then%>
										<tr bgcolor="#808080" style="color:white;">
											<td>&nbsp;<b>Grau</b><br><img src="../../imagens/px.gif" width="65" height="1"></td>
											<td>&nbsp;<b>Instituição</b><br><img src="../../imagens/px.gif" width="190" height="1"></td>
											<td>&nbsp;<b>Curso</b><br><img src="../../imagens/px.gif" width="190" height="1"></td>
											<td>&nbsp;<b>Andamento</b><br><img src="../../imagens/px.gif" width="80" height="1"></td>
											<td>&nbsp;<b>Término</b><br><img src="../../imagens/px.gif" width="65" height="1"></td>
											<td>&nbsp;<b>Observação</b><br><img src="../../imagens/px.gif" width="190" height="1"></td>
											<td>&nbsp;</td>
										</tr>
										<%end if%>
										<tr>
											<td>&nbsp;<%xsql = "Select * From TBL_escolaridade_grau where cd_codigo='"&cd_escolaridade_grau&"'"
															Set rs_grau = dbconn.execute(xsql)
															if not rs_grau.EOF then
															cd_grau = rs_grau("cd_codigo")
															nm_grau = rs_grau("nm_grau")%>
																<%=nm_grau%>
															<%rs_grau.movenext
															end if%></td>											
											<td>&nbsp;<%=nm_escolaridade_instituicao%></td>
											<td>&nbsp;<%=nm_escolaridade_curso%></td>
											<td>&nbsp;<%xsql = "Select * From TBL_escolaridade_andamento where cd_codigo='"&cd_escolaridade_andamento&"'"
															Set rs_andamento = dbconn.execute(xsql)
															if not rs_andamento.EOF then
															cd_andamento = rs_andamento("cd_codigo")
															nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
																<%=nm_ensino_andamento%>
															<%rs_andamento.movenext
															end if%></td>
											<td>&nbsp;<%=dt_escolaridade_termino%></td>
											<td>&nbsp;<%=nm_escolaridade_obs%></td>					
											<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsEscolaridadeDelete('<%=str_funcionario%>','<%=cd_escolaridade%>')"></td>
										</tr>
										<%cabeca_escolaridade = cabeca_escolaridade + 1
										rs_escolaridade.movenext
										wend%>
									</table>
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ffffff"%>
							<tr bgcolor="#a1a1a1" >
								<td colspan="4" style="color:white;"><b>EXPERIÊNCIA PROFISSIONAL</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<b>Empresa Anterior</b><img src="../../imagens/px.gif" alt="" width="75" height="1" border="0">
									<b>Função</b><img src="../../imagens/px.gif" alt="" width="125" height="1" border="0">
									<b>entrada</b><img src="../../imagens/px.gif" alt="" width="105" height="1" border="0">
									<b>Saída</b><img src="../../imagens/px.gif" alt="" width="115" height="1" border="0">
									<b>Obs.</b></td>
							</tr>
							<!--tr><td colspan="4"--><input type="hidden" name="emprego_total" value="" size="100">
								<input type="hidden" name="emprego_result" size="70"><!--/td></tr-->
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<input type="text" name="nm_emprego_empresa" class="inputs" maxlength="200" size="20"><img src="../../imagens/px.gif" alt="" width="10" height="1" border="0">
									<input type="text" name="nm_emprego_cargo" class="inputs" maxlength="200" size="20"><img src="../../imagens/px.gif" alt="" width="10" height="1" border="0">
									<input type="text" name="dt_emprego_entrada_dia" maxlength="2" size="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_emprego_entrada_mes" maxlength="2" size="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_emprego_entrada_ano" maxlength="4" size="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this,4)" onFocus="PararTAB(this);"><img src="../../imagens/px.gif" alt="" width="15" height="1" border="0">
									<input type="text" name="dt_emprego_saida_dia" maxlength="2" size="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_emprego_saida_mes" maxlength="2" size="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_emprego_saida_ano" maxlength="4" size="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"><img src="../../imagens/px.gif" alt="" width="15" height="1" border="0">
									<input type="text" name="dt_emprego_obs" size="15" maxlength="200" class="inputs">
									<input type="button" name="inclui_emprego" value="+" onFocus="adiciona_emprego(document.form.nm_emprego_empresa.value,'',document.form.nm_emprego_cargo.value, document.form.dt_emprego_entrada_dia.value, document.form.dt_emprego_entrada_mes.value, document.form.dt_emprego_entrada_ano.value, document.form.dt_emprego_saida_dia.value, document.form.dt_emprego_saida_mes.value, document.form.dt_emprego_saida_ano.value, document.form.dt_emprego_obs.value, document.form.emprego_total.value)">
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><span id='divEmprego_lista'> &nbsp;</span></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_emprego_anterior Where cd_funcionario='"&str_funcionario&"' order by dt_entrada"
											Set rs_emprego = dbconn.execute(strsql)
											
										'if rs_escolaridade.EOF Then%>
											<!--td colspan="100%" style="width:300px; font-weight:bold;" align="center">&nbsp;</td-->
										<%'end if
										while not rs_emprego.EOF
											
											cd_emprego = rs_emprego("cd_codigo")
											nm_emprego_empresa = rs_emprego("nm_empresa")
											nm_emprego_cargo = rs_emprego("nm_cargo")
											dt_emprego_entrada = rs_emprego("dt_entrada")
											dt_emprego_saida = rs_emprego("dt_saida")
											nm_emprego_obs = rs_emprego("nm_obs")
											
											conta_linha = conta_linha + 1%>
										<%if cabeca_emprego = "" Then%>
										<tr bgcolor="#808080" style="color:white;">
											<td>&nbsp;</td>
											<td>&nbsp;<b>Emprego Anterior</b><br><img src="../../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
											<td>&nbsp;<b>Função</b><br><img src="../../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
											<td>&nbsp;<b>Data Entrada</b><br><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
											<td>&nbsp;<b>Data saída</b><br><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
											<td>&nbsp;<b>Observação</b><br><img src="../../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
											<td>&nbsp;</td>
										</tr>
										<%end if%>
										<tr>
											<td>&nbsp;<%=conta_linha%></td>
											<td>&nbsp;<%=nm_emprego_empresa%></td>											
											<td>&nbsp;<%=nm_emprego_cargo%></td>
											<td>&nbsp;<%=dt_emprego_entrada%></td>
											<td>&nbsp;<%=dt_emprego_saida%></td>
											<td>&nbsp;<%=nm_emprego_obs%></td>
											<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsEmpregodeDelete('<%=str_funcionario%>','<%=cd_emprego%>')"></td>
										</tr>
										<%cabeca_emprego = cabeca_emprego + 1
										rs_emprego.movenext
										wend%>
									</table>
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ebebeb"%>
							<tr bgcolor="#a1a1a1" >
								<td colspan="4" style="color:white;"><b>INFORMAÇÕES COMPLEMENTARES</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Possui parente na Empresa? </b><%if strnm_parente <> "" AND str_funcionario <> "" Then%>SIM<%Elseif strnm_parente = "" AND str_funcionario <> "" then%>NÃO<%end if%></td>
								<td colspan="2"><b>Nome: </b><input type="text" name="nm_parente" value="<%=strnm_parente%>" size="50" maxlength="50" class="inputs"></td>
								<td><b>Parentesco: </b>
									<select name="cd_parente_parentesco" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_parentesco order by nm_parentesco DESC"
										Set rs_parentesco = dbconn.execute(xsql)
										while not rs_parentesco.EOF
										cd_parentesco = rs_parentesco("cd_codigo")
										nm_parentesco = rs_parentesco("nm_parentesco")%>
											<option value="<%=cd_parentesco%>" <%if int(strcd_parente_parentesco) = cd_parentesco then%> SELECTED<%end if%>><%=nm_parentesco%></option>
										<%rs_parentesco.movenext
										Wend%>
									</select>
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Residência: (tipo)</b></td>
								<!--td><b>Financiada?</b></td-->
								<td><b>Reside</b></td>
								<td><b>Veículo: (tipo)</b></td>
								<!--td>&nbsp;</td-->
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><select name="cd_residencia_tipo" class="inputs">
										<option value="0"></option>
										<option value="1" <%if int(strcd_residencia_tipo) = 1 Then%>selected<%end if%>><b>Própria<b></option>
										<option value="2" <%if int(strcd_residencia_tipo) = 2 Then%>selected<%end if%>><b>Alugada<b></option>
									</select>
								<!--td><select name="cd_residencia_financ" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_residencia_financ) = 1 Then%>selected<%end if%>>Sim</option>
										<option value="2" <%if int(strcd_residencia_financ) = 2 Then%>selected<%end if%>>Não</option>
									</select></td-->
								<td><select name="cd_residencia_reside" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_residencia_reside) = 1 Then%>selected<%end if%>>Só</option>
										<option value="2" <%if int(strcd_residencia_reside) = 2 Then%>selected<%end if%>>Companheiro</option>
										<option value="3" <%if int(strcd_residencia_reside) = 3 Then%>selected<%end if%>>Pais</option>
										<option value="4" <%if int(strcd_residencia_reside) = 4 Then%>selected<%end if%>>Parentes</option>
										<option value="5" <%if int(strcd_residencia_reside) = 5 Then%>selected<%end if%>>Amigos</option>
									</select></td>
								<td><select name="cd_veiculo_tipo" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_veiculo_tipo) = 1 Then%>selected<%end if%>>Carro</option>
										<option value="2" <%if int(strcd_veiculo_tipo) = 2 Then%>selected<%end if%>>Moto</option>
									</select>				
								</td>
								<!--td>&nbsp;</td-->
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<!--tr bgcolor="<%=cor_campo%>">
								
								<td><b>Financiado?</b></td>
								<td colspan="2">&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								
								<td><select name="cd_veiculo_financ" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_veiculo_financ) = 1 Then%>selected<%end if%>>Sim</option>
										<option value="2" <%if int(strcd_veiculo_financ) = 2 Then%>selected<%end if%>>Não</option>
									</select></td>
								<td colspan="2">&nbsp;</td>
							</tr-->
							<input type="hidden" name="valores_total" value="" size="100">
							<input type="hidden" name="valores_result" size="70">
							<!--tr bgcolor="<%=cor_campo%>"><td colspan="4"><a name="beneficios"></a><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
							
							<%'if session_cd_usuario = 3 or session_cd_usuario = 52 or session_cd_usuario = 56 or session_cd_usuario = 46 then%>
<%cor_campo = "#fffff"%>
							<tr bgcolor="#a1a1a1" >
								<%if int(cd_user) <> usuario_restrito then%>
								<td colspan="2" style="color:white;"><b>SALÁRIOS</b></td><%end if%>
								<td colspan="2" style="color:white;"><b>TRANSPORTE</b><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td>
								<!--td colspan="1" style="color:white;"><b>REFEIÇÃO</b></td-->
							</tr>
							<tr bgcolor="#a1a1a1" >
								<%if int(cd_user) <> usuario_restrito then%><td colspan="2" style="color:white;"></td><%end if%>
								<td align="left" colspan="2" style="color:white;">
									<%strsql = "SELECT * FROM View_Beneficios WHERE cd_beneficio_tipo = 2 ORDER BY nm_beneficio"
									SET rs_ben = dbconn.execute(strsql)%>
									<select name="cd_transporte">
										<option value="0" SELECTED></option>
											<%while not rs_ben.EOF
												cd_beneficio = rs_ben("cd_codigo")
												nm_beneficio_abrev = rs_ben("nm_beneficio_abrev")
												nm_categoria = rs_ben("nm_categoria")
												nm_regiao = rs_ben("nm_regiao")
													strsql = "SELECT * FROM TBL_beneficios_valores WHERE cd_beneficio = '"&cd_beneficio&"' ORDER BY dt_atualizacao DESC"
													SET rs_valor = dbconn.execute(strsql)
														if NOT rs_valor.EOF then nr_transp_atual = rs_valor("nr_valor")%>
												<option value="<%=cd_beneficio%>"><%=nm_beneficio_abrev&" - "&nm_categoria&": "&formatnumber(nr_transp_atual)%></option>
											<%nr_transp_atual = "0"
											rs_ben.movenext
											wend%>
									</select>
									Atualiz.:<input type="text" name="mes_transp_atualizacao" size="2" maxlength="2">/<input type="text" name="ano_transp_atualizacao" size="3" maxlength="4"></td>
								<!--td align="center" colspan="1" style="color:white;">
									<%'strsql = "SELECT * FROM View_Beneficios WHERE cd_beneficio_tipo = 1 ORDER BY nm_beneficio"
									'SET rs_ben = dbconn.execute(strsql)%>
									<select name="cd_refeicao">
										<option value="0" SELECTED></option>
											<%'while not rs_ben.EOF
												'cd_beneficio = rs_ben("cd_codigo")
												'nm_beneficio_abrev = rs_ben("nm_beneficio_abrev")
												'nm_categoria = rs_ben("nm_categoria")%>
												<option value="<%'=cd_beneficio%>"><%'=nm_beneficio_abrev&" - "&nm_categoria%></option>
											<%'rs_ben.movenext
											'wend%>
									</select><br>
									Atualização:<input type="text" name="mes_refeic_atualizacao" size="2" maxlength="2">/<input type="text" name="ano_refeic_atualizacao" size="4" maxlength="4"></td-->
							</tr>
							<!--tr bgcolor="<%=cor_campo%>"><td colspan="4"><span id='divValores_lista'> &nbsp;</span></td></tr-->
							<tr>
								<%if int(cd_user) <> usuario_restrito then%>
								<td valign="top" colspan="2">
								
									<table border="1" style="border-collapse:collapse;">
										<tr bgcolor="#eaeaea">
											<td align="center"><b>Últimos Salários<%'=cd_user&" <> "&usuario_restrito%></b><br><img src="../../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
											<td align="center"><b>Atualização</b><br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
											<td align="center"><b>% ant.</b><br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
											<td align="center"><b>Obs.</b><br><img src="../../imagens/px.gif" alt="" width="105" height="1" border="0"></td>
										</tr>
										<%'conta_sal = 1
										strsql = "Select top 3 * From TBL_funcionario_valores Where cd_funcionario='"&str_funcionario&"' AND cd_tipo=1 order by dt_atualizacao desc"
										Set rs_valor = dbconn.execute(strsql)
										
											while not rs_valor.EOF
												
												conta_sal = conta_sal + 1
												
												cd_valor = rs_valor("cd_codigo")
												cd_valor_tipo = rs_valor("cd_tipo")
												nr_valor = rs_valor("nr_valor")
												cd_funcionario = rs_valor("cd_funcionario")
												dt_valor_atualizacao = rs_valor("dt_atualizacao")
												nm_valor_obs = rs_valor("nm_obs")
													
													if conta_sal = 1 then
														ultimo_sal = nr_valor
														ultimo_valor_atualizacao = dt_valor_atualizacao
														
													elseif conta_sal = 2 then
														penultimo_sal = nr_valor
														penultimo_valor_atualizacao = dt_valor_atualizacao
													
													elseif conta_sal = 3 then
														response.write(nm_valor)
														antepenultimo_sal = nr_valor
														antepenultimo_valor_atualizacao = dt_valor_atualizacao
													end if%>
											
											<%ultimo_salario = nr_valor
											
											rs_valor.movenext
											wend%>
											<%for i = 1 to conta_sal
												if i = 1 then 
													cor_salario = "#008000"
												else
													cor_salario = "silver"
												end if%>
											<tr style="color:<%=cor_salario%>">
												<td align="right" <%if i = 1 then%> onClick="window.open('funcionarios_folhapagamento_salario.asp?tipo=cadastro&cod=<%=cd_funcionario%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','','width=470,height=200,scrollbars=1')"<%end if%>><span><b>
														<%if i = 1 Then%>
															<%=formatcurrency(ultimo_sal)%>
														<%elseif i = 2 Then%>
															<%=formatcurrency(penultimo_sal)%>
														<%elseif i = 3 then%>
															<%=formatcurrency(antepenultimo_sal)%>
														<%end if%></b></span></td>
														
												<td align="right"><b><%if i = 1 Then%>
															<%=ultimo_valor_atualizacao%>
														<%elseif i = 2 Then%>
															<%=penultimo_valor_atualizacao%>
														<%elseif i = 3 then%>
															<%=antepenultimo_valor_atualizacao%>
														<%end if%></b></td>
														
												<td align="right"><b>
													<%if i = 1 Then
														dif_sal = ultimo_sal - penultimo_sal
														variacao_sal = (dif_sal / ultimo_sal)*100%>
														<%=formatnumber(variacao_sal,2)&"%"%>
													<%elseif i = 2 Then
														dif_sal = penultimo_sal - antepenultimo_sal
														variacao_sal = (dif_sal / penultimo_sal)*100%>
														<%=formatnumber(variacao_sal,2)&"%"%>
													<%elseif i = 3 then%>
														<%=" - "%>
													<%end if%></b>
												<td>&nbsp;<b><%=nm_valor_obs%></b></td>																					
											</tr>
											<%dif_sal = 0
											variacao_sal = 0
											next%>
									</table>
								
								</td><%end if%>
					<!-- ########## T R A N S P O R T E #####################-->
								<td colspan="2" valign="top">
									<table border="1" style="border-collapse:collapse;">
										<tr bgcolor="#eaeaea">
											<td align="center"><b>Tipo</b><br><img src="../../imagens/px.gif" width="130" height="1"></td>
											<td align="center"><b>Valor</b><br><img src="../../imagens/px.gif" width="50" height="1"></td>
											<td align="center"><b>Região</b><br><img src="../../imagens/px.gif" width="100" height="1"></td>
											<td align="center"><b>Desde</b><br><img src="../../imagens/px.gif" width="60" height="1"></td>
										</tr>
										<%cor = 1
										strsql = "SELECT * FROM TBL_funcionario_beneficios where cd_funcionario = '"&str_funcionario&"' AND cd_beneficio_tipo = 2 order by dt_atualizacao desc"
										SET rs_benef = dbconn.execute(strsql)
										while not rs_benef.EOF
											cd_beneficio_transp = rs_benef("cd_beneficio")
										
											
											strsql = "SELECT top 3 * FROM View_Beneficios_valores WHERE cd_codigo="&cd_beneficio_transp&" ORDER BY dt_atualizacao DESC"
											SET rs_transp = dbconn.execute(strsql)
											'while not rs_transp.EOF
											if not rs_transp.EOF then
												nm_beneficio = rs_transp("nm_beneficio")
												nm_categoria = rs_transp("nm_categoria")
												nr_valor = rs_transp("nr_valor")
												dt_atualizacao = rs_transp("dt_atualizacao")
												nm_regiao = rs_transp("nm_regiao")
													'if cor = 1 then
														'color_transp = "#008000"
														color_transp = "#000000"
													'	cor = cor + 1
													'else
													'	color_transp = "silver"
													'end if%>
											<tr style="color:<%=color_transp%>;">
												<td align="left"><%=nm_categoria%></td>
												<td align="right"><%if nr_valor <> "" Then response.write(formatcurrency(nr_valor*2))%></td>
												<td align="center"><%=nm_regiao%></td>
												<td align="center"><%=dt_atualizacao%></td>
											</tr>
											<%'rs_transp.movenext
											'wend
											end if%>
										<%rs_benef.movenext
										wend%>
									</table>
								</td>
					<!-- ########## R E F E I Ç Ã O #####################-->								
								<!--td valign="top">
									<table border="1" style="border-collapse:collapse;"-->
										<%'strsql = "SELECT * FROM TBL_funcionario_beneficios where cd_funcionario = '"&str_funcionario&"' AND cd_beneficio_tipo = 1"
										'SET rs_benef = dbconn.execute(strsql)
										'while not rs_benef.EOF
										'	cd_beneficio_refeic = rs_benef("cd_beneficio")
										'rs_benef.movenext
										'wend
										%>
										<!--tr bgcolor="#eaeaea">
											<td align="center"><b>Tipo</b><br><img src="../../imagens/px.gif" width="130" height="1"></td>
											<td align="center"><b>Valor</b><br><img src="../../imagens/px.gif" width="50" height="1"></td-->
											<!--td align="center"><b>Atual.</b><br><img src="../../imagens/px.gif" width="50" height="1"></td-->
										<!--/tr-->
										<%'if cd_beneficio_refeic <> "" Then
										
										'	cor = 1
										'	strsql = "SELECT top 3 * FROM View_Beneficios_valores WHERE cd_codigo="&cd_beneficio_refeic&" ORDER BY dt_atualizacao DESC"
										'	SET rs_transp = dbconn.execute(strsql)
										'	while not rs_transp.EOF
										'		nm_beneficio = rs_transp("nm_beneficio")
										'		nr_valor = rs_transp("nr_valor")
										'		dt_atualizacao = rs_transp("dt_atualizacao")
										'			if cor = 1 then
										'				color_transp = "#008000"
										'				cor = cor + 1
										'			else
										'				color_transp = "silver"
										'			end if%>
											<!--tr style="color:<%'=color_transp%>;">
												<td align="left"><%'=nm_beneficio%></td>
												<td align="right"><%'if nr_valor <> "" Then response.write(formatcurrency(nr_valor))%></td-->
												<!--td align="right"><%'=dt_atualizacao%></td-->
											<!--/tr-->
											<%'rs_transp.movenext
											'wend
										'else%>
											<!--tr><td colspan="3">Não há vale refeição</td></tr-->
										<%'end if%>
									<!--/table-->
								</td>
							</tr>
							
								
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<%'end if%>	
<%cor_campo = "#ffffff"%>
							<tr bgcolor="#a1a1a1" >
								<td colspan="4" style="color:white;"><b>BENEFÍCIOS</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Cartão SPTRANS</b></td>
								<td><b>Cartão BOM</b></td>
								<td><b>Vale Refeição</b></td>
								<td><b>Cesta Básica</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><input type="text" name="cd_sptrans" value="<%=str_sptrans%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_cmt_bom" value="<%=str_cmt_bom%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_vr" value="<%=str_vr%>" size="25" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_vale_alimentacao" value="<%=str_vale_alimentacao%>" size="25" maxlength="20" class="inputs"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Banco</b></td>
								<td><b>Agência</b></td>
								<td><b>Conta</b></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><%if str_banco = "" Then%><%str_banco = "Real"%><%end if%>
								<%if str_banco_ag = "" Then%><%str_banco_ag = "1342"%><%end if%>
								<td><input type="text" name="nm_banco" value="<%=str_banco%>" size="20" maxlength="30" class="inputs"></td>
								<td><input type="text" name="cd_banco_ag" value="<%=str_banco_ag%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_banco_conta" value="<%=str_banco_conta%>" size="20" maxlength="20" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td>&nbsp;<b>Assistência Médica</b></td>
								<td>&nbsp;<b>Co-participação</b></td>
								<td>&nbsp;<b>Assistência Odontológica</b></td>
								<td>&nbsp;<b>Co-participação</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><input type="text" name="cd_assistencia_medica_matricula" value="<%=str_assistencia_medica_matricula%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="nr_assistencia_medica_coparticipacao" size="3" maxlength="3" value="<%=str_assistencia_medica_coparticipacao%>" class="inputs">%</td>
								<td><input type="text" name="cd_assistencia_odonto_matricula" value="<%=str_assistencia_odonto_matricula%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="nr_assistencia_odonto_coparticipacao" size="3" maxlength="3" value="<%=str_assistencia_odonto_coparticipacao%>" class="inputs">%</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<%cor_campo = "#ebebeb"%>
							<%if strcod = "" Then%>
							<tr bgcolor="#a1a1a1" >
								<td colspan="4" style="color:white;"><b>CONTRATO</b></td>
							</tr>							
							<tr bgcolor="<%=cor_campo%>">
								<td style="color:red;"><b>Contrato</b> &nbsp;</td>
								<td style="color:red;"><b>Data de admissão</b></td>
									<%If strcod <> "" Then%>
										<td><b>Data de recisão</b></td>
										<td>&nbsp;</td>
									<%else%>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
									<%End If%>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><%strsql = "Select * From TBL_tipo_contrato "'Where cd_codigo=1 or cd_codigo = 2"
			  					Set rs_contr = dbconn.execute(strsql)%>
								<select name="cd_regime_trabalho" class="inputs">
									<option value="">Selecione</option>
									<%Do While Not rs_contr.eof
									nm_regime_trab = rs_contr("nm_regime_trab")
									cd_cod = rs_contr("cd_codigo")
									if cd_cod = strregime_trabalho then%><%ckcontr = "selected"%><%Else%><%ckcontr=""%><%end if%>	
										<option value="<%=cd_cod%>" <%=ckcontr%>><%=cd_cod%>-<%=nm_regime_trab%></option>
									<%rs_contr.movenext
									loop
									rs_contr.close
									Set rs_contr = nothing%>
								</select> </td>
								<td><!--input type="text" name="dt_contratado" value="<%'=Strdt_contratado%>" size="20" maxlength="10" class="inputs"-->
									<input type="text" name="cd_dia" value="" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
									<input type="text" name="cd_mes" value="" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
									<input type="text" name="cd_ano" value="" size="4" maxlength="4" class="inputs"></td>								
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<%else%>						
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Contrato</b> &nbsp;</td>
								<td style="color:red;"><b>Data de admissão</b></td>
								<%If strcod <> "" Then%>
									<td><b>Data de recisão</b></td>
									<td>&nbsp;</td>
								<%else%>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
								<%End If%>
							</tr>
							
								<%'strsql = "Select * From VIEW_funcionario_contrato_lista Where cd_funcionario = "&str_funcionario&" order by dt_admissao desc"
			  					strsql = "Select * From VIEW_funcionario_contrato_lista Where cd_funcionario = "&strcod&" order by dt_admissao desc"
			  					Set rs_contrato = dbconn.execute(strsql)%>
								<%while not rs_contrato.EOF
									cd_funcionario = rs_contrato("cd_funcionario")
									
									nm_regime_trab = rs_contrato("nm_regime_trab")
									dt_admissao = rs_contrato("dt_admissao")
									dt_demissao = rs_contrato("dt_demissao")%>
								<tr bgcolor="<%=cor_campo%>">
									<td>&nbsp;<a href="javascript:void();" onclick="window.open('../../empresa/janelas/funcionarios_contrato_altera.asp?cod_contrato=<%=cod_contrato%>&cd_funcionario=<%=str_funcionario%>', 'janela_vdlap', 'width=335, height=205'); return false;"><%=nm_regime_trab%></a></td>
									<td>&nbsp;<%=zero(day(dt_admissao))%>/<%=mesdoano(month(dt_admissao))%>/<%=year(dt_admissao)%></td>
									<td>&nbsp;<%=dt_demissao%></td>
									<td>&nbsp;&nbsp;
									<%if dt_demissao <> "" then
										if nr_contrato = "" then%>
											<a href="javascript:;" onclick="window.open('../../empresa/janelas/funcionarios_contrato.asp?cd_funcionario=<%=str_funcionario%>', 'janela_vdlap', 'width=350, height=400'); return false;"><b>*** Novo Contrato ***<!--img src="../../imagens/ic_aprovado.gif" alt="Novo Contrato." width="10" height="12" border="0" --></b></a>
										<%end if%>
									<%else%>
											<a href="javascript:;" onclick="window.open('../../empresa/janelas/funcionarios_recisao.asp?cod_contrato=<%=cod_contrato%>&cd_funcionario=<%=cd_funcionario%>','principal','width=370, height=400'); return false;">Rescindir contrato?</a>&nbsp; <!--day(dt_demissao) = "" then%-->
									<%end if%></td>
								</tr>
								<%nr_contrato = nr_contrato + 1
								rs_contrato.movenext
								wend%>
							
							<%end if%>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="2" style="color:red;"><b>Cargo</b></td>
								<td style="color:red;"><b>Situação</b></td>
								<td style="color:red;"><b>Unidade</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<%strsql = "TBL_cargos where cd_especificidade = 3"
								Set rs_func = dbconn.execute(strsql)%>
								<%if strcod = "" Then%>
								<td colspan="2">
									<select name="cd_cargo" class="inputs"> 
									<option value="">Selecione</option>
									<%Do While Not rs_func.eof%>
									<option value="<%=rs_func("cd_codigo")%>"<%if strcd_cargo = CInt(rs_func("cd_codigo")) then%>selected<%End if%>><%=rs_func("nm_cargo")%></option>
									<%rs_func.movenext
									loop
									rs_func.close
									Set rs_func = nothing%>
									</select> *
								<%Else
									xsql_cargo = "select * From View_funcionario_cargo Where cd_funcionario='"&strcod&"'"
									Set rs_cargo = dbconn.execute(xsql_cargo)
									if not rs_cargo.EOF Then
									nm_cargo = rs_cargo("nm_cargo")
								End if%>
								
								<%strsql = "select * from VIEW_funcionario_cargo where cd_funcionario='"&strcod&"' and cd_suspenso <> 1 order by dt_atualizacao desc"
								Set rs_cargo = dbconn.execute(strsql)%>
								<%if not rs_cargo.eof then%><%nm_cargo=rs_cargo("nm_cargo")%><%end if%>
								<td class="txt_cinza" colspan="2" <%=tipo%> style="background: #c7c7c7;"><a href="javascript:;" onclick="window.open('../../empresa/janelas/funcionarios_cargo.asp?cd_funcionario=<%=strcod%>&cd_funcao=<%=cd_funcao%>', 'janela_vdlap', 'width=350, height=400'); return false;"><%=nm_cargo%>.</a>
																
								<%end if%></td>								
								<td style="background: #c7c7c7;">
								<%if strcod = "" Then%>
									<%strsql ="up_ADM_Status_lista"
									Set rs_cat = dbconn.execute(strsql)%>
								
									<select name="cd_status" class="inputs"> 
									<option value="">Selecione</option>
									<%Do While Not rs_cat.eof
									cd_cod_status = rs_cat("cd_codigo")
									%>
									<option value="<%=cd_cod_status%>"><%=cd_cod_status%>-<%=rs_cat("nm_status")%></option>
									<%rs_cat.movenext
									loop%>
									</select>
								<%else%>
								<%strsql ="SELECT * From View_funcionario_status where cd_funcionario = "&strcod&"  and cd_suspenso <> 1 order by dt_atualizacao"
								  	Set rs_status = dbconn.execute(strsql)
									
									do while not rs_status.EOF
									cd_status = rs_status("cd_status")
									nm_status = rs_status("nm_status")
									dt_atualizacao = rs_status("dt_atualizacao")
									
									rs_status.movenext
									loop%>
									<a href="javascript:;" onclick="window.open('../../empresa/janelas/funcionarios_status.asp?cd_funcionario=<%=strcod%>&cd_status=<%=cd_status%>', 'janela_vdlap', 'width=350, height=400'); return false;"><%=nm_status%>.</a>
								<%end if%>
									</td>
									<td style="background: #c7c7c7;">
									<%strsql ="TBL_unidades where cd_hospital > 0 and cd_status = 1"
								  	Set rs_uni = dbconn.execute(strsql)%>
								<%if strcod = "" Then%>
									<select name="cd_unidade" class="inputs"> 
										<option value="">Selecione</option>
										<%Do While Not rs_uni.eof%>
										<option value="<%=rs_uni("cd_codigo")%>"><%=rs_uni("nm_Unidade")%></option>
										<%rs_uni.movenext
										loop
										rs_uni.close
										Set rs_uni = nothing%>
									</select> 	
								<%else%>
									<%strsql ="SELECT * From View_funcionario_unidade where cd_funcionario = "&strcod&" and cd_suspenso <> 1 order by dt_atualizacao asc"
								  	Set rs_unidade = dbconn.execute(strsql)
									
									do while not rs_unidade.EOF
										cd_unidade = rs_unidade("cd_unidade")
										nm_unidade = rs_unidade("nm_unidade")
										dt_atualizacao = rs_unidade("dt_atualizacao")
									rs_unidade.movenext
									loop%>
									 <a href="javascript:;" onclick="window.open('../../empresa/janelas/funcionarios_unidade.asp?cd_funcionario=<%=strcod%>&cd_unidade=<%=cd_unidade%>', 'janela_vdlap', 'width=350, height=400'); return false;"><%=nm_unidade%>.</a>
									<%rs_uni.close
									Set rs_uni = nothing%>
								<%end if%></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4" style="color:red;"><b>Função</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<%strsql = "TBL_cargos where cd_especificidade = 4"
								Set rs_func = dbconn.execute(strsql)%>
								<%if strcod = "" Then%>
								<td colspan="2">
									<select name="cd_funcao" class="inputs"> 
									<option value="">Selecione</option>
									<%Do While Not rs_func.eof%>
									<option value="<%=rs_func("cd_codigo")%>"<%if strcd_funcao = CInt(rs_func("cd_codigo")) then%>selected<%End if%>"><%=rs_func("nm_cargo")%></option>
									<%rs_func.movenext
									loop
									rs_func.close
									Set rs_func = nothing%>
									</select> *
								<%Else
									xsql_funcao = "select * From View_funcionario_funcao Where cd_funcionario='"&strcod&"'"
									Set rs_funcao = dbconn.execute(xsql_funcao)
									if not rs_funcao.EOF Then
									nm_funcao = rs_funcao("nm_funcao")
								End if%>
								
								<%strsql = "select * from VIEW_funcionario_funcao where cd_funcionario='"&strcod&"' and cd_suspenso <> 1 order by dt_atualizacao desc"
								Set rs_funcao = dbconn.execute(strsql)%>
								<%if not rs_funcao.eof then%><%nm_funcao=rs_funcao("nm_funcao")%><%end if%>
								<td class="txt_cinza" colspan="2" <%=tipo%> style="background: #c7c7c7;"><a href="javascript:;" onclick="window.open('../../empresa/janelas/funcionarios_funcao.asp?cd_funcionario=<%=strcod%>&cd_funcao=<%=cd_funcao%>', 'janela_vdlap', 'width=350, height=400'); return false;"><%=nm_funcao%>.</a>
																
								<%end if%></td>								
								<td style="background: #c7c7c7;"></td>
									<td style="background: #c7c7c7;"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr bgcolor="<%=cor_campo%>" style="color:red;">
								<td><b>Entrada</b></td>
								<td><b>Intervalo</b></td>
								<td><b>Saída</b></td>
								<td>&nbsp;</td>
							</tr>
							<%xsql = "Select * From TBL_funcionario_horario Where cd_funcionario = '"&strcod&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
							'xsql = "SELECT  * FROM         TBL_funcionario_horario WHERE     (cd_funcionario = 97) ORDER BY dt_atualizacao DESC"
							Set rs_hora = dbconn.execute(xsql)
								if not rs_hora.EOF then
									hr_entrada = rs_hora("hr_entrada")
									hr_saida = rs_hora("hr_saida")
									nm_intervalo = rs_hora("nm_intervalo")
								end if
							  %>														
							<tr bgcolor="<%=cor_campo%>">
								<td>
									<%if strcod = "" Then%>
										<input type="text" name="hr_entrada" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:<input type="text" name="min_entrada" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">
									<%else%>
										<%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%>
									<%end if%>
								</td>
								<td>
									<%if strcod = "" Then%>
										<select name="nm_intervalo" size="1" class="inputs">
											<option value=""></option>
											<option value="1 Hora">1 Hora</option>
											<option value="15 minutos">15 minutos</option>
										</select>
									<%else%>
										<%=nm_intervalo%>
									<%end if%>
								</td>
								<td>
									<%if strcod = "" Then%>
										<input type="text" name="hr_saida" size="2" maxlength="2"  class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:<input type="text" name="min_saida" size="2" maxlength="2"  class="inputs"></td>
									<%else%>
										<%=zero(hour(hr_saida))&":"&zero(minute(hr_saida))%>
										&nbsp;<a href="javascript:;" onclick="window.open('../janelas/funcionarios_horario.asp?cd_funcionario=<%=strcod%>', 'janela_vdlap', 'width=350, height=430'); return false;"><img src="../../imagens/ic_novo.gif" alt="" width="10" height="12" border="0"></a></td>
									<%end if%>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
								<%id_origem = str_funcionario
								cd_origem = 4
								pag_retorno = ".."&mem_posicao
								variaveis = "../../empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro$cod="&str_funcionario
								pos_caminho = "../../"
								if str_funcionario <> "" then
									ocorrencias_mostrar = 1
								end if%>
								<!--#include file="../../ocorrencia/ocorrencia_formulario.asp"-->
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td colspan="4" valign="middle"><!--<A href="empresa.asp?tipo=lista"><img src="../../imagens/bt_lista.gif" alt="Listar" border="0" width="119" height="22" border="1"></a>-->
								<input type="image" src="../../imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0">
								<%=stracao%></td>
							</tr>
							<%if int(session_cd_usuario) = 46 AND strcod <> "" then%>
							<tr><td colspan="4">&nbsp;</td></tr>
							<tr>
								<td colspan="3" >&nbsp;</td>
								<td align="right">
									<!--input onclick="javascript:JsDelete('<%=strcod%>','1')" type="button" value="Apagar funcionario" title="Apagar 1"-->
									<img src="../../imagens/ic_del.gif" alt="Apaga funcionario" width="13" height="14" border="0" onclick="javascript:JsDelete('<%=strcod%>','1')">
								</td>
							</tr>
							<%end if%>
							</form>
							<%If strcod <> "" then%>
								<tr>
									<td colspan="4">
									<form name="form_ex" action="empresa/acoes/empr_funcionarios_acao.asp" id="forms" method="post"  enctype="multipart/form-data">
										<input type="hidden" name="acao" value="excluir">
										<input type="hidden" name="cod" value="<%=cd_funcionario%>">
									</form></td>
								</tr>
							<%End if%>
						</table>
									 								
					<%End if%>
			<SCRIPT LANGUAGE="javascript">
 {
        //MaskInput(document.forms.nm_cep, "99999-99");
		//MaskInput(document.forms.dt_nascimento, "99/99/9999");	    
 }
</SCRIPT>
			
			<SCRIPT LANGUAGE="javascript">
				function JsDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o funcionario?"))
				  		{
							  //document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir&protege_exclusao='+cod2+'');
							  window.location=('../acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir&protege_exclusao='+cod2+'');
							  //window.location=("www.google.com.br");
						}
					}
					
				
				function JsContatoDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o contato?"))
				  		{
				  		//document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_cont='+cod2+'&acao=apaga_contato');
				  		document.location=('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_cont='+cod2+'&acao=apaga_contato');
						}
					}  
					  
				function JsDependenteDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o dependente?"))
				  		{
				 		//document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_dep='+cod2+'&acao=apaga_dependente');
				  		document.location=('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_dep='+cod2+'&acao=apaga_dependente');
						}
					}  
				
				function JsEscolaridadeDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse registro de escolaridade?"))
				  		{
				  		//document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_esc='+cod2+'&acao=apaga_escolaridade');
						document.location=('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_esc='+cod2+'&acao=apaga_escolaridade');
						}
					}    
				
				function JsEmpregodeDelete(cod,cod2)
					{
					if (confirm ("Tem certeza que deseja excluir esse registro de emprego anterior?"))
				  		{
				  		//document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_emp='+cod2+'&acao=apaga_emprego');
						document.location=('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_emp='+cod2+'&acao=apaga_emprego');
						}
					} 
				
				function JsValorDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse valor?"))
				  {
				  document.location=('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_valor='+cod2+'&acao=apaga_valor');
					}
					  } 
				
					  
				//function JsContatoDelete(cod){
					//if (confirm ("Tem certeza que deseja excluir o contato?")){
				  //document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir');
					//}
					  //}
			</SCRIPT>