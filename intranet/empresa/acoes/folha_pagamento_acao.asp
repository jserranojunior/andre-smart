
<%@ Language=VBScript %>
<%

'Set Upload = Server.CreateObject("Persits.Upload.1")

'Count = Upload.Save("E:\users\vdlap.com.br\website\foto\temp\")
%>

<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->

<%
cd_user = request("cd_user")
strjan = request("jan")

strcod  = request("cod")
stracao = request("acao")
strcd_funcionario = request("cd_funcionario")
'*************************************************
'*** VALORES   ***
'*************************************************
cd_salario = request("cd_salario")
nr_salario = request("nr_salario")
	if nr_salario <> "" Then
		nr_salario = formatnumber(nr_salario,2)
	end if

cd_aj_custo = request("cd_aj_custo")
nr_aj_custo = request("nr_aj_custo")

cd_insalubridade_forca = request("cd_insalubridade_forca")
insalubridade_forca = request("insalubridade_forca")
	if insalubridade_forca <> 1 then	insalubridade_forca = 0

cd_qtd_plantoes = request("cd_qtd_plantoes")
qtd_plantoes = request("qtd_plantoes")

cd_ad_noturno = request("cd_ad_noturno")
ad_noturno = request("ad_noturno")

cd_hextra = request("cd_hextra")
qtd_he = request("qtd_he")

list_depend = request("list_depend")
'cd_dependente = request("cd_dependente")

cd_vrefeic_cancela = request("cd_vrefeic_cancela")
vrefeic_cancela = request("vrefeic_cancela")
	if vrefeic_cancela <> 1 then	vrefeic_cancela = 0

cd_vtransp_cancela = request("cd_vtransp_cancela")
vtransp_cancela = request("vtransp_cancela")
	if vtransp_cancela <> 1 then	vtransp_cancela = 0

cd_conv_medico = request("cd_conv_medico")
conv_medico = request("conv_medico")

cd_contr_sind = request("cd_contr_sind")
contr_sind = request("contr_sind")

cd_desc_diversos = request("cd_desc_diversos")
desc_diversos = request("desc_diversos")

cd_faltas_justif = request("cd_faltas_justif")
nr_faltas_justif = request("nr_faltas_justif")

cd_faltas_injust = request("cd_faltas_injust")
nr_faltas_injust = request("nr_faltas_injust")

cd_dsr_faltas = request("cd_dsr_faltas")
nr_dsr_faltas = request("nr_dsr_faltas")
	if nr_faltas <> "0" AND nr_dsr_faltas = "0" then	nr_dsr_faltas = 1

cd_atrasos =  request("cd_atrasos")
nr_atrasos =  request("nr_atrasos")

cd_credito_he = request("cd_credito_he")
nr_credito_he = request("nr_credito_he")

cd_desconto_he = request("cd_desconto_he")
nr_desconto_he = request("nr_desconto_he")

cd_vr_extra = request("cd_vr_extra")
nr_vr_extra = request("nr_vr_extra")

cd_vt_extra = request("cd_vt_extra")
nr_vt_extra = request("nr_vt_extra")

dt_ano = request("dt_ano")
dt_mes = request("dt_mes")
dt_dia = day(now)
dia_final = ultimodiames(dt_mes,dt_ano)

	if dia_final < dt_dia then
		dt_dia = dia_final
	end if

'dt_ano = YEAR(Now)
'dt_mes = Month(Now)
'dt_dia = Day(now)
'dt_hora = hour(now)
'dt_minuto = minute(now)%>

<%
'**********************************************************
'*** 					Insere 						    ***
'**********************************************************
if cd_user = "46" then
	x = ""
end if

if x = "" then
	'******** Valores *******************************************************
	'***	S A L A R I O	***
	if nr_salario <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=1) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_salario = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE Salário ***
				if cod_salario = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=1, @nr_valor='"&nr_salario&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"',@cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA Salário ***
				else
					xsql = "up_funcionario_valor_update @cd_valor="&cd_salario&", @nr_valor='"&nr_salario&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

	'**********************************
	'***	A D.  N O T U R N O		***
	'**********************************
	if ad_noturno <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=8) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_ad_noturno = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_ad_noturno = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=8, @nr_valor='"&ad_noturno&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"',@cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_ad_noturno&"', @nr_valor='"&ad_noturno&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

	'***	A J.  C U S T O	***
	if nr_aj_custo <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=5) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_aj_custo = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE Ajuda de custo ***
				if cod_aj_custo = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=5, @nr_valor='"&nr_aj_custo&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"',@cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA Ajuda de custo ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_aj_custo&"', @nr_valor='"&nr_aj_custo&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

	'***	F O R Ç A	I N S A L U B R I D A D E	***
	if insalubridade_forca <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=7) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_insalubridade_forca = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_insalubridade_forca = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=7, @nr_valor='"&insalubridade_forca&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_insalubridade_forca&"', @nr_valor='"&insalubridade_forca&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

	'**************************************
	'***	Q T D	  P L A N T Õ E S	***
	'**************************************
	if qtd_plantoes <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=21) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_plantoes = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_plantoes = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=21, @nr_valor='"&qtd_plantoes&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"',@cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_qtd_plantoes&"', @nr_valor='"&qtd_plantoes&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

	'**********************************
	'***	H O R A		E X T R A	***
	'**********************************
	if qtd_he <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=11) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_he = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_he = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=11, @nr_valor='"&qtd_he&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"',@cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_hextra&"', @nr_valor='"&qtd_he&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					'xsql = "up_funcionario_valor_update @cd_valor='"&cd_hextra&"', @nr_valor=0, @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

	'***	A U X.	 C R E C H E	***
	dependentes_array = split(list_depend,",")
		For Each cod_depend In dependentes_array
			cd_auxcreche = request("cd_auxcreche"&trim(cod_depend))
			auxcreche_institui = request("auxcreche_institui"&trim(cod_depend))
				if auxcreche_institui <> 1 then auxcreche_institui = "0"
			'response.write("///"&auxcreche_institui&"\\\")
		
				'*** Verifica se já há o valor para o mês desejado ***
				xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=6) AND (cd_indice="&cod_depend&") AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
				set rs = dbconn.execute(xsql)
				while not rs.EOF
					cod_auxcreche_institui = rs("cd_codigo")
					rs.movenext
				wend
					'*** INSERE ***
						if cod_auxcreche_institui = "" then
							xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=6, @nr_valor='"&auxcreche_institui&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = "&trim(cod_depend)&",  @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
							dbconn.execute(xsql)
							response.write("x ")
					'*** ALTERA ***
						else
							xsql = "up_funcionario_valor_update @cd_valor='"&cd_auxcreche&"', @nr_valor='"&auxcreche_institui&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
							dbconn.execute(xsql)
							response.write(cod_depend&":"&auxcreche_institui&";")
						end if
					'**********************
				
		next

	'***	V A L E		R E F E I Ç Ã O	***
	if vrefeic_cancela <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=18) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_vrefeic_cancela = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_vrefeic_cancela = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=18, @nr_valor='"&vrefeic_cancela&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_vrefeic_cancela&"', @nr_valor='"&vrefeic_cancela&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

	'***	V A L E		T R A N S P O R T E	***
	if vtransp_cancela <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=4) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_vtransp_cancela = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_vtransp_cancela = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=4, @nr_valor='"&vtransp_cancela&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL,  @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_vtransp_cancela&"', @nr_valor='"&vtransp_cancela&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

	'***	C O N V Ê N I O		M É D I C O ***
	if conv_medico <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario="&strcd_funcionario&" AND cd_tipo=9 AND dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_conv_medico = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_conv_medico = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=9, @nr_valor='"&conv_medico&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_conv_medico&"', @nr_valor='"&conv_medico&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if


	'***	C O N T R I B U I Ç Ã O 	S I N D I C A L ***
	if contr_sind <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=13) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_contr_sind = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE Salário ***
				if cod_contr_sind = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=13, @nr_valor='"&contr_sind&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA Salário ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_contr_sind&"', @nr_valor='"&contr_sind&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

	'***	D E S C O N T O S	D I V E R S O S ***
	if desc_diversos <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=14) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_desc_diversos = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE Salário ***
				if cod_desc_diversos = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=14, @nr_valor='"&desc_diversos&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA Salário ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_desc_diversos&"', @nr_valor='"&desc_diversos&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if
	
	'**************************
	'***	F A L T A S 	***
	'**************************
	if nr_faltas_justif <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=22) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_faltas_justif = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_faltas_justif = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=22, @nr_valor='"&nr_faltas_justif&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_faltas_justif&"', @nr_valor='"&nr_faltas_justif&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if
	if nr_faltas_injust <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=15) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_faltas_injust = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_faltas_injust = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=15, @nr_valor='"&nr_faltas_injust&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_faltas_injust&"', @nr_valor='"&nr_faltas_injust&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if
	
	'**************************************************************
	'***	Q T D.	D E S C O N T O 	D S R	F A L T A S 	***
	'**************************************************************
	if nr_dsr_faltas <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=19) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_dsr_faltas = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_dsr_faltas = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=19, @nr_valor='"&nr_dsr_faltas&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_dsr_faltas&"', @nr_valor='"&nr_dsr_faltas&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if
	
	'**************************************************************
	'***	D E S C O N T O 		H O R A S	 A T R A S O S 	***
	'**************************************************************
	if nr_atrasos <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=20) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_atrasos = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_atrasos = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=20, @nr_valor='"&nr_atrasos&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cod_atrasos&"', @nr_valor='"&nr_atrasos&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if
	
	'***	C R E D I T O 	H O R A 	E X T R A ***
	if nr_credito_he <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=16) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_credito_he = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE Salário ***
				if cod_credito_he = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=16, @nr_valor='"&nr_credito_he&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = NULL, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA Salário ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_credito_he&"', @nr_valor='"&nr_credito_he&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if
	
	'***	D E S C O N T O 	H O R A 	E X T R A ***
	if nr_desconto_he <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=17) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_desc_he = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_desc_he = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=17, @nr_valor='"&nr_desconto_he&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = null, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_desconto_he&"', @nr_valor='"&nr_desconto_he&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if
	
	'***	V R 	E X T R A ***
	if nr_vr_extra <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=23) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_vr_extra = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_vr_extra = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=23, @nr_valor='"&nr_vr_extra&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = null, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_vr_extra&"', @nr_valor='"&nr_vr_extra&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if
	
	'***	V T 	E X T R A ***
	if nr_vt_extra <> "" Then
		'*** Verifica se já há o valor para o mês desejado ***
		xsql = "SELECT * FROM TBL_Funcionario_valores WHERE (cd_funcionario="&strcd_funcionario&") AND (cd_tipo=24) AND (dt_atualizacao BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"')"
		set rs = dbconn.execute(xsql)
		while not rs.EOF
			cod_vt_extra = rs("cd_codigo")
			rs.movenext
		wend
			'*** INSERE ***
				if cod_vt_extra = "" then
					xsql = "up_funcionario_valor_insert @cd_funcionario='"&strcd_funcionario&"', @cd_valor_tipo=24, @nr_valor='"&nr_vt_extra&"', @dt_valor_atualizacao='"&dt_mes&"/"&dt_dia&"/"&dt_ano&"', @cd_indice = null, @nm_valor_obs='"&nm_valor_obs&"', @cd_sessao='"&cd_user&"', @dt_cadastro='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
			'*** ALTERA ***
				else
					xsql = "up_funcionario_valor_update @cd_valor='"&cd_vt_extra&"', @nr_valor='"&nr_vt_extra&"', @nm_valor_obs='"&str_valor_obs&"', @cd_sessao='"&session_cd_usuario&"',@dt_atualizador='"&month(now)&"/"&day(now)&"/"&year(now)&"'"
					dbconn.execute(xsql)
				end if
			'**********************
	end if

Else
	response.write("Pula manipulação do banco")
End if
%>
   

	<%'response.redirect "../funcionario/funcionarios_folhapagamento_demonstrativo.asp?tipo=cadastro&cod="&strcd_funcionario&"&dt_dia="&dt_dia&"&dt_mes="&dt_mes&"&dt_ano="&dt_ano&""
	'jan = 10
	if jan = 1 then%>   
		<body onload="window.close();">
	<%else
		'response.redirect "../../empresa.asp?tipo=cadastro&pag=&cod="&strcod&""
		'response.redirect "../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod="&strcod&""%>
		<body onload="window.close();">
	<%end if%>
	
	<br><br>
<%'end if%>
cd_user = <%=cd_user%><br>
cd_funcionario = <%=strcd_funcionario%><br>
strjan = <%=strjan%><br><br>

cd_aj_custo = <%=cd_aj_custo%><br>
nr_aj_custo = <%=nr_aj_custo%><br><br>


cd_salario = <%=cd_salario%><br>
nr_salario = <%=nr_salario%><br><br>


cd_ad_noturno = <%=cd_ad_noturno%><br>
ad_noturno = <%=ad_noturno%><br><br>

cd_hextra = <%=cd_hextra%><br>
qtd_he = <%=qtd_he%><br><br>

cd_hextra = <%=cd_hextra%><br>
qtd_he = <%=qtd_he%><br><br>

cd_hextra = <%=cd_hextra%><br>
qtd_he = <%=qtd_he%><br><br>

cd_qtd_plantoes = <%=cd_qtd_plantoes%><br>
qtd_plantoes = <%=qtd_plantoes%><br><br>



dt_valor_atualizacao = <%=dt_mes&"/"&dt_dia&"/"&dt_ano%><br>
<br>
vrefeic_cancela = <%=vrefeic_cancela%><br>
cd_vrefeic_cancela = <%=cd_vrefeic_cancela%><br><br>
*****************************************************
list_depend = <%=list_depend%><br>
cd_dependente = <%=cd_dependente%><br>




</body>