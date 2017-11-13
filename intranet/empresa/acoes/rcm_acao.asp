<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->

<%
cd_user = request("cd_user")
nao_fecha_janela = request("nao_fecha_janela")

cd_rcm = request("cd_rcm")

dt_atual = month(now)&"/"&day(now)&"/"&year(now)
cd_funcionario = request("cd_funcionario")
selecao_func = request("selecao_func")
cd_plantoes_1d = request("cd_plantoes_1d")

nm_obs = request("nm_obs")
acao = request("acao")



dt_ano = request("dt_ano")
dt_mes = request("dt_mes")
dt_dia = request("dt_dia")

dt_mes_retorno = dt_mes
dt_ano_retorno = dt_ano
	
	if dt_dia > 20 then
		dt_mes = dt_mes - 1
			if dt_mes < 1 then
				dt_mes = 12
				dt_ano = dt_ano - 1
			end if
	end if
	dt_rcm = dt_mes&"/"&dt_dia&"/"&dt_ano

dt_hora_i = request("dt_hora_i")
dt_min_i = request("dt_min_i")
	if dt_hora_i <> "" AND dt_min_i <> "" Then
		dt_hora_inicio = "'"&dt_mes&"/"&dt_dia&"/"&dt_ano&" "&dt_hora_i&":"&dt_min_i&"'"
	else
		dt_hora_inicio = "NULL"
	end if
	
dt_hora_f = request("dt_hora_f")
dt_min_f = request("dt_min_f")
	if dt_hora_f <> "" AND dt_min_f <> "" then
		'** ajuste da data final da hora extra **
		if dt_hora_f < dt_hora_i then
			dt_dia = dt_dia + 1
			
				if dt_dia > ultimodiames(dt_mes,dt_ano) then
					dt_dia = 1
					dt_mes = dt_mes + 1
						if dt_mes > 12 then
							dt_mes = 1
							dt_ano = dt_ano + 1
						end if
				end if
		end if
		
		dt_hora_fim = "'"&dt_mes&"/"&dt_dia&"/"&dt_ano&" "&dt_hora_f&":"&dt_min_f&"'"
	else
		dt_hora_fim = "NULL"
	end if
response.write(ultimodiames(dt_mes,dt_ano)&"/"&dt_mes&"/"&dt_ano&"<br>")
	
nr_ad_noturno = request("nr_ad_noturno")
nr_vr_extra = request("nr_vr_extra")
nr_vt_extra = request("nr_vt_extra")
nr_folga_enf = request("nr_folga_enf")
nr_faltas = request("nr_faltas")
	if nr_faltas = 1 then 
		nr_faltas_justif = 1
		dt_hora_inicio = "NULL"
		dt_hora_fim = "NULL"
	elseif nr_faltas = 2 then
		nr_faltas_injust = 1
		dt_hora_inicio = "NULL"
		dt_hora_fim = "NULL"
	end if

nr_atrasos = request("nr_atrasos")

'******* Plantoes *********

nr_plantoes_6h = request("nr_plantoes_6h")
nr_plantoes_8h = request("nr_plantoes_8h")
nr_plantoes_12h = request("nr_plantoes_12h")

dt_atualizacao = dt_mes&"/"&dt_dia&"/"&dt_ano



if(dt_mes = 1)then
	dt_mes_redirect = 12
	dt_ano_redirect = dt_ano - 1
	else 
	dt_mes_redirect = dt_mes - 1
	dt_ano_redirect = dt_ano
	end if



if acao = "rcm" then
	xsql = "up_funcionario_rcm_insert @cd_funcionario='"&cd_funcionario&"',@dt_rcm='"&dt_rcm&"', @dt_hora_inicio="&dt_hora_inicio&", @dt_hora_fim="&dt_hora_fim&", @nr_ad_noturno='"&nr_ad_noturno&"', @nr_vr_extra='"&nr_vr_extra&"', @nr_vt_extra='"&nr_vt_extra&"', @nr_falta_justif='"&nr_faltas_justif&"', @nr_falta_injust='"&nr_faltas_injust&"', @nr_atraso='"&nr_atrasos&"', @nr_folga_enf='"&nr_folga_enf&"', @nm_obs='"&nm_obs&"',@cd_user="&cd_user&",@data_atual='"&dt_atual&"'"
	dbconn.execute(xsql)
	
	if nao_fecha_janela = 1 then
		response.redirect("../funcionario/funcionarios_folhapagamento_rcm_cad.asp?tipo=cadastro&cod="&cd_funcionario&"&dt_dia=25&dt_mes="&dt_mes_retorno&"&dt_ano="&dt_ano_retorno&"&nao_fecha_janela=1&var=1&cat=enfermagem")
	end if


elseif acao = "rcm_excluir" then
	'***************************************************************
	'*** 					Apaga RCM			 			     ***
	'***************************************************************
		xsql= "up_funcionario_rcm_delete @cd_rcm="&cd_rcm&""
		dbconn.execute(xsql)
		response.redirect("../funcionario/funcionarios_folhapagamento_rcm_cad.asp?tipo=cadastro&cod="&cd_funcionario&"&dt_mes="&dt_mes_retorno&"&dt_ano="&dt_ano_retorno&"&var=1&cat=enfermagem")
		
		

elseif acao = "plantoes" then
	'*** Verifica se já existe registros para o mês selecionado ***
	strsql = "SELECT * FROM TBL_funcionario_plantoes WHERE dt_atualizacao='"&dt_atualizacao&"'"
	Set rs_plantoes = dbconn.execute(strsql)
	while not rs_plantoes.EOF
		cd_plantoes = rs_plantoes("cd_codigo")
		'nr_plantoes_6h = rs_plantoes("nr_plantoes_6h")
		'nr_plantoes_8h = rs_plantoes("nr_plantoes_6h")
		'nr_plantoes_18h = rs_plantoes("nr_plantoes_6h")
	rs_plantoes.movenext
	wend
	
	'**** Insere plantoes ****
	if cd_plantoes = "" Then 
		xsql = "up_funcionario_plantoes_insert @nr_plantoes_6h='"&nr_plantoes_6h&"', @nr_plantoes_8h='"&nr_plantoes_8h&"', @nr_plantoes_12h='"&nr_plantoes_12h&"', @dt_atualizacao='"&dt_atualizacao&"', @cd_user="&cd_user&",@data_atual='"&dt_atual&"'"
		dbconn.execute(xsql)
		'response.redirect("enfermagem.asp?tipo=qtd_plantoes")
		response.redirect("../../enfermagem.asp?tipo=qtd_plantoes&dt_mes="&dt_mes_redirect&"&dt_ano="&dt_ano_redirect&"&mens=Plantão_inserido_com_sucesso")
	else
		xsql = "up_funcionario_plantoes_update @cd_plantoes = "&cd_plantoes&", @nr_plantoes_6h='"&nr_plantoes_6h&"', @nr_plantoes_8h='"&nr_plantoes_8h&"', @nr_plantoes_12h='"&nr_plantoes_12h&"', @cd_user="&cd_user&",@data_atual='"&dt_atual&"'"
		dbconn.execute(xsql)
		response.redirect("../../enfermagem.asp?tipo=qtd_plantoes&dt_mes="&dt_mes_redirect&"&dt_ano="&dt_ano_redirect&"&mens=Alterado_com_sucesso")
	end if

elseif acao = "plantoes_1d_insert" then
	xsql = "up_funcionario_plantoes_1d_insert @nm_funcionarios='"&selecao_func&"', @dt_data='"&dt_rcm&"', @cd_user="&cd_user&",@data_atual='"&dt_atual&"'"
	dbconn.execute(xsql)
		'response.redirect("enfermagem.asp?tipo=qtd_plantoes")
		response.redirect("../../enfermagem.asp?tipo=qtd_plantoes&dt_mes="&dt_mes&"&dt_ano="&dt_ano&"&mens=")
		'response.write("plantoes +1d = OK!")
elseif acao = "plantoes_1d_update" then
	xsql = "up_funcionario_plantoes_1d_update @cd_plantoes_1d='"&cd_plantoes_1d&"', @nm_funcionarios='"&selecao_func&"', @cd_user="&cd_user&",@data_atual='"&dt_atual&"'"
	dbconn.execute(xsql)
		'response.redirect("enfermagem.asp?tipo=qtd_plantoes")
		response.redirect("../../enfermagem.asp?tipo=qtd_plantoes&dt_mes="&dt_mes&"&dt_ano="&dt_ano&"&mens=")
		'response.write("plantoes +1d = OK!")
end if
%>



<body onload="window.close();">







<!-- ----- Resultadods ----- -->
cd_user: <%=cd_user%><br><br>
dt_atual: <%=dt_autal%><br>
ação: <%=acao%><br><br>

cd_rcm = <%=cd_rcm%><br>
<br>

funcionario: <%=cd_funcionario%><br><br>

selecao_func: <%=selecao_func%><br><br>


data: <%=dt_dia%>/<%=dt_mes%>/<%=dt_ano%><br>
hora i:	<%=dt_hora_i%>:<%=dt_min_i%><br>
hora f: <%=dt_hora_f%>:<%=dt_min_f%><br>
<br>
nr_vr_extra: <%=nr_vr_extra%><br>
nr_vt_extra: <%=nr_vt_extra%><br>

nr_faltas_justif: <%=nr_faltas_justif%><br>
nr_faltas_injust: <%=nr_faltas_injust%><br>

nr_atrasos: <%=nr_atrasos%><br>
************************************<br>
nr_plantoes_6h: <%=nr_plantoes_6h%><br>
nr_plantoes_8h: <%=nr_plantoes_8h%><br>
nr_plantoes_12h: <%=nr_plantoes_12h%><br>
dt_atualizacao: <%=dt_atualizacao%><br>
cd_user: <%=cd_user%><br>
data_atual: <%=dt_atual%><br>

</body>


