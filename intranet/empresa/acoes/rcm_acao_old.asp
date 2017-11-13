<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->

<%
cd_user = request("cd_user")

dt_atual = month(now)&"/"&day(now)&"/"&year(now)
cd_funcionario = request("cd_funcionario")

dt_ano = request("dt_ano")
dt_mes = request("dt_mes")
dt_dia = request("dt_dia")
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
				if dt_dia > ultimodiames(dt_ano,dt_mes) then
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

	
nr_ad_noturno = request("nr_ad_noturno")
nr_vr_extra = request("nr_vr_extra")
nr_vt_extra = request("nr_vt_extra")
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

nm_obs = request("nm_obs")


	xsql = "up_funcionario_rcm_insert @cd_funcionario='"&cd_funcionario&"',@dt_rcm='"&dt_rcm&"', @dt_hora_inicio="&dt_hora_inicio&", @dt_hora_fim="&dt_hora_fim&", @nr_ad_noturno='"&nr_ad_noturno&"', @nr_vr_extra='"&nr_vr_extra&"', @nr_vt_extra='"&nr_vt_extra&"', @nr_falta_justif='"&nr_faltas_justif&"', @nr_falta_injust='"&nr_faltas_injust&"', @nr_atraso='"&nr_atrasos&"', @nm_obs='"&nm_obs&"',@cd_user="&cd_user&",@data_atual='"&dt_atual&"'"
	dbconn.execute(xsql)
	'response.redirect("../funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=&cod="&strcod&"")
%>



<body onload="window.close();">







<!-- ----- Resultadods ----- -->
cd_user: <%=cd_user%><br><br>
dt_atual: <%=dt_autal%><br><br>

funcionario: <%=cd_funcionario%><br><br>


data: <%=dt_dia%>/<%=dt_mes%>/<%=dt_ano%><br>
hora i:	<%=dt_hora_i%>:<%=dt_min_i%><br>
hora f: <%=dt_hora_f%>:<%=dt_min_f%><br>
<br>
nr_vr_extra: <%=nr_vr_extra%><br>
nr_vt_extra: <%=nr_vt_extra%><br>

nr_faltas_justif: <%=nr_faltas_justif%><br>
nr_faltas_injust: <%=nr_faltas_injust%><br>

nr_atrasos: <%=nr_atrasos%><br>

</body>


