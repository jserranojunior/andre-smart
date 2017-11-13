<!--#include file="../../../../includes/inc_open_connection.asp"-->
<!--#include file="../../../../includes/util.asp"-->

<%'*** Variáveis ***
data_atual = Month(now)&"/"&Day(now)&"/"&Year(now)'&" "&Hour(now)&":"&Minute(now)&":"&Second(now)
cd_user = session("cd_codigo")

cd_funcionario = request("cd_funcionario")
cd_empresa = request("cd_empresa")
cd_unidade = request("cd_unidade")
cd_cid = request("cd_cid")
	cd_cid = replace(cd_cid,".","")
cd_motivo_falta = request("cd_motivo_falta")

dt_dia = request("dt_dia")
dt_mes = request("dt_mes")
dt_ano = request("dt_ano")

		if dt_dia <> "" OR dt_mes <> "" OR dt_ano <> "" Then
			dt_ausencia = "'"&dt_mes&"/"&dt_dia&"/"&dt_ano&"'"
		else
			dt_ausencia = "NULL"
		end if

per_dia_i = request("per_dia_i")
per_mes_i = request("per_mes_i")
per_ano_i = request("per_ano_i")
per_hor_i = request("per_hor_i")
	if per_hor_i = "" Then
		per_hor_i = "00:00"
	end if
per_min_i = request("per_min_i")
	if per_min_i = "" Then
		per_min_i = "00:00"
	end if
		if per_dia_i <> "" OR per_mes_i <> "" OR per_ano_i <> "" Then
			dt_afastamento_i = "'"&per_mes_i&"/"&per_dia_i&"/"&per_ano_i&" "&per_hor_i&":"&per_min_i&"'"
		else
			dt_afastamento_i = "NULL"
		end if

per_dia_f = request("per_dia_f")
per_mes_f = request("per_mes_f")
per_ano_f = request("per_ano_f")
per_hor_f = request("per_hor_f")
	if per_hor_f = "" Then
		per_hor_f = "00:00"
	end if
per_min_f = request("per_min_f")
	if per_min_f = "" Then
		per_min_f = "00:00"
	end if
		if per_dia_f <> "" OR per_mes_f <> "" OR per_ano_f <> "" Then
			dt_afastamento_f = "'"&per_mes_f&"/"&per_dia_f&"/"&per_ano_f&" "&per_hor_f&":"&per_min_f&"'"
		else
			dt_afastamento_f = "NULL"
		end if

nm_obs = request("obs_ausencia")

cd_user = request("cd_user")
dt_cadastro = request("dt_cadastro")

'*** Fim das variáveis ***

'*** Registra a falta ***
	xsql ="up_funcionario_ausencia_insert @cd_funcionario="&cd_funcionario&", @cd_empresa="&cd_empresa&", @cd_unidade="&cd_unidade&", @cd_motivo="&cd_motivo_falta&", @nm_cid='"&cd_cid&"', @cd_user="&cd_user&", @dt_cadastro='"&data_atual&"', @dt_ausencia="&dt_ausencia&", @dt_afastamento_i="&dt_afastamento_i&", @dt_afastamento_f="&dt_afastamento_f&", @nm_obs='"&nm_obs&"'"
	dbconn.execute(xsql)
	response.redirect "../../../../empresa.asp?tipo=ausencia&func=ativo"
'response.write(xsql)

%>
xsql: <%=xsql%><br><br>


cd_funcionario : <%=cd_funcionario%><br>
cd_empresa :<%=cd_empresa%><br>
cd_unidade :<%=cd_unidade%><br>
cd_cid :<%=cd_cid%><br>
cd_motivo :<%=cd_motivo_falta%><br>

dt_ausencia :<%=dt_ausencia%><br>
dt_afastamento_i :<%=dt_afastamento_i%><br>
dt_afastamento_f :<%=dt_afastamento_f%><br>
nm_obs :<%=nm_obs%><br>

cd_user : <%=cd_user%><br>
dt_cadastro : <%=data_atual%><br>



