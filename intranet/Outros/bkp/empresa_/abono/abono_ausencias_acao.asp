<!--#include file="../../includes/inc_open_connection.asp"-->
<!--include file="../../includes/util.asp"-->

<%
cd_funcionario = request("cd_funcionario")

dt_dia = request("dt_dia")
dt_mes = request("dt_mes")
dt_ano = request("dt_ano")
dt_hora = request("dt_hora")
dt_minuto = request("dt_minuto")

	Hora = Int(dt_hora*60)   '60 minutos
	Minutos = dt_minuto' Mod 60 '60 segundos
	dt_abono_minutos = hora + minutos











dt_dia_i = request("dt_dia_i")
dt_mes_i = request("dt_mes_i")
dt_ano_i = request("dt_ano_i")
dt_hora_i = request("dt_hora_i")
dt_minuto_i = request("dt_minuto_i")

	if dt_hora_i <> "" OR dt_minuto_i <> "" Then 
	dt_horas_i = dt_hora_i&":"&dt_minuto_i
	end if
	
	dt_abono_inicio = dt_mes_i&"/"&dt_dia_i&"/"&dt_ano_i & dt_horas_i

dt_dia_f = request("dt_dia_f")
dt_mes_f = request("dt_mes_f")
dt_ano_f = request("dt_ano_f")
dt_hora_f = request("dt_hora_f")
dt_minuto_f = request("dt_minuto_f")


'dt_abono = dt_dia&"/"&dt_mes&"/"&dt_ano&" "&dt_hora&":"&dt_minuto
dt_abono = dt_hora&":"&dt_minuto

dt_abono_fim = dt_dia_f&"/"&dt_mes_f&"/"&dt_ano_f&" "&dt_hora_f&":"&dt_minuto_f
	if dt_dia = "" Then 
	dt_abono = ""
	end if
	
	if dt_dia_i = "" Then
	dt_abono_inicio = ""
	end if
	
	if dt_dia_f = "" Then 
	dt_abono_fim = ""
	end if
	
cd_motivo = request("cd_motivo")
nm_obs = request("nm_obs")


'******************************************
'**               Inserir				 **
'******************************************

	xsql ="up_funcionario_abono_insert @cd_funcionario="&cd_funcionario&",@dt_abono='"&dt_abono_minutos&"', @dt_abono_inicio = '"&dt_abono_inicio&"', @dt_abono_fim = '"&dt_abono_fim&"',	@cd_motivo='"&cd_motivo&"',@nm_obs='"&nm_obs&"'"
	dbconn.execute(xsql)
	

%>
<Script language="JavaScript">
window.close()
</SCRIPT>
<br>
Funcionario: <%=cd_funcionario%><br>
<br>
Abono: <%=dt_abono%><br>
Abono em minutos: <%=dt_abono_minutos%><br>
Abono_inicio: <%=dt_abono_inicio%><br>
Abono_fim: <%=dt_abono_fim%><br>
<br>
Motivo: <%=cd_motivo%><br>
<br>
Observações:<br>
<%=nm_obs%>
<br>




