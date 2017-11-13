<!--DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"-->
<%
cd_reserva = request("cd_reserva")
cd_situacao = request("cd_situacao")
nm_paciente = request("nm_paciente")
dt_cirurgia = request("dt_cirurgia")
	dt_horario = hour(dt_cirurgia)&":"&hour(dt_cirurgia)
	dt_data = day(dt_cirurgia)&"/"&month(dt_cirurgia)&"/"&year(dt_cirurgia)
nm_medico = request("nm_medico")
nm_centro = request("nm_centro")
nm_cirurgia = request("nm_cirurgia")
'	nm_cirurgia = int(nm_cirurgia)

	'while cirurgia_inicio 
	
	
	
nm_material = request("nm_material")
nm_equipamentos = request("nm_equipamentos")
nm_obs = request("nm_obs")




%>


<b>Reserva:</b> <%=cd_reserva%> - <%=cd_situacao%><br>
<b>Paciente:</b> <%=nm_paciente%><br>
<b>Data:</b> <%=dt_data%> <br>
<b>Hora:</b> <%=dt_horario%><br>
<b>Medico:</b> <%=nm_medico%><br>
<b>Centro:</b> <%=nm_centro%><br>
<b>Cirurgia:</b> <%=nm_cirurgia%><br>
<b>Material:</b> <%=nm_material%><br>
<b>Equipamentos:</b> <%=nm_equipamentos%><br>
<b>Obs:</b> <%=nm_obs%> <br>
