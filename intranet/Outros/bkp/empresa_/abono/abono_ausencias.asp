
<%
cd_funcionario = request("cd_funcionario")
dt_data = request("dt_data")
%>
<html>
<head>
	<title>Abono</title>
</head>

<body onunload="window.opener.location.reload();">
ABONOS - Funcionário: <%=cd_funcionario%>
<br>


<form action="../../intranet/empresa/abono/abono_ausencias_acao.asp" method="post" name="abonos" id="abonos">
<input type="hidden" name="cd_funcionario" value="<%=cd_funcionario%>">
<!--início do abono<br>
<input type="text" name="dt_dia_i" maxlength="2" size="2" value="<%=day(dt_data)%>">/
<input type="text" name="dt_mes_i" maxlength="2" size="2" value="<%=month(dt_data)%>">/
<input type="text" name="dt_ano_i" maxlength="4" size="4" value="<%=year(dt_data)%>"> - 
<input type="text" name="dt_hora_i" size="2" maxlength="2">:
<input type="text" name="dt_minuto_i" size="2" maxlength="2"><br>
<br>
fim do abono<br>
<input type="text" name="dt_dia_f" maxlength="2" size="2">/
<input type="text" name="dt_mes_f" maxlength="2" size="2">/
<input type="text" name="dt_ano_f" maxlength="4" size="4"> - 
<input type="text" name="dt_hora_f" size="2" maxlength="2">:
<input type="text" name="dt_minuto_f" size="2" maxlength="2"><br><br>
OU<br><br-->
<input type="hidden" name="dt_dia_i" maxlength="2" size="2" value="<%=day(dt_data)%>">
<input type="hidden" name="dt_mes_i" maxlength="2" size="2" value="<%=month(dt_data)%>">
<input type="hidden" name="dt_ano_i" maxlength="4" size="4" value="<%=year(dt_data)%>">

Data do abono<br>
<input type="text" name="dt_dia" maxlength="2" size="2" value="<%=day(dt_data)%>">/
<input type="text" name="dt_mes" maxlength="2" size="2" value="<%=month(dt_data)%>">/
<input type="text" name="dt_ano" maxlength="4" size="4" value="<%=year(dt_data)%>"> <br>
Horas abonadas<br> 
<input type="text" name="dt_hora" size="2" maxlength="2">:<input type="text" name="dt_minuto" size="2" maxlength="2"><br>
<br>
Motivo<br>
<input type="radio" name="cd_motivo" value="1">Compensação<br>
<input type="radio" name="cd_motivo" value="2">Atestado<br>
<input type="radio" name="cd_motivo" value="3">Declaração de horas<br>
<input type="radio" name="cd_motivo" value="5">Férias<br>
<br>

Observações<br>
<textarea cols= rows= name='nm_obs'></textarea>
<br><br>
<input type="submit" name="OK" value="OK">
</form>
</body>
</html>
