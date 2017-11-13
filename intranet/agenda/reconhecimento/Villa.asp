<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
<%'agendamento = request("agendamento")
	agendamento = Replace(request("agendamento"),vbcrlf,";<br>")

'leforte






%>

</head>

<body>

<%=agendamento%>

<form action="../agenda.asp" name="agendamento" id="agendamento">
	<input type="hidden" name="tipo" value="teste">
	<textarea cols="80" rows="10" name="agendamento"></textarea>
	<input type="submit" name="agendar" value="Agendar">
</form>
<table>
	<%agenda_array = split(agendamento,";")
		linha_numero = 1
		
		For Each agenda_item In agenda_array
		if linha_numero = 1 then
		linha = "DATA"
		elseif linha_numero = 2 then 
			linha = "HORARIO"
		elseif linha_numero = 2 then
			linha = "NOME DO PACIENTE"
		elseif linha_numero = 3 then
			linha = "CIRURGIA"
		elseif linha_numero = 4 then 
			linha = "CIRURGIÃO"
		elseif linha_numero = 5 then
			linha = "MATERIAL"
		end if%>
 
		<tr>
			<td valign="top"><%=linha%></td>
			<td valign="top"><%=agenda_item%></td>
		</tr>
	<%linha_numero = linha_numero + 1
	next%>
</table>



</body>
</html>
