<html>
<head>
	<title>Untitled</title>
</head>
<%
dt_dia = request("dia_sel")
dt_mes = request("mes_sel")
dt_ano = request("ano_sel")
data_atual = dt_dia&"/"&dt_mes&"/"&dt_ano
'data_atual = dt_mes&"/"&dt_dia&"/"&dt_ano

if dt_dia = "" OR dt_mes = "" OR  dt_ano = "" then
	data_atual = now
end if

'tipo = request("tipo")
'data_atual = "1/04/2010"
hoje = day(data_atual)
ultimo_dia = UltimoDiaMes(month(data_atual),year(data_atual))
%>


<body>
<table border="0" width="50">
	<!--tr><td colspan="8" align="center"><a href="#"><<</a> &nbsp;<b>Calendario</b> - <%=mes_selecionado(month(data_atual))&"-"&year(data_atual)%> &nbsp;>></td></tr-->
	<tr><td><a href="<%=pagina_atual%>?tipo=<%=tipo%>&dia_sel=<%=day(data_atual)%>&mes_sel=<%=month(data_atual)-1%>&ano_sel=<%=year(data_atual)%>"><<</a></td><td colspan="5"> &nbsp;<b>Calendario</b> - <%=mes_selecionado(month(data_atual))&"-"&year(data_atual)%></td><td><a href="<%=pagina_atual%>?tipo=<%=tipo%>&dia_sel=<%=day(data_atual)%>&mes_sel=<%=month(data_atual)+1%>&ano_sel=<%=year(data_atual)%>">>></a></a></td></tr>
	
	<tr>
		<td align="center" bgcolor="gray" style="color:white;"><b>S</b></td>
		<td align="center" bgcolor="gray" style="color:white;"><b>T</b></td>
		<td align="center" bgcolor="gray" style="color:white;"><b>Q</b></td>
		<td align="center" bgcolor="gray" style="color:white;"><b>Q</b></td>
		<td align="center" bgcolor="gray" style="color:white;"><b>S</b></td>
		<td align="center" bgcolor="gray" style="color:white;"><b>S</b></td>
		<td align="center" bgcolor="gray" style="color:white;"><b>D</b></td>
	</tr>	
	<tr>
		<%dia = 1
		numero = 1
		semana = 7
		ajusta_dia = weekday(dia&"/"&month(data_atual)&"/"&year(data_atual))-1
		
		while not numero = ajusta_dia%>
		<td></td>
		<%numero = numero  + 1
		wend
		'numero = 1
		while not dia = ultimo_dia+1
			if numero => 6 then
				marca_dia = "red"
				marca_dia_tam = 10
				if int(dt_dia) = int(dia) then
					color_bg = "yellow"
				else
					color_bg = "#e5e5e5"
				end if
			elseif int(dia) = int(hoje) then
				marca_dia = "yellow"
				marca_dia_tam = 10
				color_bg = "black"
			else
				marca_dia = "black"
				marca_dia_tam = 10
				cor_celula = "gray"
				color_bg = "#e5e5e5"
			end if%>
		<td bgcolor="<%=color_bg%>" align="center" style=""><a href="<%=pagina_atual%>?tipo=<%=tipo%>&dia_sel=<%=dia%>&mes_sel=<%=month(data_atual)%>&ano_sel=<%=year(data_atual)%>" style="font-size:<%=marca_dia_tam%>; color:<%=marca_dia%>; background-color:<%=color_bg%>;">&nbsp;<%=dia%>&nbsp;</a></td>			
			<%if numero = semana then '*** Pula linha**%>
			</tr><tr>
			<%numero = 0
			end if%>
					
		<%dia = dia + 1
		numero = numero + 1
		wend%>
	</tr>
	<!--tr>
		<td colspan="100%"><%=dt_dia%>/<%=dt_mes%>/<%=dt_ano%></td>
	</tr-->
</table>
</body>
</html>