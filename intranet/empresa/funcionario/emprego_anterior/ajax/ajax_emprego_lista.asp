<!--#include file="../../../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->

<%emprego_total = request("emprego_total")%>
<%if len(emprego_total) > 4 Then%>
<table border="1" width="600" style="border:2 solid gray; border-collapse:collapse;">
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:291px;">&nbsp;Empresa<br><img src="../../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Fun&ccedil;&acirc;o<br><img src="../../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Entrada<br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Sa&iacute;da<br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Obs.<br><img src="../../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
	</tr>
	
	<%'if x = 1 then
		conta_linha = 1
		emprego_array = split(emprego_total,"$")
		For Each emprego_item_junto In emprego_array
		
		if emprego_item_junto <> "" Then
			
			emprego_item_junto = replace(emprego_item_junto,"%20"," ")
				emprego_item_junto = replace(emprego_item_junto,"Ã§","ç")
				
				separador1 = int(instr(emprego_item_junto,"!"))
				separador2 = instr(emprego_item_junto,"{")
				separador3 = instr(emprego_item_junto,"}")
				separador4 = instr(emprego_item_junto,"[")
				separador5 = instr(emprego_item_junto,"]")
				'separador6 = instr(emprego_item_junto,"_")
						
				emprego_empresa = mid(emprego_item_junto,1,separador1)
					emprego_empresa = replace(emprego_empresa,"!","")
				emprego_funcao = mid(emprego_item_junto,separador1,separador2-separador1)
					emprego_funcao = replace(emprego_funcao,"!","")
				emprego_entrada = mid(emprego_item_junto,separador2,separador3-separador2)
					emprego_entrada = replace(emprego_entrada,"{","")
					if len(emprego_entrada) > 3 Then
						'emprego_entrada = "'"&month(emprego_entrada)&"/"&day(emprego_entrada)&"/"&year(emprego_entrada)&"'"
						emprego_entrada = month(emprego_entrada)&"/"&day(emprego_entrada)&"/"&year(emprego_entrada)
					else
						emprego_entrada = ""
					end if
				emprego_saida = mid(emprego_item_junto,separador3,separador4-separador3)
					emprego_saida = replace(emprego_saida,"}","")
						if len(emprego_saida) > 3 Then
							'emprego_saida = "'"&month(emprego_saida)&"/"&day(emprego_saida)&"/"&year(emprego_saida)&"'"
							emprego_saida = month(emprego_saida)&"/"&day(emprego_saida)&"/"&year(emprego_saida)
						else
							emprego_saida = ""
						end if
				emprego_obs = mid(emprego_item_junto,separador4,separador5-separador4)
					emprego_obs = replace(emprego_obs,"["," ")
			%>		
		
		<tr>
			<td bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
			<td>&nbsp;<%=emprego_empresa%></td>
			<td>&nbsp;<%=emprego_funcao%></td>
			<td>&nbsp;<%=emprego_entrada%></td>
			<td>&nbsp;<%=emprego_saida%></td>
			<td>&nbsp;<%=emprego_obs%></td>
			<td>
				<input type="button" name="remove_emprego" value="-" onclick="adiciona_emprego('','<%=emprego_item_junto%>','','','','','','','','','<%=emprego_total%>')">
			</td>
		</tr>
		
		<%
		total_emprego = conta_linha
		conta_linha = conta_linha + 1
		end if%>	
	<%next
	'end if%>
		<tr><td colspan="6">Total de dependentes: &nbsp;<b><%=total_emprego%></b></td></tr>
	
</table>
<%end if%>