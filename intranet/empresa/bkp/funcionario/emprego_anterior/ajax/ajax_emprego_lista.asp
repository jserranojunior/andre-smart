<!--#include file="../../../../includes/inc_open_connection.asp"-->
<%' Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%emprego_total = request("emprego_total")%>

<table border="1" style="border:2 solid gray; border-collapse:collapse;">
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;"><img src="../../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:291px;">&nbsp;Empresa<br><img src="../../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Fun&ccedil;&acirc;o<br><img src="../../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Entrada<br><img src="../../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Sa&icute;da<br><img src="../../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Obs.<br><img src="../../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
	</tr>
	
	<%'if x = 1 then
		conta_linha = 1
		emprego_array = split(emprego_total,"$")
		For Each emprego_item_junto In emprego_array
		
		if emprego_item_junto <> "" Then
			emprego_item_len = len(emprego_item_junto) 'tamanho da sentença toda
			
			separador1 = int(instr(emprego_item_junto,"#")) 'diz onde está o sinal, para separação de cada item
			separador2 = int(instr(mid(emprego_item_junto,(separador1)+1,emprego_item_len),"#")) 
			separador3 = int(instr(mid(emprego_item_junto,(separador1+separador2)+1,emprego_item_len),"#")) 
			separador4 = int(instr(mid(emprego_item_junto,(separador1+separador2+separador3)+1,emprego_item_len),"#"))
			separador5 = int(instr(mid(emprego_item_junto,(separador1+separador2+separador3+separador4)+1,emprego_item_len),"#"))
			
			emprego_empresa = mid(emprego_item_junto,1,separador1-1) 'Mostra o primeiro compo
			emprego_funcao = mid(emprego_item_junto,separador1+1,separador2-1)'Mostra o segundo campo
			emprego_entrada = mid(emprego_item_junto,separador1+separador2+1,separador3-1)'Mostra o terceiro campo
			emprego_saida = mid(emprego_item_junto,separador1+separador2+separador3+1, separador4-1)
			emprego_obs = mid(emprego_item_junto,separador1+separador2+separador3+separador4+1, separador5-1)
			%>		
		
		<tr>
			<td bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
			<td>&nbsp;<%=emprego_empresa%></td>
			<td>&nbsp;<%=emprego_funcao%></td>
			<td>&nbsp;<%=emprego_entrada%></td>
			<td>&nbsp;<%=emprego_saida%></td>
			<td>&nbsp;<%=emprego_obs%></td>
		</tr>
		
		<%
		total_emprego = conta_linha
		conta_linha = conta_linha + 1
		end if%>	
	<%next
	'end if%>
		<tr><td colspan="6">Total de dependentes: &nbsp;<b><%=total_emprego%></b></td></tr>
	
</table>
