<!--#include file="../../../../includes/inc_open_connection.asp"-->
<%' Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%valores_total = request("valores_total")%>

<table border="1" style="border:2 solid gray; border-collapse:collapse;">
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;">&nbsp;</td>
		<td bgcolor="#d8d8d8" style="width:100px;">&nbsp;Tipo</td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Valor</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Data</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Obs.</td>
	</tr>
	
	<%'if x = 1 then
		conta_linha = 1
		valores_array = split(valores_total,"$")
		For Each valores_item_junto In valores_array
		
		if valores_item_junto <> "" Then
			valores_item_len = len(valores_item_junto) 'tamanho da sentença toda
			
			separador1 = int(instr(valores_item_junto,"#")) 'diz onde está o sinal, para separação de cada item
			separador2 = int(instr(mid(valores_item_junto,(separador1)+1,valores_item_len),"#")) 
			separador3 = int(instr(mid(valores_item_junto,(separador1+separador2)+1,valores_item_len),"#")) 
			separador4 = int(instr(mid(valores_item_junto,(separador1+separador2+separador3)+1,valores_item_len),"#"))
			separador5 = int(instr(mid(valores_item_junto,(separador1+separador2+separador3+separador4)+1,valores_item_len),"#"))
			separador6 = int(instr(mid(valores_item_junto,(separador1+separador2+separador3+separador4+separador5)+1,valores_item_len),"#"))
			
			valores_tipo = 	mid(valores_item_junto,1,separador1-1) 'Mostra o primeiro compo e assim por diante
			nr_valor = 	mid(valores_item_junto,separador1+1,separador2-1)
			dt_valor_atualizacao = mid(valores_item_junto,separador1+separador2+1,separador3-1)
			nm_valor_obs = 	mid(valores_item_junto,separador1+separador2+separador3+1, separador4-1)
			%>		
		<!--tr><td colspan="100%"><%=separador1%>.<%=separador2%>.<%=separador3%>.<%=separador4%>.<%=separador5%> (Somatudo:<%=separador1+separador2%>)</td></tr-->
		
		<tr>
			<td bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
			<td>&nbsp;
				<%xsql = "Select * From TBL_funcionario_valores_tipos where cd_codigo='"&valores_tipo&"'"
					Set rs_valor = dbconn.execute(xsql)
					if not rs_valor.EOF then
					cd_tipo = rs_valor("cd_codigo")
					nm_valor = rs_valor("nm_valor")%>
						<%=nm_valor%>
					<%'rs_valor.movenext
					end if%>
			</td>
			<td>&nbsp;<%=formatnumber(nr_valor,2)%></td>
			<td>&nbsp;<%=dt_valor_atualizacao%>&nbsp;</td>
			<td >&nbsp;<%=nm_valor_obs%></td>			
		</tr>
		
	<%total_dependentes = conta_linha
		conta_linha = conta_linha + 1
		end if%>	
	<%next
	'end if%>
</table>
