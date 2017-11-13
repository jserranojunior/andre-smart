<!--#include file="../../../../includes/inc_open_connection.asp"-->
<%' Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%escolaridade_total = request("escolaridade_total")%>

<table border="1" style="border:2 solid gray; border-collapse:collapse;">
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;">&nbsp;</td>
		<td bgcolor="#d8d8d8" style="width:100px;">&nbsp;Grau</td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Institui&ccedil;&atilde;o</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Curso</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Andamento</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;T&eacute;rmino</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Obs.</td>
	</tr>
	
	<%'if x = 1 then
		conta_linha = 1
		escolaridade_array = split(escolaridade_total,"$")
		For Each escolaridade_item_junto In escolaridade_array
		
		if escolaridade_item_junto <> "" Then
			escolaridade_item_len = len(escolaridade_item_junto) 'tamanho da sentença toda
			
			separador1 = int(instr(escolaridade_item_junto,"#")) 'diz onde está o sinal, para separação de cada item
			separador2 = int(instr(mid(escolaridade_item_junto,(separador1)+1,escolaridade_item_len),"#")) 
			separador3 = int(instr(mid(escolaridade_item_junto,(separador1+separador2)+1,escolaridade_item_len),"#")) 
			separador4 = int(instr(mid(escolaridade_item_junto,(separador1+separador2+separador3)+1,escolaridade_item_len),"#"))
			separador5 = int(instr(mid(escolaridade_item_junto,(separador1+separador2+separador3+separador4)+1,escolaridade_item_len),"#"))
			separador6 = int(instr(mid(escolaridade_item_junto,(separador1+separador2+separador3+separador4+separador5)+1,escolaridade_item_len),"#"))
			
			escolaridade_ensino_grau = 	mid(escolaridade_item_junto,1,separador1-1) 'Mostra o primeiro compo e assim por diante
			escolaridade_instituicao = 	mid(escolaridade_item_junto,separador1+1,separador2-1)
			escolaridade_curso = 		mid(escolaridade_item_junto,separador1+separador2+1,separador3-1)
			escolaridade_andamento = 	mid(escolaridade_item_junto,separador1+separador2+separador3+1, separador4-1)
			escolaridade_termino = 		mid(escolaridade_item_junto,separador1+separador2+separador3+separador4+1, separador5-1)
			escolaridade_obs = 			mid(escolaridade_item_junto,separador1+separador2+separador3+separador4+separador5+1, separador6-1)
			%>		
		<!--tr><td colspan="100%"><%=separador1%>.<%=separador2%>.<%=separador3%>.<%=separador4%>.<%=separador5%> (Somatudo:<%=separador1+separador2%>)</td></tr-->
		
		<tr>
			<td bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
			<td>&nbsp;
				<%xsql = "Select * From TBL_escolaridade_grau where cd_codigo='"&escolaridade_ensino_grau&"'"
					Set rs_grau = dbconn.execute(xsql)
					if not rs_grau.EOF then
					cd_grau = rs_grau("cd_codigo")
					nm_grau = rs_grau("nm_grau")%>
						<%=nm_grau%>
					<%'rs_grau.movenext
					end if%>
			</td>
			<td>&nbsp;<%=escolaridade_instituicao%></td>
			<td>&nbsp;<%=escolaridade_curso%>&nbsp;</td>
			<td align="center">&nbsp;<%'=escolaridade_andamento%>
				<%xsql = "Select * From TBL_escolaridade_andamento where cd_codigo='"&escolaridade_andamento&"'"
					Set rs_andamento = dbconn.execute(xsql)
					if not rs_andamento.EOF then
					cd_andamento = rs_andamento("cd_codigo")
					nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
						<%=nm_ensino_andamento%>
					<%'rs_andamento.movenext
					end if%>
										</td>
			<td>&nbsp;<%=escolaridade_termino%></td>
			<td>&nbsp;<%=escolaridade_obs%></td>			
		</tr>
		
		<%
		
		total_dependentes = conta_linha
		conta_linha = conta_linha + 1
		end if%>	
	<%next
	'end if%>
</table>
