<!--#include file="../../../../includes/inc_open_connection.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%escolaridade_total = request("escolaridade_total")%>

<%if len(escolaridade_total) > 4 then%>
<table border="1" width="700" style="border:2 solid gray; border-collapse:collapse;">
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;"><br><img src="../../imagens/px.gif" width="20" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:100px;">&nbsp;Grau<br><img src="../../imagens/px.gif" width="80" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Institui&ccedil;&atilde;o<br><img src="../../imagens/px.gif" width="140" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Curso<br><img src="../../imagens/px.gif" width="110" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Andamento<br><img src="../../imagens/px.gif" width="70" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;T&eacute;rmino<br><img src="../../imagens/px.gif" width="70" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Obs.<br><img src="../../imagens/px.gif" width="170" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:15px;"><img src="../../imagens/px.gif" width="25" height="1"></td>
	</tr>
	
	<%'if x = 1 then
		conta_linha = 1
		escolaridade_array = split(escolaridade_total,"$")
		For Each escolaridade_item_junto In escolaridade_array
		
		if escolaridade_item_junto <> "" Then
			escolaridade_item_junto = replace(escolaridade_item_junto,"%20"," ")
			escolaridade_item_junto = replace(escolaridade_item_junto,"Ã§","ç")
			
			separador1 = int(instr(escolaridade_item_junto,"!")) 'diz onde está o sinal, para separação de cada item
			separador2 = instr(escolaridade_item_junto,"{")
			separador3 = instr(escolaridade_item_junto,"}")
			separador4 = instr(escolaridade_item_junto,"[")
			separador5 = instr(escolaridade_item_junto,"]")
			separador6 = instr(escolaridade_item_junto,"_")
			
			escolaridade_ensino_grau = 	mid(escolaridade_item_junto,1,separador1) 
				escolaridade_ensino_grau = replace(escolaridade_ensino_grau,"!","")
			escolaridade_instituicao = 	mid(escolaridade_item_junto,separador1,separador2-separador1)
				escolaridade_instituicao = replace(escolaridade_instituicao,"!","")
			escolaridade_curso = 		mid(escolaridade_item_junto,separador2,separador3-separador2)
				escolaridade_curso = replace(escolaridade_curso,"{","")
			escolaridade_andamento = 	mid(escolaridade_item_junto,separador3,separador4-separador3)
				escolaridade_andamento = replace(escolaridade_andamento,"}","")
			escolaridade_termino = 		mid(escolaridade_item_junto,separador4,separador5-separador4)
				escolaridade_termino = replace(escolaridade_termino,"[","")
					if str_escolaridade_termino <> "" Then
						escolaridade_termino = "'"&month(str_escolaridade_termino)&"/"&day(str_escolaridade_termino)&"/"&year(str_escolaridade_termino)&"'"
					else
						escolaridade_termino = ""
					end if
			escolaridade_obs = 			mid(escolaridade_item_junto,separador5,separador6-separador5)
				escolaridade_obs = replace(escolaridade_obs,"]","")
			%>		
		<!--tr><td colspan="100%"><%=separador1%>.<%=separador2%>.<%=separador3%>.<%=separador4%>.<%=separador5%> (Somatudo:<%=separador1+separador2%>)</td></tr-->
		
		<tr>
			<td align="right" bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
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
			<td>
				<input type="button" name="remove_escolaridade" value="-" onclick="adiciona_escolaridade('','<%=escolaridade_item_junto%>','','','','','','','','<%=escolaridade_total%>')">
			</td>	
		</tr>
		
		<%
		
		total_dependentes = conta_linha
		conta_linha = conta_linha + 1
		end if%>	
	<%next
	'end if%>
</table>
<%end if%>