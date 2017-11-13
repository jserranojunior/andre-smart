<!--#include file="../../../includes/inc_open_connection.asp"-->
<%' Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%dependentes_total = request("dependentes_total")%>

<table border="1" style="border-collapse:collapse;">
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;">&nbsp;</td>
		<td bgcolor="#d8d8d8" style="width:291px;">&nbsp;Nome</td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Parentesco</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Nascimento</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Idade</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Obs.</td>
	</tr>
	
	<%
		conta_linha = 1
		dependentes_array = split(dependentes_total,"$")
		For Each dependentes_item_junto In dependentes_array
		
		if dependentes_item_junto <> "" Then
			dependentes_item_len = len(dependentes_item_junto) 'tamanho da sentença toda
			
			separador1 = int(instr(dependentes_item_junto,"#")) 'diz onde está o sinal, para separação de cada item
			separador2 = int(instr(mid(dependentes_item_junto,(separador1)+1,dependentes_item_len),"#")) 
			separador3 = int(instr(mid(dependentes_item_junto,(separador1+separador2)+1,dependentes_item_len),"#")) 
			separador4 = int(instr(mid(dependentes_item_junto,(separador1+separador2+separador3)+1,dependentes_item_len),"#"))
			'separador5 = int(instr(mid(dependentes_item_junto,(separador1+separador2+separador3+separador4)+1,dependentes_item_len),"#"))
			'separador6 = int(instr(mid(dependentes_item_junto,(separador1+separador2+separador3+separador4+separador5)+1,dependentes_item_len),"#"))
			
			
			dependentes_nome = mid(dependentes_item_junto,1,separador1-1) 'Mostra o primeiro compo
			dependentes_parentesco = mid(dependentes_item_junto,separador1+1,separador2-1)'Mostra o segundo campo
			dependentes_nascimento = mid(dependentes_item_junto,separador1+separador2+1,separador3-1)'Mostra o terceiro campo
			dependentes_obs = mid(dependentes_item_junto,separador1+separador2+separador3+1, separador4-1)
			'dependentes_numero = mid(dependentes_item_junto,separador1+separador2+separador3+separador4+1, separador5-1)
			'dependentes_obs = mid(dependentes_item_junto,separador1+separador2+separador3+separador4+separador5+1, separador6-1)
			%>		
		<!--tr><td colspan="100%"><%=separador1%>.<%=separador2%>.<%=separador3%>.<%=separador4%>.<%=separador5%> (Somatudo:<%=separador1+separador2%>)</td></tr-->
		<tr>
			<td bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
			<td>&nbsp;<%=dependentes_nome%></td>
			<td>&nbsp;
			<%strsql = "Select * From TBL_dependente_parentesco where cd_codigo='"&dependentes_parentesco&"'"'Where cd_codigo=1 or cd_codigo = 2"
				Set rs_parentesco = dbconn.execute(strsql)%>
				<%Do While Not rs_parentesco.eof
					nm_parentesco = rs_parentesco("nm_parentesco")
					cd_dependente_parentesco = rs_parentesco("cd_codigo")%>	
						<%=nm_parentesco%>
					<%rs_parentesco.movenext
				loop
						rs_parentesco.close
						Set rs_parentesco = nothing%>
			
			
			<%'=dependentes_parentesco%></td>
			<td>&nbsp;<%=dependentes_nascimento%>&nbsp;</td>
			
			<%idade = datediff("m",day(dependentes_nascimento)&"/"&month(dependentes_nascimento)&"/"&year(dependentes_nascimento),now)
				if idade >= 12 then
					if idade = 12 then
						tempo = "ano"
					else
						tempo = "anos"
					end if
					idade = int(idade/12)&" "&tempo
				else 
					if idade <= 1 then
						tempo = "mes"
					else
						tempo = "meses"
					end if
					idade = idade&" "&tempo
				end if%>
			<td align="center">&nbsp;<%=idade%></td>
			<td>&nbsp;<%=dependentes_obs%></td>
			<td bgcolor="#d8d8d8"><input type="button" name="adiciona_dependente" value="-<%=dependentes_item_junto%>" onclick="adiciona_dependente('','<%'=dependentes_item_junto%>','','','','','','<%=dependentes_total%>')"></td>
			
		</tr>
		
		<%
		total_dependentes = conta_linha
		conta_linha = conta_linha + 1
		end if%>	
	<%next%>
		<tr><td colspan="2">Total de dependentes: &nbsp;<b><%=total_dependentes%></b></td></tr>
	
</table>
