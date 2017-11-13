<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->

<%contatos_total = request("contatos_total")

if len(contatos_total) > 4 Then%>

<table border="1" style="border:1px; border-collapse:collapse;" width="700">
	<!--tr><td colspan="6"><%'=contatos_total%></td></tr-->
	<tr>
		<td bgcolor="#d8d8d8" style="width:40px;">&nbsp;</td>
		<td bgcolor="#d8d8d8" style="width:55px;">&nbsp;Tipo</td>
		<td bgcolor="#d8d8d8" style="width:150px;">&nbsp;Tel. Cel. E-mail Nextel</td>
		<td bgcolor="#d8d8d8" style="width:140px;">&nbsp;Contato</td>
		<td bgcolor="#d8d8d8" style="width:150px;">&nbsp;Grau relac.</td>
		<td bgcolor="#d8d8d8" style="width:150px;">&nbsp;Obs.</td>
		<td bgcolor="#d8d8d8" style="width:15px;"><img src="../../imagens/px.gif" width="25" height="1"></td>
	</tr>
	<%
		conta_linha = 1
		contatos_array = split(contatos_total,"$")
		For Each contatos_item_junto In contatos_array
		
		if contatos_item_junto <> "" Then
			'contatos_item_len = len(contatos_item_junto) 'tamanho da sentença toda
			
			separador1 = int(instr(contatos_item_junto,"!")) 'diz onde está o sinal, para separação de cada item
			separador2 = instr(contatos_item_junto,"{")
			separador3 = instr(contatos_item_junto,"}")
			separador4 = instr(contatos_item_junto,"[")
			separador5 = instr(contatos_item_junto,"]")
			separador6 = instr(contatos_item_junto,"_")
			
			contatos_tipo =  mid(contatos_item_junto,1,separador1) 'Mostra o primeiro compo
				contatos_tipo = replace(contatos_tipo,"!","")
			contatos_nome =  mid(contatos_item_junto,separador1,separador2-separador1)'Mostra o segundo campo
				contatos_nome = replace(contatos_nome,"!","")
			contatos_cargo = mid(contatos_item_junto,separador2,separador3-separador2)'Mostra o terceiro campo
				contatos_cargo = replace(contatos_cargo,"{","")
			contatos_ddd =   mid(contatos_item_junto,separador3,separador4-separador3)
				contatos_ddd = replace(contatos_ddd,"}","")
			contatos_numero = mid(contatos_item_junto,separador4,separador5-separador4)
				contatos_numero = replace(contatos_numero,"[","")
			contatos_obs =   mid(contatos_item_junto,separador5,separador6-separador5)
				contatos_obs = replace(contatos_obs,"]","")
			%>		
		<!--tr><td colspan="6"><%=separador1%>.<%=separador2%>.<%=separador3%>.<%=separador4%>.<%=separador5%> (Somatudo:<%=separador1+separador2%>)</td></tr-->
		<tr>
			<td bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
			<td>&nbsp;
			<%strsql = "Select * From TBL_contato_cat where cd_codigo='"&contatos_tipo&"'"'Where cd_codigo=1 or cd_codigo = 2"
				Set rs_contato_cat = dbconn.execute(strsql)%>
				<%Do While Not rs_contato_cat.eof
					nm_parentesco = rs_contato_cat("categoria_contato")
					cd_dependente_parentesco = rs_contato_cat("cd_codigo")%>	
						<%=nm_parentesco%>
					<%rs_contato_cat.movenext
				loop
						rs_contato_cat.close
						Set rs_contato_cat = nothing%>
			</td>
			<td>&nbsp;<%=contatos_ddd%><%if contatos_ddd <> "" Then%> - <%end if%> <%=contatos_numero%></td>
			<td>&nbsp;<%=contatos_nome%></td>
			<td>&nbsp;<%=contatos_cargo%></td>
			<td>&nbsp;<%=contatos_obs%></td>
			<td>
				<input type="button" name="remove_contato" value="-<%'=contatos_item_junto%>" onclick="adiciona_contato('','<%=contatos_item_junto%>','','','','','','<%=contatos_total%>')">
			</td>
			
		</tr>
		
		<%
		total_contatos = conta_linha
		conta_linha = conta_linha + 1
		end if%>	
	<%next%>
		<tr><td colspan="6">Total de contatos a incluir: &nbsp;<b><%=total_contatos%></b></td></tr>
		<tr><td colspan="6"><a href="javascript:void();" onclick="fn()">.</a> </td></tr>
			
</table>
<%end if%>