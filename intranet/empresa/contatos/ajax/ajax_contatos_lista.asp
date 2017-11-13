
<%contatos_total = request("contatos_total")%>

<table border="1" style="border-collapse:collapse;">
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;">&nbsp;</td>
		<td bgcolor="#d8d8d8" style="width:291px;">&nbsp;Nome</td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Quem &eacute;</td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Tipo</td>
	</tr>
	
	<%
		conta_linha = 1
		contatos_array = split(contatos_total,"$")
		For Each contatos_item_junto In contatos_array
		
		if contatos_item_junto <> "" Then
			contatos_item_len = len(contatos_item_junto) 'tamanho da sentença toda
			
			separador1 = int(instr(contatos_item_junto,"#")) 'diz onde está o sinal, para separação de cada item
			separador2 = int(instr(contatos_item_junto,"£")) 'diz onde está o sinal, para separação de cada item
			separador3 = int(instr(contatos_item_junto,"§")) 'diz onde está o sinal, para separação de cada item
				'separador4 = int(instr(contatos_item_junto,"«»"))-1 'diz onde está o sinal, para separação de cada item
			
			contatos_nome = mid(contatos_item_junto,1,separador1-1) 'Mostra o primeiro compo
			'contatos_pos = mid(contatos_item_junto,separador1+1,separador2-separador1)'Mostra o segundo campo
			'contatos_cat = mid(contatos_item_junto,separador,separador3-separador2)'Mostra o terceiro campo%>		
		<tr><td colspan="100%"><%=separador1%>.<%=separador2%>.<%=separador3%>.<%=separador4%>.<%=separador5%></td></tr>
		<tr>
			<td bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
			<td>&nbsp;<%=contatos_nome%></td>
			<td>&nbsp;<%=contatos_pos%></td>
			<td>&nbsp;<%=contatos_cat%></td>
			<td bgcolor="#d8d8d8"><input type="button" name="adiciona_contato" value="-" onFocus="adiciona_contato('','<%=contatos_nome%>#<%=contatos_pos%>','','<%=contatos_total%>')"><%'=contatos_nome%><%'=contatos_pos%><%'=contatos_total%></td>
		</tr>
		
		<%
		conta_linha = conta_linha + 1
		end if%>	
	<%next%>
	
</table>
