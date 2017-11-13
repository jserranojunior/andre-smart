<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->

<%
OS_total = request("OS_total")
valor_total_orc = request("valor_total")
	if valor_total_orc <> "" Then
		valor_total_orc = formatnumber(valor_total_orc,2)
	end if
valor_lanc_os = request("valor_lanc_os")	
%>

<%'=OS_total&":"&valor_lanc_os&"+"&valor_total_orc%>
	<table>
	
	<%
		valor_total_total_os = 0
		conta_linha = 1
		OS_array = split(OS_total,"$")
		For Each item_junto In OS_array
		
		if item_junto <> "" Then
			item_len = len(item_junto) 'tamanho da sentença toda
			
			'response.write(OS_total)
			
			separador1 = int(instr(item_junto,"!")) 'diz onde está o sinal, para separação de cada item
			separador2 = int(instr(item_junto,"_")) 'diz onde está o sinal, para separação de cada item
			separador3 = int(instr(item_junto,"@")) 'diz onde está o sinal, para separação de cada item
				'separador4 = int(instr(contatos_item_junto,"«»"))-1 'diz onde está o sinal, para separação de cada item
			
			OS_interna = mid(item_junto,1,separador1-1) 'Mostra o primeiro compo
				'OS_interna = replace(OS_interna,",",".")
			valor_cada = mid(item_junto,separador1+1,separador2-separador1-1)'Mostra o segundo campo
				'if valor_cada <> "" AND isnumeric(valor_cada)Then
				'	valor_cada =(formatnumber(valor_cada,2))
				'end if
			'contatos_cat = mid(contatos_item_junto,separador,separador3-separador2)'Mostra o terceiro campo%>		
	<%if valor_total_total_os = 0 Then%>
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;">&nbsp;nº</td>
		<td bgcolor="#d8d8d8" style="width:291px;">&nbsp;Ordem de serviço</td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Valor unitários</td>
		<td><img src="../../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
	</tr>
	<%end if%>
	<!--tr><td colspan="100%"><%=separador1%>.<%=separador2%>.<%=separador3%>.<%=separador4%>.<%=separador5%></td></tr-->
		<tr>
			<td bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
			<td><%=replace(OS_interna,",",".")%></td>
			<td align="right"><%if isnumeric(valor_cada) then response.write(formatnumber(valor_cada,2))%></td>
			<td bgcolor="#d8d8d8">
				<input type="button" name="adiciona_OS" value="-" onFocus="adiciona_OS('','<%=OS_interna%>!<%=valor_cada%>_$',document.form.valor_unitario.value,'<%=valor_lanc_os%>','<%=valor_total_orc%>',document.form.OS_total.value)">
			</td>
		</tr>
		
		
		<%valor_total_os = valor_total_os + abs(valor_cada)
		
		conta_linha = conta_linha + 1
		end if%>	
		<%next%>
		
		<%soma_valores_os = valor_total_os + valor_lanc_os
		diferenca_valor_x_osunica = formatnumber(soma_valores_os - valor_total_orc,2)%>
		<tr><td colspan="4"><img src="../../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td></tr>
		<tr>
			<td colspan="2">Total ordens serviço</td>
			<td align="right"><%if isnumeric(valor_total_os) then response.write(formatnumber(valor_total_os,2))%></td>
		</tr>
		<tr><td colspan="4" align="right"><input type="button" value="Vincular" onclick="window.open('../rel_orc_aprov')"></td></tr>
		<tr><td colspan="4">&nbsp;</td></tr>
		<tr>
			<td colspan="2">Total O.S. lançadas</td>
			<td align="right"><%if isnumeric(valor_lanc_os) then response.write(formatnumber(valor_lanc_os,2))%></td>
		</tr>
		<tr>
			<td colspan="2">Total orçamento</td>
			<td align="right"><%=valor_total_orc%></td>
		</tr>
		<tr><td colspan="4"><img src="../../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td></tr>
		<tr>
			<td colspan="2">Diferença</td>
			<td align="right" style="color:red;"><b><%=diferenca_valor_x_osunica%></b></td>
		</tr>
</table>
