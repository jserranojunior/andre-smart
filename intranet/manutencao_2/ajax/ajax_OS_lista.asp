<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<!-- meta http-equiv="Content-Type" content="text/html;charset=utf-8" -->
<%
OS_total = request("OS_total")
valor_total_orc = request("valor_total")
	if valor_total_orc <> "" Then
		valor_total_orc = formatnumber(valor_total_orc,2)
		end if
valor_lanc_os = request("valor_lanc_os")
	response.write valor_total_orc
'response.write(OS_total)
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
			
			'response.write item_junto
			
			'response.write(OS_total)
			
			separador1 = int(instr(item_junto,"!")) 'diz onde está o sinal, para separação de cada item
			separador2 = int(instr(item_junto,"_")) 'diz onde está o sinal, para separação de cada item
			separador3 = int(instr(item_junto,"@")) 'diz onde está o sinal, para separação de cada item
				'separador4 = int(instr(contatos_item_junto,"«»"))-1 'diz onde está o sinal, para separação de cada item
			
			OS_interna = mid(item_junto,1,separador1-1) 'Mostra o primeiro compo
				Num_unidade_os = left(OS_interna,2)
				Len_os_interna = len(OS_interna)
				ponto_string_os = instr(1,OS_interna,".",1)
				virgula_string_os = instr(1,OS_interna,",",1)
				
					if ponto_string_os = 0 AND virgula_string_os = 0 then 
						Num_os_interna = right(OS_interna,Len_os_interna-2)
						'OS_interna = Num_unidade_os&"."&Num_os_interna
					elseif ponto_string_os = 0 AND virgula_string_os > 0 then 
						Num_os_interna = mid(OS_interna,4,Len_os_interna-3)
					elseif ponto_string_os = 0 AND virgula_string_os > 0 then 
						Num_os_interna = right(OS_interna,Len_os_interna-2)
					else
						Num_os_interna = mid(OS_interna,4,Len_os_interna-3)
					end if
					
					
					
				'response.write(ponto_string_os&"-"&virgula_string_os&" * "&Num_unidade_os)
				'OS_interna = replace(OS_interna,",",".")
				'*** Verifica se a OS já foi digitada no sistema ***
				strsql = "SELECT TOP (1) cd_codigo FROM TBL_OrdemServico2 WHERE cd_unidade = '"&Num_unidade_os&"' AND num_os = '"&Num_os_interna&"'"
				Set rs = dbconn.execute(strsql)
				if NOT rs.EOF Then 
					os_existe = 1
					cd_andamento = rs("cd_codigo")
				else
					os_existe = 0
					cd_andamento = ""
				end if
		
		
			valor_cada = mid(item_junto,separador1+1,separador2-separador1-1)'Mostra o segundo campo
				valor_cada_item = valor_cada
				
			
				
			qtd_itens = mid(item_junto,separador2,separador3-separador2)'Mostra o terceiro campo
			response.write qtd_itens
				if qtd_itens <> "" Then qtd_itens = replace(qtd_itens,"_","")
				if qtd_itens >= 1 AND Num_os_interna <> "" Then
					valor_cada_item = valor_cada_item * qtd_itens
				end if
				%>
						
	<%if conta_linha = 1 then
	'if valor_total_total_os = 0 Then%>
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;">&nbsp;nº</td>
		<td bgcolor="#d8d8d8" style="width:291px;">&nbsp;Ordem de serviço</td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Valor unitários</td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Qtd. Itens</td>
		<td><img src="../../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
	</tr>
	<%end if%>
	<!--tr><td colspan="100%"><%=separador1%>.<%=separador2%>.<%=separador3%>.<%=separador4%>.<%=separador5%></td></tr-->
		<tr>
			<td bgcolor="#d8d8d8">&nbsp;<%=conta_linha%><%'="-"&cd_andamento%></td>
			<td><%=Num_unidade_os&"."&Num_os_interna%> <%if os_existe <> 1 then response.write(" <i style='color:red;'>(O.S. não encontrada)</i>")%></td>
			<td align="right"><%if isnumeric(valor_cada) then response.write(formatnumber(valor_cada,2))%></td>
			<td align="center"><%if isnumeric(qtd_itens) then response.write(qtd_itens)%></td>
			<td bgcolor="#d8d8d8">
				<input type="button" name="adiciona_OS" value="-" onFocus="adiciona_OS('','<%=OS_interna%>!<%=valor_cada%>_<%=qtd_itens%>@$',document.form.valor_unitario.value,'<%=valor_lanc_os%>','<%=valor_total_orc%>','','',document.form.OS_total.value)">
			</td>
		</tr>
		
		
		<%if valor_cada = "" Then valor_cada = 0
		valor_total_os = valor_total_os + abs(valor_cada_item)
		if os_existe = 1 then cod_os_vincular = cod_os_vincular&";"&cd_andamento&"!"&valor_cada&"_"&qtd_itens
		
		conta_linha = conta_linha + 1
		end if%>	
		<%next%>
		
		<%cod_os_vincular = mid(cod_os_vincular,2,len(cod_os_vincular))
		
		if isnumeric(valor_lanc_os) Then soma_valores_os = valor_total_os + valor_lanc_os
		if isnumeric(soma_valores_os) then diferenca_valor_x_osunica = formatnumber(soma_valores_os - valor_total_orc,2)%>
		<tr><td colspan="4"><img src="../../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td></tr>
		<tr>
			<td colspan="2">Total ordens serviço</td>
			<td align="right"><%if isnumeric(valor_total_os) then response.write(formatnumber(valor_total_os,2))%></td>
		</tr>
		<tr><td colspan="4" align="right"><!--input type="button" value="Vincular" onclick="window.open('orcamentos_aprovados_os.asp?cd_orcamento=')"-->
			<input type="button" value="Vincular" onclick="vincula_os('<%=cod_os_vincular%>');"><%'=cod_os_vincular%></td></tr>
		<tr><td colspan="4"><span id='divOS_vinculo'><!-- ***abre o arquivo de acoes da manutenção: acoes/manutencao_acao.asp*** --><span></td></tr>
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
