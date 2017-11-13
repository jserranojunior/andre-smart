<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<%dia_sel_i = request("dia_sel_i")
mes_sel_i = request("mes_sel_i")
ano_sel_i = request("ano_sel_i")

dia_sel_f = request("dia_sel_f")
mes_sel_f = request("mes_sel_f")
ano_sel_f = request("ano_sel_f")

ordem = request("ordem")

avaliacao = request("avaliacao")

cd_unidade = request("cd_unidade")%>
<table align="center" border="1" id="no_print">
	<form action="../manutencao.asp" name="relatorio" id="relatorio">
	<input type="hidden" name="tipo" value="periodo">
	<tr>
		<td>Período</td>
		<td>de: <input type="text" name="mes_sel_i" size="2" maxlength="2" value="<%=mes_sel_i%>">/<input type="text" name="ano_sel_i" size="4" maxlength="4" value="<%=ano_sel_i%>"> &nbsp; 
		até: <input type="text" name="mes_sel_f" size="2" maxlength="2" value="<%=mes_sel_f%>">/<input type="text" name="ano_sel_f" size="4" maxlength="4" value="<%=ano_sel_f%>"></td>
	</tr>
	<tr>
		<%strsql ="SELECT * FROM TBL_unidades where cd_status <> 0 "
		Set rs_uni = dbconn.execute(strsql)
		%>
		<td>Unidade: </td>
		<td><select name="cd_unidade" class="inputs" onFocus="nextfield ='num_hospital';">
					<option value="0">Todas unidades</option>
					<%Do While Not rs_uni.eof
					if int(cd_unidade) = rs_uni("cd_codigo") then%>
						<%unidade_check = "selected"
						nm_sigla = rs_uni("nm_sigla")
						nome_unidade = rs_uni("nm_unidade")%>
					<%end if%>
					<option value="<%=rs_uni("cd_codigo")%>" <%=unidade_check%>><%=rs_uni("nm_unidade")%></option>
					<%rs_uni.movenext
					unidade_check = ""
					loop%>
					</select>
		</td>
	</tr>
	<tr>
		<td>Foco: </td>
		<td><input type="radio" name="ordem" value="1" <%if ordem = 1 OR ordem = "" then response.write("checked")end if%> style="border-color:white;"> Motivo<br>
			<input type="radio" name="ordem" value="2" <%if ordem = 2 then response.write("checked")end if%> style="border-color:white;"> Data<br>
			<input type="radio" name="ordem" value="3" <%if ordem = 3 then response.write("checked")end if%> style="border-color:white;"> Equipamento</td>
	</tr>
	<tr>
		<td>Mostrar</td>
		<td><input type="radio" name="avaliacao" value="1" style="border-color:white;" <%if avaliacao = 1 OR avaliacao = "" then response.write("checked")%>>Manutenção<br>
		<input type="radio" name="avaliacao" value="2" style="border-color:white;" <%if avaliacao = 2 then response.write("checked")%>>Compra<br>
		<input type="radio" name="avaliacao" value="3" style="border-color:white;" <%if avaliacao = 3 then response.write("checked")%>>Ambos</td>
	</tr>
	<tr><td colspan="100%" align="center"><input type="submit" value="OK"></td></tr>	
	</form>
</table>
<br id="no_print">
<%if cd_unidade <> "" Then%>

<table align="center" border="0" style="font-family:arial;">	
<%if cd_unidade = 0 then
	nome_unidade = "Todas as unidades"
end if
qtd_linhas_pg = 60
qtd_dados = 1
total_parcial = 0
versao = "1.0"

if ordem = 1 or ordem = "" then
	'*** ordem por motivo, equipamento, data e o.s. ***
	ordem = "cd_espec_ordem, nm_especialidade,  nm_motivo,nm_equipamento,dt_os,num_os"
	
elseif ordem = 2 then
	'*** ordem por data, equipamento, motivo e o.s. ***
	ordem = "cd_espec_ordem, nm_especialidade,  dt_os,nm_equipamento,nm_motivo,num_os"
	
elseif ordem = 3 then
	'*** ordem por data, equipamento, motivo e o.s. ***
	ordem = "cd_espec_ordem, nm_equipamento,  nm_especialidade,dt_os,nm_motivo,num_os"
end if

if cd_unidade = "0" Then
	unidade = " "
else
	unidade = "AND (nm_unidade = '"&nm_sigla&"')"
end if

if avaliacao = 1 then
	'*** manutencao ***
	str_avaliacao = "AND cd_avaliacao = 6 "
	str_condicoes = " WHERE sequencia = '1' AND dt_os BETWEEN '"&mes_sel_i&"/1/"&ano_sel_i&"' AND '"&mes_sel_f&"/"&UltimoDiaMes(mes_sel_f,ano_sel_f)&"/"&ano_sel_f&"' "&unidade&" "&str_avaliacao&""
elseif avaliacao = 2 then
	'***  compras
	str_avaliacao = "AND cd_avaliacao = 3 "
	str_condicoes = " WHERE sequencia = '1' AND dt_os BETWEEN '"&mes_sel_i&"/1/"&ano_sel_i&"' AND '"&mes_sel_f&"/"&UltimoDiaMes(mes_sel_f,ano_sel_f)&"/"&ano_sel_f&"' "&unidade&" "&str_avaliacao&""
elseif avaliacao = 3 then
	'*** ambos ***
	str_avaliacao1 = " AND cd_avaliacao = 3 "
	str_avaliacao2 = " AND cd_avaliacao = 6 "
	str_condicoes = " WHERE sequencia = '1' AND dt_os BETWEEN '"&mes_sel_i&"/1/"&ano_sel_i&"' AND '"&mes_sel_f&"/"&UltimoDiaMes(mes_sel_f,ano_sel_f)&"/"&ano_sel_f&"' "&unidade&" "&str_avaliacao1&" OR sequencia = '1' AND dt_os BETWEEN '"&mes_sel_i&"/1/"&ano_sel_i&"' AND '"&mes_sel_f&"/"&UltimoDiaMes(mes_sel_f,ano_sel_f)&"/"&ano_sel_f&"' "&unidade&" "&str_avaliacao2&""
end if

strsql = "SELECT COUNT(num_qtd) as conta,num_os, cd_ns, cd_patrimonio, cd_equipamento,nm_equipamento, nm_unidade, cd_fornecedor,nm_fornecedor, cd_especialidade,nm_especialidade, cd_avaliacao,nm_avaliacao, cd_marca,nm_marca, num_os_assist, nm_natureza_defeito, dt_os, nm_motivo FROM VIEW_os_lista "&str_condicoes&" GROUP BY sequencia,num_os, cd_ns, cd_patrimonio, cd_equipamento,nm_equipamento, nm_unidade, cd_fornecedor,nm_fornecedor, cd_especialidade,nm_especialidade, cd_avaliacao,nm_avaliacao, cd_marca,nm_marca, num_os_assist, nm_natureza_defeito, dt_os, cd_codigo, nm_motivo, cd_espec_ordem Order by "&ordem&"  ASC"
Set rs = dbconn.execute(strsql)

		strsql_2 = "SELECT     COUNT(num_qtd) AS conta, cd_especialidade, nm_especialidade FROM VIEW_os_lista "&str_condicoes&" GROUP BY cd_especialidade, nm_especialidade, cd_espec_ordem ORDER BY cd_espec_ordem, nm_especialidade"
		Set rs2 = dbconn.execute(strsql_2)
		while not rs2.EOF 
			array_categorias = array_categorias&","&rs2("conta")'&";("&rs2("nm_especialidade")&")"
			rs2.movenext
		wend
		array_categorias = mid(array_categorias,2,len(array_categorias))

	while not rs.EOF
	
	dt_os = zero(day(rs("dt_os")))&"/"&zero(month(rs("dt_os")))&"/"&year(rs("dt_os"))
	num_os = rs("num_os")
	
	cd_especialidade = rs("cd_especialidade")
	nm_especialidade = rs("nm_especialidade")
	nm_avaliacao = rs("nm_avaliacao")

	if linha = (qtd_linhas_pg+1)  OR linha = "" Then%>
	<tr><td colspan="100%"><p>&nbsp;</p></td></tr>
	<tr>
		<td>&nbsp;</td>
		<td colspan="100%" align="center" style=" font-size:14px;">
			Relatório de Serviço de Manutenção<br>
			Referente a <%=mes_selecionado(int(mes_sel_i))%>/<%=ano_sel_i%> até <%=mes_selecionado(int(mes_sel_f))%>/<%=ano_sel_f%>
		</td>
	</tr>
	<tr><td colspan="100%"><p>&nbsp;</p></td></tr>
	
	<tr><td>&nbsp;</td><td colspan="100%" align="center" style=" font-size:14px; font-weight:bold;"><%=nome_unidade%></td></tr>
	
	<tr><td colspan="100%"><p>&nbsp;</p></td></tr>
	<tr style="background-color:silver; color: white; font-weight:bold; text-transform:uppercase;">
		<td style="background-color:white;">&nbsp;</td>
		<td>&nbsp;</td>		
		<!--td>&nbsp;</td>
		<td>&nbsp;</td-->
		<td>&nbsp;Data</td>
		<td>&nbsp;O.S.</td>
		<td>&nbsp;Material</td>
		<!--td>N/S</td-->
		<td>&nbsp;Patrimônio</td>
		<!--td>Unidade</td-->
		<!--td>Assist. Tec.</td-->
		<!--td>Especialidade</td-->
		<!--td>Marca</td-->
		<td>&nbsp;Motivo</td>
		<td>&nbsp;Tipo</td>
	</tr>
	<%linha = 1
	end if%>
	<%'****************************
	'*** Titulo da especialidade ***
	 '*****************************
	'if cd_especialidade <> ultima_especialidade AND ultima_especialidade <> "" then
	if cd_especialidade <> ultima_especialidade then 
	qtd_dados_espec = mid(array_categorias,1,instr(array_categorias,","))%>
		<tr style="background-color:silver;">
			<td style="background-color:white;">&nbsp;</td>
			<td colspan="100%" align="center"><%=nm_especialidade%> <%'=qtd_dados_espec%></td>
		</tr>
	<%linha = linha + 1
	qtd_dados = 1
	end if%>	
	<tr>
		<td>&nbsp;</td>
		<td align="center">&nbsp;<%=zero(qtd_dados)%>  <%'=linha%></td>
		<!--td>&nbsp;<%=rs("conta")%></td>
		<td>&nbsp;<%=zero(linha)&"/"&qtd_linhas_pg%></td-->
		<td>&nbsp;<%=dt_os%></td>
		<td>&nbsp;<%=num_os%></td>
		<td>&nbsp;<%=rs("nm_equipamento")%></td>
		<!--td>&nbsp;<%=rs("cd_ns")%></td-->		
		<td>&nbsp;<%=rs("cd_patrimonio")%></td>
		<!--td>&nbsp;<%=rs("nm_unidade")%></td-->
		<!--td>&nbsp;<%=rs("nm_fornecedor")%></td-->
		<!--td>&nbsp;<%=nm_especialidade%> <%'=cd_especialidade%><%'=ultima_especialidade%></td-->
		<!--td>&nbsp;<%=rs("nm_marca")%></td-->
		<td>&nbsp;<%=rs("nm_motivo")%> <%'=array_categorias%></td>
		<td>&nbsp;<%=nm_avaliacao%></td>
	</tr>
	<%'*******************************
	'*** Subtotal de cada categoria ***
	 '********************************
	 total_parcial = total_parcial + 1
	 if int(qtd_dados)&"," = qtd_dados_espec then%>
	<tr>
		<td>&nbsp;</td>
		<td colspan="100%">&nbsp;Subtotal: <b><%=qtd_dados%><%'=total_parcial%></b></td>
	</tr>
	<%array_categorias = mid(array_categorias,instr(array_categorias,",")+1,len(array_categorias))
	linha = linha + 1
	end if
	
	 '**********************************
	'*** Rodapé do corpo do documento ***
	 '**********************************
	if int(linha) = qtd_linhas_pg then%>
	
	<tr>
		<td>&nbsp;</td>
		<td align="center" colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="700" height="1" border="0">&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td colspan="3">&nbsp;Ver.:<%=versao%></td>
		<td align="center" style="page-break-after:always;">Pág: <%=zero(pagina)+1%></td>
		<td>&nbsp;</td>
		<td align="right">&nbsp;<%=mes_selecionado(month(now))&"/"&year(now)%>&nbsp;</td>
	</tr>
	<%
	pagina = pagina+1
	end if%>
<%
'array_categorias = mid(array_categorias,instr(qtd_dados_espec,","),len(array_categorias))
'if linha = qtd_linhas_pg then
ultima_especialidade =  cd_especialidade
linhas_total = linha + 1
linha = linha + 1
qtd_dados = qtd_dados + 1

'total_dados = total_dados + total_parcial
rs.movenext
wend%>
	<tr>
		<td style="background-color:white;">&nbsp;</td>
		<td colspan="100%">Subtotal: <b><%=qtd_dados-1%></b><%'=linha%><%linha = linha + 1%></td>
	</tr>
	<tr><td>&nbsp;<%'=linha%><%linha = linha + 1%></td></tr>
	<tr>
		<td>&nbsp;<%'=linha%><%linha = linha + 1%></td>
		<td align="center" colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="700" height="1" border="0">&nbsp;</td>
	</tr>
	<%strsql_2 = "SELECT     COUNT(num_qtd) AS conta, cd_especialidade, nm_especialidade FROM VIEW_os_lista "&str_condicoes&" GROUP BY cd_especialidade, nm_especialidade, cd_espec_ordem ORDER BY cd_espec_ordem, nm_especialidade"
	Set rs2 = dbconn.execute(strsql_2)
	
		while not rs2.EOF 
			'geral_categorias = geral_categorias&","&rs2("conta")&";("&rs2("nm_especialidade")&")<br>"%>
			<tr>
				<td></td>
				<td colspan="2"><%=rs2("nm_especialidade")%></td>
				<td><b><%=rs2("conta")%></b><%'=linha%></td>
			</tr>
			<%linha = linha + 1
			rs2.movenext
		wend%>
	<tr>
		<td>&nbsp;</td>
		<td colspan="100%"><%=geral_categorias %><%linha = linha + 1%></td>
	</tr>	
	<tr>
		<td>&nbsp;</td>
		<td colspan="100%" style="background">Total de manutenção da unidade: <%=nome_unidade%> = <b><%=total_parcial%></b> <%'=linha%><%linha = linha + 1%></td>
	</tr>
	<tr>
		<td>&nbsp;<%linha = linha + 1%></td>
		<td align="center" colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="700" height="1" border="0">&nbsp;</td>
	</tr>
	<%while not linha > qtd_linhas_pg' then%>
	<tr>
		<td>&nbsp;</td>
		<td colspan="100%"><%'=geral_categorias %> <%'=linha%></td>
	</tr>
	<%linha = linha + 1
	wend%>
	<tr>
		<td>&nbsp;</td>
		<td align="center" colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="700" height="1" border="0">&nbsp;</td>
	</tr><tr>
		<td>&nbsp;</td>
		<td colspan="2">&nbsp;Ver.:<%=versao%></td>
		
		<td align="center" colspan="3">Pág: <%=zero(pagina)+1%></td>
		
		<td align="right" colspan="2">&nbsp;<%=mes_selecionado(month(now))&"/"&year(now)%>&nbsp;</td>
	</tr>
	<tr>
		<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="220" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
		<!--td><img src="../imagens/blackdot.gif" alt="" width="200" height="1" border="0"></td-->
	</tr>
</table><br>
<%end if%>
</body>
</html>
