<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Top 10</title>
</head>
<%ano_i = request("ano_i")
mes_i = request("mes_i")
dia_i = request("dia_i")
	


ano_f = request("ano_f")
mes_f = request("mes_f")
dia_f = ultimodiames(mes_f,ano_f)

cd_unidade = request("cd_unidade")
	if cd_unidade <> "" Then
		str_unidade = "AND a053_codung = "&cd_unidade&""
	end if

top_busca = request("top_busca")
	if top_busca = "" then
		top_busca = 10
	end if
cd_emprestimos = request("cd_emprestimos")
	if cd_emprestimos <> 1 then
		contar_emprestimos = " AND A002_tipage <> 'E' "
	end if
cd_categoria = request("cd_categoria")
	if cd_categoria = 1 then
		campos = ",a055_crmmed, a055_desmed"
		txt_cabecalho = "MÉDICOS CIRURGIÕES"
		txt_coluna = "MÉDICO/CIRURGIÃO"
	elseif cd_categoria = 2 then
		campos = ",a054_codcon, nm_convenio"
		txt_cabecalho = "CONVÊNIOS MÉDICOS"
		txt_coluna = "CONVÊNIO MÉDICO"
	elseif cd_categoria = 3 then
		campos = ",a056_codrac, a056_desrac"
		txt_cabecalho = "RACKS"
		txt_coluna = "RACK"
	elseif cd_categoria = 4 then
		campos = ",a057_codesp, nm_especialidade"
		txt_cabecalho = "ESPECIALIDADES"
		txt_coluna = "ESPECIALIDADE"
	'elseif cd_categoria = 5 then
	'	campos = ",a056_codrac, a056_desrac"
	'	txt_cabecalho = "RACKS"
	'	txt_coluna = "RACK"
	end if
%>

<body><br id="no_print">
<table align="center" id="no_print">	
	<tr>
		<td align="center">RANKING</td>
	</tr>
	<form action="../protocolo.asp" name="top10" id="top10">
	<input type="hidden" name="tipo" value="top10">
	<tr>
		<td>Período</td>
		<td><input type="text" name="dia_i" value="<%=dia_i%>" size="2" maxlength="2">/<input type="text" name="mes_i" value="<%=mes_i%>" size="2" maxlength="2">/<input type="text" name="ano_i" value="<%=ano_i%>" size="4" maxlength="4"> &nbsp; à &nbsp; <input type="text" name="dia_f" value="<%=dia_f%>" size="2" maxlength="2">/<input type="text" name="mes_f" value="<%=mes_f%>" size="2" maxlength="2">/<input type="text" name="ano_f" value="<%=ano_f%>" size="4" maxlength="4"></td>
	</tr>
	<tr>
		<td>Unidade</td>
		<td>
			<%strsql ="TBL_unidades where cd_status = 1"
			Set rs_uni = dbconn.execute(strsql)%>
			<select name="cd_unidade" class="inputs"> 
									<option value="">Selecione</option>
									<%Do While Not rs_uni.eof%>
									<option value="<%=rs_uni("cd_codigo")%>" <%if int(cd_unidade) = rs_uni("cd_codigo") then%>SELECTED<%end if%> ><%=rs_uni("nm_Unidade")%></option>
									<%rs_uni.movenext
									loop
									'rs_uni.close
									'Set rs_uni = nothing%>
									</select> 
		</td>
	</tr>
	<tr>
		<td>Tipo</td>
		<td><select name="cd_categoria">
				<option value="">Selecione</option>
				<option value="1" <%if cd_categoria = 1 then%>Selected<%end if%>>Cirurgiões</option>
				<option value="2" <%if cd_categoria = 2 then%>Selected<%end if%>>Convênios</option>
				<option value="3" <%if cd_categoria = 3 then%>Selected<%end if%>>Racks</option>
				<option value="4" <%if cd_categoria = 4 then%>Selected<%end if%>>Especialidades</option>
			</select>
		</td>
	</tr>
	<tr>
		<td>TOP</td>
		<td><input type="text" name="top_busca" value="<%=top_busca%>" size="4" maxlength="4">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Com emprestimos <input type="checkbox" name="cd_emprestimos" value="1" <%if cd_emprestimos = 1 then%>Checked<%end if%>></td>
	</tr>
	
	<tr><td colspan="100%"><img src="../../imagens/blacdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	
	<tr>
		<td colspan="100%" valign="middle" align="center"><input type="submit" name="mostra" value="Mostra"> &nbsp; &nbsp; &nbsp; &nbsp;
		<img src="../imagens/ic_print.gif" alt="Imprimir" width="24" height="24" border="0" onClick="window.print()"> &nbsp; &nbsp; &nbsp; &nbsp; <img src="../imagens/ic_print_view.gif" alt="Imprimir" width="24" height="26" border="0" onclick="visualizarImpressao();"></td>
	</tr>
	</form>
</table>
<%if dia_i <> "" AND mes_i <> "" AND ano_i <> "" Then%>
<br id="no_print">
<table align="center">	
	<%cor_linha = "#FFFFFF"
	cor = 1
	
	'*** Conta prévia dos registros para obtenção da porcentagem ***
	xsql = "SELECT TOP "&top_busca&" COUNT (a002_numpro) AS conta_previa FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"' "&str_unidade&" "&contar_emprestimos&" ORDER BY conta_previa DESC"
	Set rs = dbconn.execute(xsql)
	while not rs.EOF
		conta_previa = rs("conta_previa")
	rs.movenext	
	wend
	
	num = 1
	qtd_linhas_pag = 45
	n_pagina = 1
	
	'*** Conta os registros ***
	xsql = "SELECT top "&top_busca&" COUNT (a002_numpro) AS conta, a053_codung, nm_sigla "&campos&" FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"' "&str_unidade&" "&contar_emprestimos&" Group by a053_codung, nm_sigla "&campos&" ORDER BY conta DESC"
	Set rs = dbconn.execute(xsql)
	
	while not rs.EOF
	
		conta = rs("conta")
		'conta = 25.1
			porcentagem_total = (conta/conta_previa)*100
				porcent_virgula = instr(porcentagem_total,",")
					if porcent_virgula > 0 then
						porcentagem_parcial = replace(mid(porcentagem_total,1,porcent_virgula),",","")&mid(porcentagem_total,porcent_virgula,3)
					else
						porcentagem_parcial = porcentagem_total
					end if
		if cd_categoria = 1 then
			variavel = rs("a055_desmed")
		elseif cd_categoria = 2 then
			variavel = rs("nm_convenio")
		elseif cd_categoria = 3 then
			variavel = rs("a056_desrac")
		elseif cd_categoria = 4 then
			variavel = rs("nm_especialidade")
		end if
		
			nm_sigla = rs("nm_sigla")
			
			xsql = "SELECT * FROM TBL_Unidades WHERE nm_sigla = '"&nm_sigla&"'"
			Set rs_und = dbconn.execute(xsql)
				while not rs_und.EOF
					unidade = rs_und("nm_unidade")
				rs_und.movenext
				wend
				
				if cd_unidade = "" Then
					unidade = "Todas unidades"
				end if
	%>
	<%if cabeca = 0 Then%>
	<tr>
		<td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td align="right" colspan="100%"><br>&nbsp;<img src="../imagens/logotipo_vdlap.gif" alt="" width="91" height="51" border="0"><br>&nbsp;</td>
	</tr>
	<tr><td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td colspan="100%" align="center" style="background-color:gray; color:white; font-size:14px;">&nbsp;<br>TOP <%=top_busca%> - <%=txt_cabecalho%><br>CIRURGIAS DE <%=ucase(left(mes_selecionado(int(mes_i)),3))&"/"&ano_i%> À <%=ucase(left(mes_selecionado(int(mes_f)),3))&"/"&ano_f%><br><%=ucase(unidade)%><br>&nbsp;</td></tr>
	<tr>
		<td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td align="center" style="background-color:silver; font-size:12px;">ORDEM</td>
		<td align="left" style="background-color:silver; font-size:12px;"><%=txt_coluna%></td>
		<td align="center" style="background-color:silver; font-size:12px;">QUANTIDADE</td>
		<td align="center" style="background-color:silver; font-size:12px;">% TOTAL</td>
	</tr>
	<%cabeca = cabeca + 1
	num_linha = 0
	end if%>
	<tr>
		<td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td align="center" style="font-size:12px; background-color:<%=cor_linha%>"><%=num%></td>
		<td align="left" style="font-size:12px; background-color:<%=cor_linha%>"><%=variavel%><%'=medico%></td>
		<td align="right" style="font-size:12px; background-color:<%=cor_linha%>"><%=conta%> &nbsp; </td>
		<td align="right" style="font-size:12px; background-color:<%=cor_linha%>"><%=porcentagem_parcial %>% &nbsp; </td>
	</tr>
	<%if cor > 0 then
		cor_linha = "#d7d7d7"
		cor = 0
	else
		cor_linha = "#FFFFFF"
		cor = 1
	end if
	
	num = num + 1
	num_linha = num_linha + 1
		if num_linha = qtd_linhas_pag then
			cabeca = 0%>
			<tr><td colspan="100%"><img src="../../imagens/px.gif" alt="" width="100%" height="1" border="0"></td></tr>
			<tr><td colspan="100%" align="right" style="page-break-after:always;">Pág.: <%=n_pagina%></td></tr>
		<%n_pagina = n_pagina + 1
		end if
	conta_geral = conta_geral + conta	
	porcentagem_geral = abs(porcentagem_geral) + porcentagem_parcial
	'porcentagem_geral = porcentagem_geral) + porcentagem_parcial
	rs.movenext
	wend%>
	<tr>
		<td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td colspan="100%"><img src="../../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td>
	</tr>
	<tr style="font-size:12px;">
		<td><img src="../../imagens/px.gif" alt="" width="80" height="16" border="0"></td>
		<td colspan="2" id="ok_print">TOTAL</td>
		<td align="right">&nbsp;<%=conta_geral%> &nbsp; </td>
		<td align="right">&nbsp;<%=formatnumber(porcentagem_geral,2)%>% &nbsp; </td>
	</tr>
	<tr>
		<td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td colspan="100%"><img src="../../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td>
	</tr>
	<tr style="font-size:12px;">
		<td ><img src="../../imagens/px.gif" alt="" width="80" height="14" border="0"></td>
		<td colspan="2" id="ok_print">OUTROS</td>
		<td align="right"><%=conta_previa - conta_geral%> &nbsp; </td>
		<td align="right"><%=formatnumber(100-porcentagem_geral,2)%>% &nbsp; </td>
	</tr>
	<tr><td ><img src="../../imagens/px.gif" alt="" width="80" height="16" border="0"></td>
		<td colspan="100%">&nbsp;</td></tr>
	<tr>
		<td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="16" border="0"></td>
		<td colspan="100%"><img src="../../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td>
	</tr>
	<tr style="font-size:12px;">
		<td>&nbsp;<img src="../../imagens/px.gif" alt="" width="80" height="14" border="0"></td>	
		<td align="left" colspan="2" id="ok_print">TOTAL DE CIRURGIAS NO PERÍODO</td>
		<td align="right">&nbsp;<%=conta_previa%> &nbsp; </td>
		<td align="right">&nbsp;<%'=porcentagem_geral%></td>
	</tr>
	<%num_linha = num_linha + 5
		while not num_linha = qtd_linhas_pag%>
		<tr><td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="14" border="0"></td>
			<td colspan="4" align="right"><%'=num_linha%></td></tr>
	<%num_linha = num_linha + 1
	wend%>
	<tr><td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td colspan="100%"><img src="../../imagens/px.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr><td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td colspan="100%" align="right">Pág.: <%=n_pagina%></td></tr>
	<tr>
		<td id="ok_print"><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="270" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="120" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
	</tr>
</table>
<%end if%>

</body>
</html>
