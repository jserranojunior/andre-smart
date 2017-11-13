


<%dt_inicio_sys = "01/01/2005"
nome_arquivo = "protocolo.asp"
'nome_arquivo = request("arquivo")

dia_sel = zero(request("dia_sel"))
mes_sel = zero(request("mes_sel"))
ano_sel = request("ano_sel")
data_atual = dia_sel&"/"&mes_sel&"/"&ano_sel

dia_prod = zero(day(now))
mes_prod = zero(month(now))
ano_prod = year(now)
data_prod = dia_prod&"/"&mes_prod&"/"&ano_prod

if year(dt_inicio_sys) = int(ano_prod) then
	ano_sel = year(dt_inicio_sys)
end if
cd_unidade = request("cd_unidade")
	if cd_unidade = "" then cd_unidade = 0 end if
mostra_emprestimos = request("mostra_emprestimos")
	if mostra_emprestimos = "" Then mostra_emprestimos = 0 end if%>




<%sel_ano = ano_sel
	mes_i = 1
	mes_f = 12
	'fim_mes = UltimoDiaMes(mes_f,sel_ano)
	mensal = 1
	numero = 1%>
<br>
<table align="center" style="border:1px solid gray; font-size: 13px; border-collapse:collapse;">
	<form action="<%=nome_arquivo%>" name="selecao" method="get">
					<input type="hidden" name="tipo" value="<%=tipo%>">
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">
					<tr>
						<td align="center">Ano</td>
						<td align="center">Unidade</td>
					</tr>
					<tr>
						<td align="center">
						<select name="ano_sel">
								<%ano_inicio = year(dt_inicio_sys)
								response.write(ano_inicio)
								do while not ano_inicio = ano_prod + 1%>
								<option value="<%=ano_inicio%>" <%if int(ano_inicio)=ano_prod then%><%response.write("selected")%><%end if%>><%=ano_inicio%></option>
								<%ano_inicio = ano_inicio + 1
								loop%>			
						</select><br>									
						</td>
						<td><%strsql_unidade = "SELECT DISTINCT TBL_Protocolo.A053_codung AS cd_codigo, TBL_unidades.nm_unidade, TBL_unidades.nm_sigla FROM TBL_Protocolo INNER JOIN TBL_unidades ON TBL_Protocolo.A053_codung = TBL_unidades.cd_codigo WHERE YEAR(TBL_Protocolo.A002_datpro)='"&ano_sel&"'"
						Set rs_unidade = dbconn.execute(strsql_unidade)%>								
						<select name="cd_unidade">
								<option value=""></option>
								<%while not rs_unidade.EOF
								cd_und = rs_unidade("cd_codigo")
								nm_unidade = rs_unidade("nm_unidade")%>
								<option value="<%=cd_und%>" <%if int(cd_unidade) = cd_und then response.write("selected") end if%>><%=nm_unidade%></option>
								<%rs_unidade.movenext
								wend%>			
						</select>
						</td>
					</tr>						
					<tr>
						<td align="center" colspan="100%">&nbsp;<br><input type="submit" value="OK"><br>&nbsp;</td>
					</tr>
					<tr><td align="center" colspan="100%"><input type="checkbox" name="mostra_emprestimos" value="1" <%if mostra_emprestimos = 1 then response.write("checked") end if%>>Marque para contar os emprestimos</td></tr>
					</form>
</table><br>
<%if ano_sel = "" AND cd_unidade = 0 Then%>
<table align="center"><tr><td>1. Selecione o ano e uma Unidade</td></tr></table>
<%elseif ano_sel <> "" AND cd_unidade = "" Then%>
<table align="center"><tr><td>2. Selecione uma Unidade</td></tr></table>
<%elseif ano_sel <> "" AND cd_unidade <> "" Then%>
<table align="center">
	
	<tr>
		<td colspan="3" align="left">&nbsp;<img src="../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></td>
		<td colspan="8" align="center">VD LAP Cirúrgica LTDA</td>
		<td colspan="5" align="right"></td>		
	</tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%if isnumeric(cd_unidade) then
		strsql_unidade = "SELECT * FROM TBL_unidades Where cd_codigo = '"&cd_unidade&"'"
			
		Set rs_unidade = dbconn.execute(strsql_unidade)%>
		<%if not rs_unidade.EOF then
		nm_unidade = rs_unidade("nm_unidade")%>	
		<%rs_unidade.movenext
		end if
	Else
		nm_unidade = "Quadro geral - Todas Unidades"
	End if%>
		
	<tr>
		<td style="border:1px; solid red;" colspan="100%" align="center" ><b>Relatório de protocolos digitados</b><br><b><%=nm_unidade%></b><br><b><%=mes_selecionado(int(mes_sel))%>/<%=sel_ano%></b></td>
	</tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr style="border:1px solid black;">
		<td align="left">&nbsp;<b>Contagem</b></td>
		<td align="right"><b>JAN</b></td>
		<td align="right"><b>FEV</b></td>
		<td align="right"><b>MAR</b></td>
		<td align="right"><b>ABR</b></td>
		<td align="right"><b>MAI</b></td>
		<td align="right"><b>JUN</b></td>
		<td align="right"><b>JUL</b></td>
		<td align="right"><b>AGO</b></td>
		<td align="right"><b>SET</b></td>
		<td align="right"><b>OUT</b></td>
		<td align="right"><b>NOV</b></td>
		<td align="right"><b>DEZ</b></td>
		<td align="right" style="border:1px; solid red;"><b> &nbsp; &nbsp; &nbsp; TOTAL</b></td>		
	</tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%linha_num =  1
	
	strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" ORDER BY nm_unidade"
	Set rs_unidade = dbconn.execute(strsql)
	While not rs_unidade.EOF
	cd_unidade = rs_unidade("cd_codigo")
	nm_unidade = rs_unidade("nm_unidade")
	
		if mostra_emprestimos <> 1 then
			emprestimos = " AND a002_tipage <> 'E' "
		end if
		
	if linha_num = 1 then 
	corlinha = "#dadada"
	else
	corlinha = "#ffffff"
	end if%>
	<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');">
		<td align="left">&nbsp;<b>Banco<%'=nm_unidade%></b></td>
		
		<%while not numero = 13
		fim_mes = UltimoDiaMes(numero,sel_ano)
		
		if isnumeric(cd_unidade) then
			'strsql_estat = "SELECT COUNT (a002_numpro) AS conta, dt_ano, dt_mes,a053_codung, nm_sigla,a057_codesp,nm_sigla, nm_especialidade FROM TBL_protocolo Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' AND a053_codung = '"&cd_unidade&"' AND a057_codesp = '"&cd_especialidae&"' "&emprestimos&" Group by dt_ano, dt_mes,a053_codung, nm_sigla, a057_codesp,nm_sigla, nm_especialidade	ORDER BY dt_ano, dt_mes  ASC"
			strsql_estat = "SELECT COUNT (a002_numpro) AS conta FROM TBL_protocolo Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' AND a053_codung = '"&cd_unidade&"' "&emprestimos&" "
		Else
		'	strsql_estat = "SELECT COUNT (a002_numpro) AS conta, dt_ano, dt_mes,a057_codesp, nm_especialidade FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' AND a057_codesp = '"&cd_especialidae&"' "&emprestimos&" Group by dt_ano, dt_mes, a057_codesp, nm_especialidade ORDER BY dt_ano, dt_mes  ASC"
			strsql_estat = "SELECT COUNT (a002_numpro) AS conta FROM TBL_protocolo Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' "&emprestimos&" "
		end if
		
		Set rs_estat = dbconn.execute(strsql_estat)
		
		if not rs_estat.EOF Then
		conta = rs_estat("conta")
		end if
			if conta = 0 Then
				contagem = "&nbsp;"
			Else
				contagem = conta
			end if%>
		
		<td align="right"><%=contagem%><%'="/"&numero%></td>		
		<%total_esp_ano_1 = total_esp_ano_1 + conta
			If numero = 1 then
				contagem_jan = contagem_jan + conta
			end if
			If numero = 2 then
				contagem_fev = contagem_fev + conta
			end if
			If numero = 3 then
				contagem_mar = contagem_mar + conta
			end if
			If numero = 4 then
				contagem_abr = contagem_abr + conta
			end if
			If numero = 5 then
				contagem_mai = contagem_mai + conta
			end if
			If numero = 6 then
				contagem_jun = contagem_jun + conta
			end if
			If numero = 7 then
				contagem_jul = contagem_jul + conta
			end if
			If numero = 8 then
				contagem_ago = contagem_ago + conta
			end if
			If numero = 9 then
				contagem_set = contagem_set + conta
			end if
			If numero = 10 then
				contagem_out = contagem_out + conta
			end if
			If numero = 11 then
				contagem_nov = contagem_nov + conta
			end if
			If numero = 12 then
				contagem_dez = contagem_dez + conta
			end if
		
		conta = 0
		numero = numero + 1
		wend%>
		<td align="right"><b><%=total_esp_ano_1%></b></td>
	</tr>
	<%total_esp_ano = 0
	if numero = 13 then
		numero = 1
	end if
	
	'total_mensal =  total_mensal + conta
	if mensal > 12 then
		mensal = 1
	end if
	
	linha_num = linha_num + 1
			if linha_num > 2 then
			linha_num = 1
			end if
	
	total_anterior = 0
	total_esp_ano = 0
	
	rs_unidade.movenext
	wend%>
	
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%linha_num =  1
	
	strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" ORDER BY nm_unidade"
	Set rs_unidade = dbconn.execute(strsql)
	While not rs_unidade.EOF
	cd_unidade = rs_unidade("cd_codigo")
	nm_unidade = rs_unidade("nm_unidade")
	
		if mostra_emprestimos <> 1 then
			emprestimos = " AND a002_tipage <> 'E' "
		end if
		
	if linha_num = 1 then 
	'corlinha = "#dadada"
	corlinha = "#ccff99"
	else
	corlinha = "#ffffff"
	end if%>
	<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#33ff66');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');">
		<td align="left">&nbsp;<b>Oficial</b></td>
		
		<%while not numero = 13
		fim_mes = UltimoDiaMes(numero,sel_ano)
		
		if isnumeric(cd_unidade) then
			'strsql_estat = "SELECT COUNT (a002_numpro) AS conta, dt_ano, dt_mes,a053_codung, nm_sigla,a057_codesp,nm_sigla, nm_especialidade FROM TBL_protocolo Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' AND a053_codung = '"&cd_unidade&"' AND a057_codesp = '"&cd_especialidae&"' "&emprestimos&" Group by dt_ano, dt_mes,a053_codung, nm_sigla, a057_codesp,nm_sigla, nm_especialidade	ORDER BY dt_ano, dt_mes  ASC"
			strsql_estat = "SELECT COUNT (a002_numpro) AS conta FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' AND a053_codung = '"&cd_unidade&"' "&emprestimos&" "
		Else
			'strsql_estat = "SELECT COUNT (a002_numpro) AS conta, dt_ano, dt_mes,a057_codesp, nm_especialidade FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' AND a057_codesp = '"&cd_especialidae&"' "&emprestimos&" Group by dt_ano, dt_mes, a057_codesp, nm_especialidade ORDER BY dt_ano, dt_mes  ASC"
			strsql_estat = "SELECT COUNT (a002_numpro) AS conta FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' "&emprestimos&" "
		end if
		
		Set rs_estat = dbconn.execute(strsql_estat)
		
		if not rs_estat.EOF Then
		conta = rs_estat("conta")
		end if
			if conta = 0 Then
				contagem = "&nbsp;"
			Else
				contagem = conta
			end if%>
		
		<td align="right"><%=contagem%><%'="/"&numero%></td>		
		<%total_esp_ano_2 = total_esp_ano_2 + conta
			If numero = 1 then
				contagem_jan_2 = contagem_jan_2 + conta
			end if
			If numero = 2 then
				contagem_fev_2 = contagem_fev_2 + conta
			end if
			If numero = 3 then
				contagem_mar_2 = contagem_mar_2 + conta
			end if
			If numero = 4 then
				contagem_abr_2 = contagem_abr_2 + conta
			end if
			If numero = 5 then
				contagem_mai_2 = contagem_mai_2 + conta
			end if
			If numero = 6 then
				contagem_jun_2 = contagem_jun_2 + conta
			end if
			If numero = 7 then
				contagem_jul_2 = contagem_jul_2 + conta
			end if
			If numero = 8 then
				contagem_ago_2 = contagem_ago_2 + conta
			end if
			If numero = 9 then
				contagem_set_2 = contagem_set_2 + conta
			end if
			If numero = 10 then
				contagem_out_2 = contagem_out_2 + conta
			end if
			If numero = 11 then
				contagem_nov_2 = contagem_nov_2 + conta
			end if
			If numero = 12 then
				contagem_dez_2 = contagem_dez_2 + conta
			end if
		
		conta = 0
		numero = numero + 1
		wend%>
		<td align="right"><b><%=total_esp_ano_2%></b></td>
	</tr>
	<%
	total_esp_ano = 0
	if numero = 13 then
		numero = 1
	end if
	
	'total_mensal =  total_mensal + conta
	if mensal > 12 then
		mensal = 1
	end if
	
	linha_num = linha_num + 1
			if linha_num > 2 then
			linha_num = 1
			end if
	
	total_anterior = 0
	total_esp_ano = 0
	
	rs_unidade.movenext
	wend%>
	
	<%if linha_num = 1 then 
	corlinha = "#dadada"
	'corlinha = "#ccff99"
	else
	corlinha = "#ffffff"
	end if%>
	
	
	<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#FFFFFF');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');">
		<td align="left">&nbsp;<b>Diferenças</b></td>
		
		<%while not numero = 13%>		
		<td align="right" style="color:red;"><%'=contagem%><%'=contagem_jan_2 - contagem_jan%>		
		<%diferenca_total = total_esp_ano_2 - total_esp_ano_1
			'link_i = "<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=2&cd_unidade="&cd_unidade&"></a>"
		
			If numero = 1 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=1&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_jan_2 - contagem_jan&"</a>")
			end if
			If numero = 2 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=2&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_fev_2 - contagem_fev&"</a>")
			end if
			If numero = 3 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=3&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_mar_2 - contagem_mar&"</a>")
			end if
			If numero = 4 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=4&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_abr_2 - contagem_abr&"</a>")
			end if
			If numero = 5 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=5&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_mai_2 - contagem_mai&"</a>")
			end if
			If numero = 6 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=6&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_jun_2 - contagem_jun&"</a>")
			end if
			If numero = 7 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=7&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_jul_2 - contagem_jul&"</a>")
			end if
			If numero = 8 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=8&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_ago_2 - contagem_ago&"</a>")
			end if
			If numero = 9 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=9&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_set_2 - contagem_set&"</a>")
			end if
			If numero = 10 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=10&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_out_2 - contagem_out&"</a>")
			end if
			If numero = 11 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=11&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_nov_2 - contagem_nov&"</a>")
			end if
			If numero = 12 then
				response.write("<a href=protocolo.asp?tipo=qtd_digitados&arquivo=protocolo.asp&ano_sel="&ano_sel&"&mes_sel=12&cd_unidade="&cd_unidade&"&mostra_emprestimos="&mostra_emprestimos&">"&contagem_dez_2 - contagem_dez&"</a>")
			end if%>
		</td>
		<%conta = 0
		numero = numero + 1
		wend%>
		<td align="right"><b><%=diferenca_total%></b></td>
	</tr>
	<tr><td colspan="100%">&nbsp;</td></tr>	
	<%if mes_sel <> "" Then
		ultimo_dia = UltimoDiaMes(mes_sel,ano_sel)%>
	<tr bgcolor="#808080" style="color:white;">
		<td>&nbsp;Protocolo</td>
		<td>&nbsp;Agenda</td>
		<td colspan="5">&nbsp;Nome</td>
		<td>&nbsp;Data</td>
	</tr>
	<%	strsql_diff = "SELECT * FROM TBL_Protocolo WHERE (NOT EXISTS (SELECT * FROM VIEW_protocolo_lista WHERE TBL_Protocolo.a002_numseq = VIEW_protocolo_lista.a002_numseq)) AND (A002_datpro BETWEEN '"&mes_sel&"/01/"&ano_sel&"' AND '"&mes_sel&"/"&ultimo_dia&"/"&ano_sel&"') AND a053_codung = "&cd_unidade&""
		Set rs_diff = dbconn.execute(strsql_diff)
		
		WHILE Not rs_diff.EOF
			cd_unidade = zero(rs_diff("a053_codung"))
			cd_protocolo = proto(rs_diff("a002_numpro"))
				num_protocolo = digitov(cd_unidade&"."&cd_protocolo)
			nm_paciente = rs_diff("a002_pacnom")
			data = rs_diff("a002_datpro")
			tipo_agenda = rs_diff("a002_tipage")
				if tipo_agenda = "E" Then
					emprestimo = "E"
				end if
				
				
				'digitov(cd_unidade&"."&proto(cd_protocolo))%>
			<tr>
				<td>&nbsp;<a href="javascript: void(0);" onclick="window.open('protocolo/protocolo_folha.asp?tipo=ver&cd_unidade=<%=cd_unidade%>&cd_protocolo=<%=cd_protocolo%>&cd_digito=<%=right(digitov(cdun(a053_codung)&proto(a002_numpro)),2)%>&janelas_cadastros', '', 'scrollbars=1, width=461, height=570' ); return false;">
				<%=num_protocolo%></a></td>
				<td align="center">&nbsp;<%=emprestimo%></td>
				<td colspan="5">&nbsp;<%=nm_paciente%></td>
				<td>&nbsp;<%=data%></td>
			</tr>
		<%emprestimo = ""
		rs_diff.movenext
		WEND%>
	<%end if%>
	<tr>
		<!--td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td-->
		<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>		
	</tr>
</table>
<%end if%>

