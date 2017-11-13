


<%
dt_inicio_sys = "01/01/2005"
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
mostra_emprestimos = request("mostra_emprestimos")

'*********************************************************
'*				      1ª Parte	 					  	'*
'*********************************************************
'* 1.1 - Seleção do ANO *
'************************%>
<table border="1" align="center">
<%'if ano_sel = "" Then
if cd_unidade = "" AND ano_sel = "" Then%>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr><td align=center colspan="100%">&nbsp;</td></tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
					    <td colspan="100%" align="center">Selecione o ano abaixo:</td>
					</tr>
					<form action="<%=nome_arquivo%>" name="Ano" method="get">
					<input type="hidden" name="tipo" value="<%=tipo%>">
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">					
					
					<tr>
						<td align="center" colspan="100%"><br>
						<select name="ano_sel">
								<%ano_inicio = year(dt_inicio_sys)
								response.write(ano_inicio)
								do while not ano_inicio = ano_prod + 1%>
								<option value="<%=ano_inicio%>" <%if ano_inicio=ano_prod then%><%response.write("selected")%><%end if%>><%=ano_inicio%></option>
								<%ano_inicio = ano_inicio + 1
								loop%>			
						</select><br>
						<br>
						<input type="submit" value="OK">						
						</td>
					</tr>
					</form>
<%'**************************
'* 1.2 - Seleção da Unidade *
'****************************
elseif cd_unidade = "" AND ano_sel <> "" Then%>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
						<td colspan="100%"><br>&nbsp;</td>
					</tr>
					<form action="<%=nome_arquivo%>" name="mes" method="get">
					<input type="hidden" name="tipo" value="<%=tipo%>">
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">
					<input type="text" name="ano_sel" value="<%=ano_sel%>">						
					<tr>
					    <td colspan="100%" align="center">Selecione uma Unidade abaixo:</td>
					</tr>
					<tr>
						<td align="center" colspan="100%">
						<%'strsql_unidade = "SELECT DISTINCT A053_codung AS cd_codigo FROM TBL_Protocolo WHERE (YEAR(A002_datpro) = 2009)"
						'Set rs_unidade = dbconn.execute(strsql_unidade)
						
						strsql_unidade = "SELECT DISTINCT TBL_Protocolo.A053_codung AS cd_codigo, TBL_unidades.nm_unidade, TBL_unidades.nm_sigla FROM TBL_Protocolo INNER JOIN TBL_unidades ON TBL_Protocolo.A053_codung = TBL_unidades.cd_codigo WHERE (YEAR(TBL_Protocolo.A002_datpro) ="&ano_sel&")"
						Set rs_unidade = dbconn.execute(strsql_unidade)
						
						'strsql_unidade = "SELECT * FROM TBL_unidades Where cd_status = 1 ORDER BY nm_unidade  ASC"
						'Set rs_unidade = dbconn.execute(strsql_unidade)%>
								
						<select name="cd_unidade">
								<option value="todos">Todas as unidades</option>
								<%while not rs_unidade.EOF
								cd_unidade = rs_unidade("cd_codigo")
								nm_unidade = rs_unidade("nm_unidade")%>
								<option value="<%=cd_unidade%>"><%=nm_unidade%></option>
								<%rs_unidade.movenext
								wend%>			
						</select><br>
						<br>
						<input type="submit" value="OK">						
						</td>
					</tr>
					<tr><td align="center"><input type="checkbox" name="mostra_emprestimos" value="1">Marque para contar os emprestimos</td></tr>
					</form>
					
					<tr>
					    <td colspan="100%">
							<p>&nbsp;</p>
							<p>&nbsp;</p></td>
					</tr>
					<tr>

<%'***********************
'* 1.2 - Corpo da página *
'*************************
Else
'elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo = "" Then%>
<!--tr><td><%'=data_atual%></td></tr-->
<%'end if%>
</table>



<%
sel_ano = ano_sel
	mes_i = 1
	mes_f = 12
	'fim_mes = UltimoDiaMes(mes_f,sel_ano)
	mensal = 1
	numero = 1
	%>

<table align="center" style="border:0px solid gray; font-size: 13px; border-collapse:collapse;">
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
		<td style="border:1px; solid red;" colspan="100%" align="center" ><b>MAPA ANUAL DE CIRURGIAS POR ESPECIALIDADE</b><br><b><%=nm_unidade%></b><br><b><%=sel_ano%></b></td>
	</tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr style="border:1px solid black;">
		<td align="left">&nbsp;<b>Especialidade</b></td>
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
		<td align="right"><b>TOTAL</b></td>
		<td align="right"><b><%=sel_ano-1%></b></td>
		<td align="right"><b>DIF</b>&nbsp;</td>
	</tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%linha_num =  1
	
	strsql = "SELECT * FROM TBL_espec ORDER BY cd_codigo"
	Set rs_espec = dbconn.execute(strsql)
	While not rs_espec.EOF
	cd_especialidae = rs_espec("cd_codigo")
	nm_especialidade = rs_espec("nm_especialidade")
	
		if mostra_emprestimos <> 1 then
			emprestimos = " AND a002_tipage <> 'E' "
		end if
		
	if linha_num = 1 then 
	corlinha = "#dadada"
	else
	corlinha = "#ffffff"
	end if%>
	<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');">
		<td align="left">&nbsp;<b><%=nm_especialidade%></b></td>
		
		<%while not numero = 13
		fim_mes = UltimoDiaMes(numero,sel_ano)
		
		if isnumeric(cd_unidade) then
			strsql_estat = "SELECT COUNT (a002_numpro) AS conta, dt_ano, dt_mes,a053_codung, nm_sigla,a057_codesp,nm_sigla, nm_especialidade FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' AND a053_codung = '"&cd_unidade&"' AND a057_codesp = '"&cd_especialidae&"' "&emprestimos&" Group by dt_ano, dt_mes,a053_codung, nm_sigla, a057_codesp,nm_sigla, nm_especialidade	ORDER BY dt_ano, dt_mes  ASC"
		Else
			strsql_estat = "SELECT COUNT (a002_numpro) AS conta, dt_ano, dt_mes,a057_codesp, nm_especialidade FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&numero&"/1/"&sel_ano&"' AND '"&numero&"/"&fim_mes&"/"&sel_ano&"' AND a057_codesp = '"&cd_especialidae&"' "&emprestimos&" Group by dt_ano, dt_mes, a057_codesp, nm_especialidade ORDER BY dt_ano, dt_mes  ASC"
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
		<%total_esp_ano = total_esp_ano + conta
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
		<td align="right"><%=total_esp_ano%></td>
		<td align="right">
		<%	if isnumeric(cd_unidade) then
				strsql_anoanterior = "SELECT COUNT (a002_numpro) AS conta_anterior, a053_codung, nm_sigla,a057_codesp,nm_sigla, nm_especialidade FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '1/1/"&sel_ano-1&"' AND '12/31/"&sel_ano-1&"' AND a053_codung = '"&cd_unidade&"' AND a057_codesp = '"&cd_especialidae&"'  "&emprestimos&" Group by a053_codung, nm_sigla, a057_codesp,nm_sigla, nm_especialidade ORDER BY nm_especialidade,conta_anterior  ASC"
			else
				strsql_anoanterior = "SELECT COUNT (a002_numpro) AS conta_anterior,a057_codesp, nm_especialidade FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '1/1/"&sel_ano-1&"' AND '12/31/"&sel_ano-1&"' AND a057_codesp = '"&cd_especialidae&"'  "&emprestimos&" Group by a057_codesp, nm_especialidade ORDER BY nm_especialidade,conta_anterior  ASC"
			end if
			
			Set rs_anoanterior = dbconn.execute(strsql_anoanterior)
			if not rs_anoanterior.EOF then
				total_anterior = rs_anoanterior("conta_anterior")
				total_geralanterior = total_geralanterior + total_anterior
			end if%>
			<%=total_anterior%>
		</td>
		<td align="right"><%=total_esp_ano - total_anterior%>
		<%dif_total = dif_total + (total_esp_ano - total_anterior)%>&nbsp;
		</td>
			
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
	
	rs_espec.movenext
	wend
	
	%>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr>
		<td align="left">&nbsp;<b>TOTAL</b></td>
		<td align="right"><%=contagem_jan%></td>
		<td align="right"><%=contagem_fev%></td>
		<td align="right"><%=contagem_mar%></td>
		<td align="right"><%=contagem_abr%></td>
		<td align="right"><%=contagem_mai%></td>
		<td align="right"><%=contagem_jun%></td>
		<td align="right"><%=contagem_jul%></td>
		<td align="right"><%=contagem_ago%></td>
		<td align="right"><%=contagem_set%></td>
		<td align="right"><%=contagem_out%></td>
		<td align="right"><%=contagem_nov%></td>
		<td align="right"><%=contagem_dez%></td>
		<td align="right"><b><%=contagem_jan+contagem_fev+contagem_mar+contagem_abr+contagem_mai+contagem_jun+contagem_jul+contagem_ago+contagem_set+contagem_out+contagem_nov+contagem_dez%></b></td>
		<td align="right"><%=total_geralanterior%></td>
		<td align="right"><%=dif_total%>&nbsp;</td>
	</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr>
		<td align="left">&nbsp;<b><%=sel_ano-1%></b></td>
		<%numero = 1
		
		for i = 1 to 12
		
			ultimodia = UltimoDiaMes(i,sel_ano-1)
			
			if isnumeric(cd_unidade) then
				strsql_anoanterior = "SELECT COUNT (a002_numpro) AS contames_anterior, dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&i&"/1/"&sel_ano-1&"' AND '"&i&"/"&ultimodia&"/"&sel_ano-1&"' AND a053_codung = '"&cd_unidade&"'  "&emprestimos&" Group by dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla ORDER BY dt_ano, dt_mes,contames_anterior  ASC"
			else
				strsql_anoanterior = "SELECT COUNT (a002_numpro) AS contames_anterior, dt_ano, dt_mes, nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&i&"/1/"&sel_ano-1&"' AND '"&i&"/"&ultimodia&"/"&sel_ano-1&"'  "&emprestimos&" Group by dt_ano, dt_mes,nm_sigla ORDER BY dt_ano, dt_mes,contames_anterior  ASC"
			end if
			
			set rs_anoanterior = dbconn.execute(strsql_anoanterior)
			'if not rs_anoanterior.EOF then
			'while not rs_anoanterior.EOF
			if not rs_anoanterior.EOF then
				total_anoanterior = rs_anoanterior("contames_anterior")
				'rs_anoanterior.movenext
			end if
			'wend
			'end if
		%>
		<td align="right"><%=total_anoanterior%><%'=i%></td>
		<%total_geralanoanterior = total_geralanoanterior + total_anoanterior
		
			If i = 1 then
				dif_jan = contagem_jan - total_anoanterior
			end if
			If i = 2 then
				dif_fev = contagem_fev - total_anoanterior
			end if
			If i = 3 then
				dif_mar = contagem_mar - total_anoanterior
			end if
			If i = 4 then
				dif_abr = contagem_abr - total_anoanterior
			end if
			If i = 5 then
				dif_mai = contagem_mai - total_anoanterior
			end if
			If i = 6 then
				dif_jun = contagem_jun - total_anoanterior
			end if
			If i = 7 then
				dif_jul = contagem_jul - total_anoanterior
			end if
			If i = 8 then
				dif_ago = contagem_ago - total_anoanterior
			end if
			If i = 9 then
				dif_set = contagem_set - total_anoanterior
			end if
			If i = 10 then
				dif_out = contagem_out - total_anoanterior
			end if
			If i = 11 then
				dif_nov = contagem_nov - total_anoanterior
			end if
			If i = 12 then
				dif_dez = contagem_dez - total_anoanterior
			end if
		
		next%>
		<td align="right"><%=total_geralanoanterior%></td>
	</tr>
	<%%>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr>
		<td align="left">&nbsp;<b>DIFERENÇA</b></td>
		<td align="right"><%=dif_jan%></td>
		<td align="right"><%=dif_fev%></td>
		<td align="right"><%=dif_mar%></td>
		<td align="right"><%=dif_abr%></td>
		<td align="right"><%=dif_mai%></td>
		<td align="right"><%=dif_jun%></td>
		<td align="right"><%=dif_jul%></td>
		<td align="right"><%=dif_ago%></td>
		<td align="right"><%=dif_set%></td>
		<td align="right"><%=dif_out%></td>
		<td align="right"><%=dif_nov%></td>
		<td align="right"><%=dif_dez%></td>
		<td align="right"><%'=%></td>
	</tr>
	<tr><td colspan="100%" align="center" id="no_print">
	<%if mostra_emprestimos = 1 Then 
		response.write("*** Nos totais estão imbutidos a quantidade de emprestimos ***")
	else
		response.write("*** Isento da quantidade de emprestimos ***")
	end if%> </td></tr>
	<tr>
		<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
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
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="57" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
	</tr>
</table>
<%end if%>

