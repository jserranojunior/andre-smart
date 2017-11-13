


<%
dt_inicio_sys = "01/01/2007"
nome_arquivo = "protocolo.asp"
'nome_arquivo = request("arquivo")

dia_sel = zero(request("dia_sel"))
'mes_sel = zero(request("mes_sel"))
dt_mes = request("dt_mes")
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
if cd_unidade = "" AND dt_mes = "" AND ano_sel = "" Then%>
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
						<td align="center" colspan="100%" valign="baseline"><br>
						<input type="text" name="dt_mes" size="2" maxlength="2" class="inputs">
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
elseif cd_unidade = "" AND dt_mes <> "" AND ano_sel <> "" Then%>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
						<td colspan="100%"><br>&nbsp;</td>
					</tr>
					<form action="<%=nome_arquivo%>" name="mes" method="get">
					<input type="hidden" name="tipo" value="<%=tipo%>">
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">
					<input type="hidden" name="ano_sel" value="<%=ano_sel%>">					
					<input type="hidden" name="dt_mes" value="<%=dt_mes%>">
					<tr>
					    <td colspan="100%" align="center">Selecione uma Unidade abaixo:</td>
					</tr>
					<tr>
						<td align="center" colspan="100%">
						<%strsql_unidade = "SELECT * FROM TBL_unidades Where cd_status = 1 ORDER BY nm_unidade  ASC"
						Set rs_unidade = dbconn.execute(strsql_unidade)%>
								
						<select name="cd_unidade">
								<option value="">&nbsp;</option>
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

elseif cd_unidade <> "" AND dt_mes <> "" AND ano_sel <> "" Then%>
</table>



<%sel_ano = 2009
	'mes_i = 1
	'mes_f = 12
	'fim_mes = UltimoDiaMes(mes_f,sel_ano)
	'mensal = 1
	'numero = 1

'elseif cd_unidade <> "" AND dt_mes <>  "" AND ano_sel <> "" Then%>
<%if int(qtd_paginas) = int(num) OR int(qtd_paginas) = 1 then page_break = "" else page_break = " page-break-after:always;" end if 'response.write(int(num)&".."&int(qtd_paginas))%>
<table align="center" style="border:0px; solid gray; font-size: 13px; border-collapse:collapse;">
	<!--tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr-->
	<%linha_num =  1
	linha_pulo = 2
	numero = 1
	pagina = 1

	strsql = "SELECT a002_numpro, a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, a002_tipage FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage <> 'E' Group by a002_numpro,a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, a002_tipage ORDER BY a002_datpro"
	Set rs = dbconn.execute(strsql)
	
	While not rs.EOF
		a002_numpro = rs("a002_numpro")
	
	if numero = 1 then 
	corlinha = "#dadada"
	else
	corlinha = "#ffffff"
	end if%>	
	<%if linha_pulo =  "2" then%>
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td colspan="3" align="left">&nbsp;<img src="../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></td>
			<td colspan="2" align="center"><b>VD LAP Cirúrgica LTDA</b></td>
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<%strsql_unidade = "SELECT * FROM TBL_unidades Where cd_codigo = '"&cd_unidade&"'"
		Set rs_unidade = dbconn.execute(strsql_unidade)%>
		<%if not rs_unidade.EOF then
		nm_unidade = rs_unidade("nm_unidade")%>	
		<%rs_unidade.movenext
		end if%>	
		<tr style="font-size:11px;">
			<td style="border:1px; solid red;" colspan="100%" align="center" >
				<b>RELATÓRIO DE PROCEDIMENTOS MÉDICOS</b> <br>
				<b><%=nm_unidade%></b><br>
				<b><%=mes_selecionado(dt_mes)%>/<%=sel_ano%></b></td>
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<tr style="border:0px; solid black; font-size:11px;">
			<td align="right">&nbsp;</td>
			<td align="left"><b>Data</b></td>		
			<td align="left"><b>Paciente</b></td>
			<td align="right"><b>Reg</b>&nbsp;</td>
			<td align="left">&nbsp;<b>Cirurgia</b></td>
			<td align="left"><b>Convênio</b></td>
			<td align="left"><b>Equipe Médica</b></td>		
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%end if%>
	
	<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
		<td align="left">&nbsp;<%=linha_num%></td>
		<td align="left"><%=zero(day(rs("a002_datpro")))&"/"&zero(month(rs("a002_datpro")))&"/"&right(year(rs("a002_datpro")),2)%></td>		
		<td align="left"><%a002_pacnom = rs("a002_pacnom")
		nm_pac_len = len(a002_pacnom)
								if nm_pac_len > 38 then
								a002_pacnom = left(a002_pacnom,38)&"..."
								end if%>
		<%=a002_pacnom%></td>
		<td align="right"><%=rs("a002_pacreg")%>&nbsp;</td>
		<td align="left">&nbsp;
			<%strsql_proc = "SELECT * FROM View_protocolo_procedimento WHERE a002_numpro = "&a002_numpro&" AND a053_codung = "&cd_unidade&" ORDER BY a003_propri desc"
			Set rs_proc = dbconn.execute(strsql_proc)%>
				<%if not rs_proc.EOF then
				'while not rs_proc.EOF
					
					a002_numpro = rs_proc("a002_numpro")
					nm_procedimento = rs_proc("nm_procedimento")
					nm_proc_len = len(nm_procedimento)
						if nm_proc_len > 43 then
						nm_procedimento = left(nm_procedimento,43)&"..."
						end if
						a003_propri = rs_proc("a003_propri")
						if a003_propri <> "S" then%><i><font color="red"><%end if%><b><%=a058_codpro%></b></font><%'=lcase(a003_propri)%><%if a003_propri = "s" then%></i><%end if%><%=nm_procedimento%><br>							
						<%'rs_proc.movenext
						'wend
					end if
						rs_proc.close
						set rs_proc = nothing%>
		
		
		<%'=rs("a002_pacnom")%></td>
		<td align="left"><%=rs("nm_convenio")%></td>
		<td align="left"><%=rs("a055_desmed")%></td>		
	</tr>
	<%if linha_pulo = 45 then
	linha_pulo = 1	
	%>
	
	<tr id="print_ok"><td>&nbsp;</td></tr>
	<tr id="print_ok"><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr style="font-size:10px;" id="print_ok"><td colspan="3">&nbsp;<%=qtd_proc_pag%></td>
		<td colspan="2" align="center">Data de emissão: <b><%=zero(day(now))&"/"&mesdoano(month(now))&"/"&right(year(now),2)&" "&hour(now)&":"&minute(now)%></b></td>
		<td colspan="2"align="right">Pág.:<%=pagina%>&nbsp;&nbsp;&nbsp;</td></tr>
	<tr id="print_ok"><td>&nbsp;
	<div style="page-break-after:always;"" align="center"></div></td></tr>
		
	<%pagina = pagina + 1
	
	end if%>
	<%linha_num = linha_num + 1
	qtd_proc = qtd_proc + 1
	'qtd_proc_total = qtd_proc_total +1
		qtd_proc_pag = qtd_proc_pag + 1
		if qtd_proc_pag = 45 then
			qtd_proc_pag = 1
			'qtd_proc = 1
		end if
		
	linha_pulo = linha_pulo + 1
	
	
	numero = numero + 1
			if numero > 2 then
			numero = 1
			end if
	rs.movenext
	Wend
	%>
	
	
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr><td colspan="100%" style="font-size:10px; ">&nbsp;<b>Total de procedimentos: <%=qtd_proc%></b></td></tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%'qtd_proc_pag = qtd_proc_pag + 6%>
	<tr>
		<td style="border:1px; solid red;" colspan="100%" align="center"><br><br><br>
		<b>Empréstimos</b></td>
	</tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<!--tr style="border:1px solid black;">
		<td align="right">&nbsp;</td>
		<td align="left"><b>Data</b></td>		
		<td align="left"><b>Paciente</b></td>
		<td align="right"><b>Reg</b>&nbsp;</td>
		<td align="left">&nbsp;<b>Cirurgia</b></td>
		<td align="left"><b>Convênio</b></td>
		<td align="left"><b>Equipe Médica</b></td>		
	</tr-->
	<%if x = h then%>
	<%linha_num =  1
	numero = 1
	qtd_proc_pag_empr = qtd_proc_pag + 5
	'strsql_unidade = "SELECT * FROM TBL_unidades Where cd_codigo = '"&cd_unidade&"'"
	'Set rs_unidade = dbconn.execute(strsql_unidade)

	'strsql = "SELECT COUNT (a002_numpro) AS conta, a002_datpro,a053_codung, nm_sigla,a002_numpro,a002_pacreg, a002_pacnom,a054_codcon, nm_convenio,a055_crmmed, a055_desmed,a002_tipage,nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '12/1/2009' AND '12/31/2009' AND a053_codung = 3 AND a002_tipage <> 'E' Group by a002_datpro,a053_codung, nm_sigla, a002_numpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_crmmed, a055_desmed, a002_tipage,nm_sigla ORDER BY a002_datpro"	
	'strsql = "SELECT a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, a002_tipage FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = 3 AND a002_tipage <> 'E' Group by a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, a002_tipage ORDER BY a002_datpro"
	strsql = "SELECT COUNT (a002_numpro) AS conta, a002_datpro,a053_codung, nm_sigla,a002_numpro,a002_pacreg, a002_pacnom,a054_codcon, nm_convenio,a055_crmmed, a055_desmed,a002_tipage,nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage = 'E' Group by a002_datpro,a053_codung, nm_sigla, a002_numpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_crmmed, a055_desmed, a002_tipage,nm_sigla ORDER BY a002_datpro"
	
	Set rs = dbconn.execute(strsql)
	
	While not rs.EOF
	
	if numero = 1 then 
	corlinha = "#dadada"
	else
	corlinha = "#ffffff"
	end if%>
	<%if linha_pulo_empr =  "2" then%>
		<tr><td>&nbsp;</td></tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<tr style="border:0px; solid black; font-size:11px;">
			<td align="right">&nbsp;</td>
			<td align="left"><b>Data</b></td>		
			<td align="left"><b>Paciente</b></td>
			<td align="right"><b>Reg</b>&nbsp;</td>
			<td align="left">&nbsp;<b>Cirurgia</b></td>
			<td align="left"><b>Convênio</b></td>
			<td align="left"><b>Equipe Médica</b></td>		
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%end if%>
	<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
		<td align="left">&nbsp;<%=linha_num%></td>
		<td align="left"><%=zero(day(rs("a002_datpro")))&"/"&zero(month(rs("a002_datpro")))&"/"&right(year(rs("a002_datpro")),2)%></td>		
		<td align="left"><%a002_pacnom = rs("a002_pacnom")
		nm_pac_len = len(a002_pacnom)
								if nm_pac_len > 38 then
								a002_pacnom = left(a002_pacnom,38)&"..."
								end if%>
		<%=a002_pacnom%></td>
		<td align="right"><%=rs("a002_pacreg")%>&nbsp;</td>
		<td align="left">&nbsp;
		<%strsql_proc = "SELECT * FROM View_protocolo_procedimento WHERE a002_numpro = "&a002_numpro&" AND a053_codung = "&cd_unidade&" ORDER BY a003_propri desc"
			Set rs_proc = dbconn.execute(strsql_proc)%>
				<%if not rs_proc.EOF then
				'while not rs_proc.EOF
					
					a002_numpro = rs_proc("a002_numpro")
					nm_procedimento = rs_proc("nm_procedimento")
					nm_proc_len = len(nm_procedimento)
						if nm_proc_len > 43 then
						nm_procedimento = left(nm_procedimento,43)&"..."
						end if
						a003_propri = rs_proc("a003_propri")
						if a003_propri <> "S" then%><i><font color="red"><%end if%><b><%=a058_codpro%></b></font><%'=lcase(a003_propri)%><%if a003_propri = "s" then%></i><%end if%><%=nm_procedimento%><br>							
						<%'rs_proc.movenext
						'wend
					end if
						rs_proc.close
						set rs_proc = nothing%></td>
		<td align="left"><%=rs("nm_convenio")%></td>
		<td align="left"><%=rs("a055_desmed")%></td>		
	</tr>	
	<%linha_num = linha_num + 1
	qtd_empr = qtd_empr + 1
	qtd_proc_pag_empr = qtd_proc_pag_empr + 1
	
		qtd_proc_pag = qtd_proc_pag + 1
		if qtd_proc_pag = 45 then
			qtd_proc_pag = 1
			'qtd_proc = 1
		end if
		
	linha_pulo = linha_pulo + 1
	
	numero = numero + 1
			if numero > 2 then
			numero = 1
			end if
	rs.movenext
	Wend
	%>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"><%=qtd_proc_pag_empr%></td></tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr><td colspan="100%" style="font-size:10px; ">&nbsp;<b>Total de empréstimos: <%=qtd_empr%></b></td></tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%end if%>
	
	<%if linha_pulo = 45 then
	linha_pulo = 1	
	%>
	
	<tr id="print_ok"><td>&nbsp;</td></tr>
	<tr id="print_ok"><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr style="font-size:10px;" id="print_ok"><td colspan="3">&nbsp;<%=qtd_proc_pag%></td>
		<td colspan="2" align="center">Data de emissão: <b><%=zero(day(now))&"/"&mesdoano(month(now))&"/"&right(year(now),2)&" "&hour(now)&":"&minute(now)%></b></td>
		<td colspan="2"align="right">Pág.:<%=pagina%>&nbsp;&nbsp;&nbsp;</td></tr>
	<tr id="print_ok"><td>&nbsp;
	<div style="page-break-after:always;"" align="center"></div></td></tr>
		
	<%pagina = pagina + 1
	
	end if%>
	<tr>		
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="175" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="135" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="140" height="1" border="0"></td>
	</tr>
</table>

<%end if%>
