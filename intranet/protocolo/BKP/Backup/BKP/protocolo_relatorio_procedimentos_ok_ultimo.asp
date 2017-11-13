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
<%if int(qtd_paginas) = int(num) OR int(qtd_paginas) = 1 then page_break = "" else page_break = " page-break-after:always;" end if 'response.write(int(num)&".."&int(qtd_paginas))%>
<table align="center" style="border:0px; solid gray; font-size: 13px; border-collapse:collapse;">
	<%'******************************************************
	'***     Define a quantidade de linhas por página     ***
	'********************************************************
	'*** Use 41 para impressão em folha Carta e 46 para A4 ***
	
		qtd_linhas_pag = 46
	'************************************************
	
	strsql = "SELECT COUNT (a002_numpro) AS pre_conta, dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' Group by dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla ORDER BY dt_ano, dt_mes"
	Set rs_conta = dbconn.execute(strsql)
	if not rs_conta.EOF Then
		pre_conta = rs_conta("pre_conta")
			total_final_pag = round(pre_conta / qtd_linhas_pag)
	end if
	
	if cd_unidade = 20 then '"&regra_edmundo &"
		rack_ed = "a056_codrac, a056_desrac,"
		 '"&ordena_rack&"'
	end if
	
	strsql = "SELECT a002_numpro, a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, "&rack_ed&" a002_tipage FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage <> 'E' Group by a002_numpro,a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, "&rack_ed&" a002_tipage ORDER BY "&rack_ed&" a002_datpro"
		Set rs = dbconn.execute(strsql)
		
	strsql_e = "SELECT COUNT (a002_numpro) AS conta, a002_datpro,a053_codung, nm_sigla,a002_numpro,a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_crmmed, a055_desmed, a002_tipage, "&rack_ed&" nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage = 'E' Group by a002_datpro,a053_codung, nm_sigla, a002_numpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_crmmed, a055_desmed, a002_tipage, "&rack_ed&" nm_sigla ORDER BY "&rack_ed&" a002_datpro"
			Set rs_e = dbconn.execute(strsql_e)%>
	
	<%mostra_cab = 1
	mostra_rod = 1
	mostra_empr = 0
	last_rack = 0
	
	linha_numerar =  1
	linha_pagina = 1
	pagina = 1
	qtd_proc = 1
	
	for i = 1 to pre_conta
	
	if not rs.EOF then
			dt_data = zero(day(rs("a002_datpro")))&"/"&zero(month(rs("a002_datpro")))&"/"&right(year(rs("a002_datpro")),2)
			a002_pacnom = rs("a002_pacnom")
				a002_numpro = rs("a002_numpro")
				nm_pac_len = len(a002_pacnom)
					if nm_pac_len > 34 then
						a002_pacnom = left(a002_pacnom,34)&"..."
					end if
			a002_pacreg = rs("a002_pacreg")
			nm_convenio = rs("nm_convenio")
			a055_desmed = rs("a055_desmed")
			if cd_unidade = 20 then 
				a056_desrac = rs("a056_desrac")
				a056_codrac = rs("a056_codrac")
			end if
		
		rs.movenext
	Elseif not rs_e.EOF then
		'if not rs_e.EOF Then
			dt_data = zero(day(rs_e("a002_datpro")))&"/"&zero(month(rs_e("a002_datpro")))&"/"&right(year(rs_e("a002_datpro")),2)
			a002_pacnom = rs_e("a002_pacnom")
				a002_numpro = rs_e("a002_numpro")
				nm_pac_len = len(a002_pacnom)
					if nm_pac_len > 34 then
						a002_pacnom = left(a002_pacnom,34)&"..."
					end if
			a002_pacreg = rs_e("a002_pacreg")
			nm_convenio = rs_e("nm_convenio")
			a055_desmed = rs_e("a055_desmed")
				nm_med_len = len(a055_desmed)
					if nm_med_len > 28 then
						a055_desmed = left(a055_desmed,280&"...")
					end if
			if cd_unidade = 20 then 
				a056_codrac = rs_e("a056_codrac")
				a056_desrac = rs_e("a056_desrac")
			end if
			
		rs_e.movenext
		'end if
	end if%>
	
<!-- ***************************************************
***													 ***
*** INICIO DO RELATÓRIO DE CIRURGIAS E PROCEDIMENTOS ***
***													 *** 
******************************************************-->	
	<%if mostra_cab = 1 then%>
	<!--tr><td align="center">Cabeçalho</td></tr-->
		<tr>
			<td colspan="3" align="left">&nbsp;<img src="../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></td>
			<td colspan="2" align="center"><b>Vd Lap Cirúrgica LTDA</b></td>
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
				<b><%=mes_selecionado(dt_mes)%>/<%=ano_sel%></b></td>
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<%end if%>
		
		<%if int(last_rack) <> int(a056_codrac) AND cd_unidade = 20 or int(mostra_cab) = 1 AND cd_unidade = 20 then

		%>
		<tr><td colspan="100%">&nbsp;<%'=mostra_rod%>  <%'=mostra_rod+3%></td></tr>
		<!--tr>		
			<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="175" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="135" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="140" height="1" border="0"></td>
		</tr-->
		<tr><td style="border:1px; solid red;" colspan="100%" align="center" >
			<%'=qtd_vez%><%'=qtd_vez'linha_pagina&"/"& mostra_rod+3%> Procedimentos Médicos Realizados com <%=a056_desrac%> <%'=a056_codrac%><br>		
		</td></tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<tr style="border:0px; solid black; font-size:8px;">
			<td align="right"><b>&nbsp;</b></td>
			<td align="left"><b>Data</b></td>		
			<td align="left"><b>Paciente</b></td>
			<td align="right"><b>Reg</b>&nbsp;</td>
			<td align="left">&nbsp;<b>Cirurgia</b></td>
			<td align="left"><b>Convênio</b></td>
			<td align="left"><b>Equipe Médica</b></td>		
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<%mostra_cab = mostra_cab + 1
		mostra_rod =  mostra_rod + 4 
		
			if qtd_vez => 1 then
				mostra_rod = mostra_rod - 1
			end if
		qtd_vez = qtd_vez + 1
		end if
				
			if int(last_rack) <> int(a056_codrac) AND cd_unidade = 20 then
				linha_numerar = 1
			end if
		%>	
		
		<%if mostra_cab = 1  AND cd_unidade <> 20 then%>
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
	<% mostra_cab = mostra_cab + 1
	end if%>	
	
	<%if numero = 1 then
	corlinha = "#dadada"
	else
	corlinha = "#ffffff"
	end if%>	
	<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
		<td align="left">&nbsp;<%'="("&mostra_rod&")"%><%'="("&linha_pagina&":"&mostra_rod&")"%><%=linha_numerar%></td>		
		<td align="left"><%=dt_data%></td>		
		<td align="left"><%=a002_pacnom%></td>
		<td align="right"><%=a002_pacreg%>&nbsp;</td>
		<td align="left">&nbsp;
			<%'if x = "bc"  then
			strsql_proc = "SELECT * FROM View_protocolo_procedimento WHERE a002_numpro = '"&a002_numpro&"' AND a053_codung = "&cd_unidade&" ORDER BY a003_propri desc"
			Set rs_proc = dbconn.execute(strsql_proc)%>
				<%if not rs_proc.EOF then
				'while not rs_proc.EOF
					
					a002_numpro = rs_proc("a002_numpro")
					nm_procedimento = rs_proc("nm_procedimento")
					nm_proc_len = len(nm_procedimento)
						if nm_proc_len > 39 then
						nm_procedimento = left(nm_procedimento,39)&"..."
						end if
						a003_propri = rs_proc("a003_propri")
						if a003_propri <> "S" then%><!--i--><font color="red"><%end if%><b><%=a058_codpro%></b></font><%'=lcase(a003_propri)%><%if a003_propri = "s" then%><!--/i--><%end if%><%=nm_procedimento%><br>							
						<%'rs_proc.movenext
						'wend
					end if
						rs_proc.close
						set rs_proc = nothing
				'end if%></td>
				
			<%if emprestimo <> 1  then%>
				<td align="left"><%=nm_convenio%></td>
			<%else%>
				<td align="left">&nbsp;<%'=nm_convenio%></td>
				<td align="left">&nbsp;<%'=nm_convenio%></td>			
			<%end if%>
		
		<td align="left"><%=a055_desmed%></td>
		<!--td align="left"><%'=a056_codrac%></td-->
	</tr>
	<%if mostra_empr = 0 AND rs.EOF then%>
	<%mostra_rod = mostra_rod + 1%>
	
		<%if cd_unidade = 20 then
			for i_ed = 1 to 5%>				
				<tr bgcolor="ffffff<%'=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'ffffff<%'=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
					<td>&nbsp;<%=mostra_rod%></td>
				</tr>
				<%' linha_pagina = linha_pagina + 1
				mostra_rod = mostra_rod + 1%>
			<%next%>
			
			
		<%end if%>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr><td colspan="100%" style="font-size:10px; "><%'=mostra_rod%>&nbsp;<b>Total de procedimentos: <%=qtd_proc%></b></td></tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	
	
	
	
	<%if rs.EOF AND not rs_e.EOF AND qtd_proc_e > 0 then
	mostra_rod = mostra_rod + 2%>
	</table>
	<table align="center">
	<!--table align="center" style="border:0px; solid gray; font-size: 13px; border-collapse:collapse;"-->
	<tr><td colspan="100%"><%'=mostra_rod-1%>&nbsp;</td></tr>
	<tr><td style="border:1px; solid red;" colspan="100%" align="center" ><%'=mostra_rod%><b>RELATÓRIO DE EMPRÉSTIMOS</b></b></td></tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr style="border:0px; solid black; font-size:11px;">
		<td align="right"><%'=mostra_rod%>&nbsp;</td>
		<td align="left"><b>Data</b></td>		
		<td align="left"><b>Paciente</b></td>
		<td align="right"><b>Reg</b>&nbsp;</td>
		<td align="left">&nbsp;<b>Cirurgia</b></td>
		<td align="left"><b>Qtd</b></td>
		<td align="left"><b>Material</b></td>
		<td align="left"><b>Equipe Médica</b></td>		
	</tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%emprestimo = 1
	end if
	qtd_proc_e = 0
	linha_pagina = linha_pagina + 5%>
	
	<%mostra_empr = mostra_empr + 1	
	end if%>
	<%x = "adx"
	
	if rs.EOF AND qtd_proc_e > 0 AND a002_numpro <> ""  AND x = "adx" then%>
	<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffffcc');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
		<td colspan="5"></td>		
			<%strsql_mat = "SELECT * FROM View_protocolo_material WHERE a002_numseq = '"&a002_numpro&"' AND a053_codung = "&cd_unidade&" ORDER BY a010_quantidade desc"
			Set rs_mat = dbconn.execute(strsql_mat)%>
				<%qtd_linhas_mat = 1
				
				while not rs_mat.EOF
					
					'a002_numpro = rs_mat("a002_numseq")
					qtd_material = rs_mat("a010_quantidade")
					nm_material = rs_mat("a061_nommat")
						nm_mat_len = len(nm_material)
							if nm_proc_len > 43 then
							nm_material = left(nm_material,43)&"..."
							end if%>
		<%if qtd_linhas_mat > 1	then%>
		</tr>
		
		
		
		
		<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffffcc');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
			<td colspan="3">&nbsp;</td>
		<%end if%>			
			<td>&nbsp;<%=zero(qtd_material)%></td>
			<td>&nbsp;<%=nm_material%></td>		
					<%qtd_linhas_mat = qtd_linhas_mat +1
					numero = numero + 1
				rs_mat.movenext
				wend
				
				
				rs_mat.close
				set rs_mat = nothing%>
			<td>&nbsp;</td>
	</tr>
	
	
	
	
	
	<tr><td>&nbsp;</td></tr>
	<%end if%>
	<%if rs.EOF AND rs_e.EOF AND qtd_proc_e > 0 then%>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr><td colspan="100%" style="font-size:10px; "><%'=mostra_rod%>&nbsp;<b>Total de empréstimos: <%=qtd_proc_e%></b></td></tr>
	<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%mostra_empr = mostra_empr + 1	
	end if%>
		
	<%mostra_rod = mostra_rod + 1 
	last_rack = a056_codrac
	'mostra_empr = mostra_empr + 1
	if mostra_rod = qtd_linhas_pag OR i = pre_conta then
	
			if pagina >= total_final_pag then '*** Repete as linhas em branco para posicionar o rodapé 
				while not int(mostra_rod+2) = qtd_linhas_pag%>				
				<tr bgcolor="ffffff<%'=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'ffffff<%'=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
					<td>&nbsp;<%'=mostra_rod%></td>
				</tr>
				<% linha_pagina = linha_pagina + 1
				mostra_rod = mostra_rod + 1
				wend
			end if%>
	
	<!--tr id="print_ok"><td>&nbsp;</td></tr-->
	<tr>		
		<td><img src="../imagens/px.gif" alt="" width="20" height="20" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="175" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
	<%if emprestimo <> 1 then%>	
		<td><img src="../imagens/px.gif" alt="" width="135" height="1" border="0"></td>
	<%else%>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="105" height="1" border="0"></td>
	<%end if%>
		<td><img src="../imagens/px.gif" alt="" width="140" height="1" border="0"></td>
	</tr>
	<tr id="print_ok"><td colspan="100%"><%'=total_final_pag&"///"&pagina&"/*/"&linha_pagina+1&"///"&qtd_linhas_pag%><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr style="font-size:10px;" id="print_ok"><td colspan="3">&nbsp;<%'=mostra_rod &"?"&qtd_linhas_pag&"......"&total_final_pag&"-"&pagina%></td>
		<td colspan="2" align="center">Data de emissão: <b><%=zero(day(now))&"/"&mesdoano(month(now))&"/"&right(year(now),2)&" "&hour(now)&":"&minute(now)%></b></td>
		<td colspan="2"align="right">Pág.:<%=pagina%>&nbsp;&nbsp;&nbsp;<%'=mostra_empr%></td></tr>
	<tr id="no_print"><td>&nbsp;</td></tr>
		<%if mostra_rod = qtd_linhas_pag OR i = pre_conta - 1 then%>
		<tr style="page-break-after:always;"><td align="right">&nbsp</td></tr>		
		<%end if%>
	
	<%mostra_rod = 1
	mostra_cab = 1
	pagina = pagina + 1
	linha_pagina = 0
	qtd_vez = 0
	
	end if%>	
	
	<%linha_numerar = linha_numerar + 1
	linha_pagina = linha_pagina + 1
	
	qtd_proc = qtd_proc + 1
	qtd_proc_e = qtd_proc_e + 1
	qtd_proc_pag = qtd_proc_pag + 1
		if qtd_proc_pag = qtd_linhas_pag then
			qtd_proc_pag = 1
		end if
		
	linha_pulo = linha_pulo + 1
	
	numero = numero + 1
			if numero > 2 then
			numero = 1
			end if%>
	<%next%>
</table>
<%end if%>
