<%
dt_inicio_sys = "01/01/2007"
nome_arquivo = "protocolo.asp"
'nome_arquivo = request("arquivo")

dia_sel = zero(request("dia_sel"))
'mes_sel = zero(request("mes_sel"))
dt_mes = request("dt_mes")
ano_sel = request("ano_sel")

dia_sel_i = zero(request("dia_sel_i"))
mes_sel_i = request("mes_sel_i")
ano_sel_i = request("ano_sel_i")

dia_sel_f = zero(request("dia_sel_f"))
mes_sel_f = request("mes_sel_f")
ano_sel_f = request("ano_sel_f")
	if mes_sel_i = "" then
		dia_sel_i = day(now)
		mes_sel_i = month(now)
		ano_sel_i = year(now)
		
		dia_sel_f = day(now)
		mes_sel_f = month(now)
		ano_sel_f = year(now)
	end if

data_atual = dia_sel&"/"&mes_sel&"/"&ano_sel

dia_prod = zero(day(now))
mes_prod = zero(month(now))
ano_prod = year(now)
data_prod = dia_prod&"/"&mes_prod&"/"&ano_prod

if year(dt_inicio_sys) = int(ano_prod) then
	ano_sel = year(dt_inicio_sys)
end if

cd_unidade = request("cd_unidade")
cd_medico = request("cd_medico")
	if cd_medico = 0 Then
		busca_medico = ""
	else
		busca_medico = " AND (A055_codmed = '"&cd_medico&"') "
	end if

mostra_emprestimos = request("mostra_emprestimos")


'*********************************************************
'*				      1ª Parte	 					  	'*
'*********************************************************
'* 1.1 - Seleção do ANO *
'************************%>
<table border="1" align="center">
<%'***********************
'* 1.2 - Corpo da página *
'*************************

'if cd_unidade <> "" AND mes_sel_i <> "" AND ano_sel_i <> "" Then%>
</table>
<%if int(qtd_paginas) = int(num) OR int(qtd_paginas) = 1 then page_break = "" else page_break = " page-break-after:always;" end if%>
		<%strsql_unidade = "SELECT * FROM TBL_unidades Where cd_codigo = '"&cd_unidade&"'"
		Set rs_unidade = dbconn.execute(strsql_unidade)%>
		<%if not rs_unidade.EOF then
			nm_unidade = rs_unidade("nm_unidade")%>	
		<%rs_unidade.movenext
		end if%>
<br class="no_print">
<br class="no_print">		
<table align="center" style="border:1px solid gray; font-size: 13px; border-collapse:collapse;" class="no_print">	
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">protocolo_relatorio_emprestimos.asp</span></td></tr>
	<%end if%>
	<tr>
		<td align="center" style="border:1px solid gray; font-size: 15px; border-collapse:collapse;" colspan="2"><b>Relatório de Cirurgia</b></td>
	</tr>
	<form action="protocolo.asp?tipo=emprestimos">
	<input type="hidden" name="tipo" value="emprestimos">
	<tr>
		<td>Unidade</td>
		<td align="left" style="color:red;">
			<%strsql_unidade = "SELECT * FROM TBL_unidades Where cd_status = 1 ORDER BY nm_unidade  ASC"
						Set rs_unidade = dbconn.execute(strsql_unidade)%>
				<select name="cd_unidade">
						<option value="">&nbsp;</option>
						<%while not rs_unidade.EOF
						strcd_unidade = rs_unidade("cd_codigo")
						strnm_unidade = rs_unidade("nm_unidade")
						%>
						<option value="<%=strcd_unidade%>" <%if strcd_unidade = int(cd_unidade) then response.write("selected") end if%>><%=strnm_unidade%></option>
						<%rs_unidade.movenext
						wend%>
				</select></td>
	</tr>
	<tr>
		<td>Período</td>
		<td>
			<select name="dia_sel_i" class="inputs">
							<%numero = 1
							do while NOT numero = 32
							if int(dia_sel_i) = numero Then
							check = "selected"
							end if%>					
							<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<select name="mes_sel_i" class="inputs">
							<%numero = 1
							do while NOT numero = 13
							if int(mes_sel_i) = numero Then
							check = "selected"
							end if%>					
							<option value="<%=numero%>" <%=check%>><%=left(mes_selecionado(numero),3)%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<%if ano_sel_i = "" then%><%ano_sel_i=ano_atual%><%end if%>
						<input type="text" name="ano_sel_i" size="4" maxlength="4" value="<%=ano_sel_i%>" class="inputs">
						até
						<select name="dia_sel_f" class="inputs">
							<%numero = 1
							do while NOT numero = 32
								if int(dia_sel_f) = numero Then
								check = "selected"								
								end if%>					
							<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<select name="mes_sel_f" class="inputs">
							<%numero = 1
							do while NOT numero = 13
								if mes_sel_f = "" Then
									mes_sel_f = 12
									if mes_hoje = numero Then
									check = "selected"
									end if
								Elseif int(mes_sel_f) = numero Then
									check = "selected"
								end if%>		
							<option value="<%=numero%>" <%=check%>><%=left(mes_selecionado(numero),3)%></option>
							<%numero = numero+1
						check = ""
						loop%>							
						</select> /
						<%if ano_sel_f = "" then%><%ano_f=ano_atual%><%end if%> 
						<input type="text" name="ano_sel_f" size="4" maxlength="4" value="<%=ano_sel_f%>" class="inputs"></td></tr>
		<tr>
			<td>Cirurgião</td>
			<td><%strsql_medico = "SELECT * FROM TBL_medicos Where a055_status = 1 ORDER BY a055_desmed ASC"
						Set rs_medico = dbconn.execute(strsql_medico)%>
				<select name="cd_medico">
						<option value="0">&nbsp;</option>
						<%while not rs_medico.EOF
						strcd_medico = int(rs_medico("a055_codmed"))
						strnm_medico = rs_medico("a055_desmed")
						%>
						<option value="<%=strcd_medico%>" <%if strcd_medico = int(cd_medico) then response.write("selected") end if%>><%=strnm_medico%></option>
						<%rs_medico.movenext
						wend%>
				</select>
				
			</td>
		</tr>
		<tr><td colspan="2" align="center"><input type="submit" value="Buscar"> &nbsp; <img src="../imagens/ic_print.gif" alt="" border="0" onClick="window.print();"> &nbsp; <img src="../imagens/ic_print_view.gif" alt="" border="0" onclick="visualizarImpressao();"></td></tr>
	</form>
	<tr><td colspan="2"><hr></td></tr>
	<tr>
		<td align="center" colspan="2"><br><input type="button" onclick="visualizarImpressao();" value="Visualizar" id="inputs"><br>
		<br>
		<br>
		<input type="button" onClick="window.print()" value="IMPRIMIR" id="inputs"></td>
		
	</tr>
	<tr><td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td></tr>
	<!--tr><td colspan="100">SELECT cd_protocolo, a002_datpro,a053_codung, nm_sigla,a002_numpro,a002_pacreg, a002_pacnom,a054_codcon, nm_convenio,a055_crmmed, a055_desmed,a002_tipage,nm_sigla, a055_desmed, a055_crmmed <br>
	FROM VIEW_protocolo_lista 	<br>
	WHERE A002_datpro BETWEEN '<%=mes_sel_i&"/"&dia_sel_i&"/"&ano_sel_i%>' AND '<%=mes_sel_f&"/"&dia_sel_f&"/"&ano_sel_f%>' AND a053_codung = <%=cd_unidade%>  <%=busca_medico%> AND a002_tipage = 'E' 	<br>
	Group by a002_datpro,a053_codung, nm_sigla, a002_numpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_crmmed, a055_desmed, a002_tipage,nm_sigla, a055_desmed, a055_crmmed, cd_protocolo 	ORDER BY a002_datpro</td></tr-->
</table>
<%if cd_unidade <> "" OR cd_medico <> "" OR dia_sel_i <> "" OR mes_sel_i <> "" OR ano_sel_i <> "" OR dia_sel_f <> "" OR mes_sel_f <> "" OR ano_sel_f <> ""  then%>
<table align="center" style="border:0px solid gray; font-size: 13px; border-collapse:collapse;" id="oks_print">
	<%'******************************************************
	'***     Define a quantidade de linhas por página     ***
	'********************************************************
	'*** Use 41 para impressão em folha Carta e 46 para A4 ***
	
		qtd_linhas_pag = 45
	
	'*********************************
	'*** Contagem de procedimentos ***
	'*********************************
	
		
	'A.2 - Conta todos os emprestimos
		'strsql = "SELECT COUNT (a002_numpro) AS pre_conta, dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage = 'E' Group by dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla ORDER BY dt_ano, dt_mes"
		'strsql = "SELECT COUNT (a061_codmat) AS pre_conta , cd_protocolo, a002_datpro,a053_codung, nm_sigla,a002_numpro,a002_pacreg, a002_pacnom,a054_codcon, nm_convenio,a055_crmmed, a055_desmed,a002_tipage,nm_sigla, a055_desmed, a055_crmmed FROM VIEW_material_protocolo_lista Where A002_datpro BETWEEN '10/1/2009' AND '4/30/2010' AND a053_codung = 3 AND a055_desmed like '%duarte%' AND a002_tipage = 'E' Group by a002_datpro,a053_codung, nm_sigla, a002_numpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_crmmed, a055_desmed, a002_tipage,nm_sigla, a055_desmed, a055_crmmed, cd_protocolo ORDER BY a002_datpro"
		'strsql = "SELECT     COUNT(a061_codmat) AS pre_conta FROM         VIEW_material_protocolo_lista WHERE     (A002_datpro BETWEEN '10/1/2009' AND '3/31/2010') AND (A053_codung = 3) AND (A055_desmed LIKE '%duarte%') AND (a002_tipage = 'E')"
		
		strsql = "SELECT     COUNT(a002_numpro) AS pre_conta FROM VIEW_protocolo_lista WHERE (A002_datpro BETWEEN '"&mes_sel_i&"/"&dia_sel_i&"/"&ano_sel_i&"' AND '"&mes_sel_f&"/"&dia_sel_f&"/"&ano_sel_f&"') AND (A053_codung = '"&cd_unidade&"') "&busca_medico&" AND (a002_tipage = 'E')"
			Set rs_conta = dbconn.execute(strsql)
			if not rs_conta.EOF Then
				pre_conta_e = rs_conta("pre_conta")
			end if
		
		
		
			'*** Fim das contagens de procedimentos	
			'procedimentos = pre_conta - pre_conta_e' - pre_conta_co
			emprestimos = pre_conta_e			
			'proc_co = pre_conta_co
			Total_procedimentos = procedimentos + emprestimos + proc_co
			
				'Regra Edmundo Vasconcelos ******************
					if cd_unidade = 20 then 
						rack_ed = "a056_codrac, a056_desrac,"
					end if
			'************************************************%>

<%
'if emprestimos <> "" Then%>
<%'******************************************************************************************************
'********************************************************************************************************
	'B.3 - Mostra os procedimentos - Empréstimos
'********************************************************************************************************
'********************************************************************************************************
	
	'strsql = "SELECT a002_numpro, a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, cd_protocolo, "&rack_ed&" a002_tipage = 'E' FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage = 'E' Group by a002_numpro,a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, cd_protocolo, "&rack_ed&" a002_tipage ORDER BY "&rack_ed&" a002_datpro"
	strsql = "	SELECT cd_protocolo, a002_datpro,a053_codung, nm_sigla,a002_numpro,a002_pacreg, a002_pacnom,a054_codcon, nm_convenio,a055_codmed, a055_desmed,a002_tipage,nm_sigla, a055_desmed, a055_crmmed 	FROM VIEW_protocolo_lista 	WHERE A002_datpro BETWEEN '"&mes_sel_i&"/"&dia_sel_i&"/"&ano_sel_i&"' AND '"&mes_sel_f&"/"&dia_sel_f&"/"&ano_sel_f&"' AND a053_codung = '"&cd_unidade&"' "&busca_medico&" AND a002_tipage = 'E' GROUP BY a002_datpro,a053_codung, nm_sigla, a002_numpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_codmed, a055_desmed, a002_tipage,nm_sigla, a055_desmed, a055_crmmed, cd_protocolo 	ORDER BY a055_codmed,a002_datpro"
	Set rs = dbconn.execute(strsql)
	
	linha_numerar_e =  1
	conta_linha_e = 0
	rack_numerar_e = 1
	reg_num_e = 0
	
	for i_e = 1 to emprestimos
	
		if numero = 1 then
			corlinha = "#ffffff"
		else
			corlinha = "#dadada"
		end if
		
		cd_protocolo = rs("cd_protocolo")
		a002_numpro = rs("a002_numpro")
		dt_data = zero(day(rs("a002_datpro")))&"/"&zero(month(rs("a002_datpro")))&"/"&right(year(rs("a002_datpro")),2)
		a002_pacnom = rs("a002_pacnom")
			nm_pac_len = len(a002_pacnom)
				if nm_pac_len > 34 then
					a002_pacnom = left(a002_pacnom,34)&"..."
				end if
		a055_codmed = rs("a055_codmed")
		a055_desmed = rs("a055_desmed")
		nm_convenio = rs("nm_convenio")%>			
		
			<%if conta_linha_e < 1 then%>
			<tr>
				<td colspan="3" align="left">&nbsp;<img src="../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></td>
				<td colspan="3" align="center"><b>Vd Lap Cirúrgica LTDA</b></td>
			</tr>
			<tr>		
			<!--td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td-->
			<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="160" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="125" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="140" height="1" border="0"></td>
		</tr>
			<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
			<tr style="font-size:11px;">
				<td style="border:1px; solid red;" colspan="100%" align="center" >
					<b>RELATÓRIO DE EMPRÉSTIMO DE MATERIAIS</b> <br>
					<%if cd_medico > 0 Then%><b>Equipe: Dr. <%=a055_desmed%></b><br><%end if%>					
					<b>De <%=dia_sel_i%>/<%=mes_selecionado(int(mes_sel_i))%>/<%=ano_sel_i%> até <%=dia_sel_f%>/<%=mes_selecionado(int(mes_sel_f))%>/<%=ano_sel_f%></b><br>
					<b><%=nm_unidade%></b><br>
					&nbsp;</td>
			</tr>
			<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%if cd_unidade = 20 then%>
					<tr><td align="center" colspan="7"><%=a056_desrac%></td></tr>
				<%end if%>
			
				<%if cd_medico > 0 then%>
				<tr style="border:0px; solid black; font-size:11px;">
					<td align="right">&nbsp;</td>
					<td align="left"><b>Data</b></td>		
					<td align="left"><b>Paciente</b></td>
					<td align="right"><b>Reg</b>&nbsp;</td>
					<td align="left">&nbsp;<b>Cirurgia</b></td>
					<td align="left" colspan=2><b>Convênio</b></td>
					<td align="left"><b>Qtd. - ITENS</b></td>
					<!--td align="left"><b></b></td-->						
				</tr>
				<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%end if%>
			
			<%reg_num_e = reg_num_e + 1
			'conta_linha_e = 1
			end if%>
			
				<%'***********************************************************************
				'*** se o relatorio abranger todos os medicos, mostra o nome da equipe ***
				'*************************************************************************
				if cd_medico = 0 AND  a055_codmed <> medico_anterior Then%>
					<tr><td>&nbsp;</td></tr>
					<tr><td colspan="8" align="center"><b>Equipe: Dr. <%=a055_desmed%></b><br></td></tr>
					<tr><td colspan="8"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				
				<tr style="border:0px; solid black; font-size:11px;">
					<td align="right">&nbsp;</td>
					<td align="left"><b>Data</b></td>		
					<td align="left"><b>Paciente</b></td>
					<td align="right"><b>Reg</b>&nbsp;</td>
					<td align="left">&nbsp;<b>Cirurgia</b></td>
					<td align="left" colspan=2><b>Convênio</b></td>
					<td align="left"><b>Qtd. - ITENS</b></td>
					<!--td align="left"><b></b></td-->						
				</tr>
				<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%end if%>
			
			
				<tr bgcolor="<%=corlinha%>" valign="top" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
					<!--td align="center"><%=conta_linha_e%>/<%=qtd_linhas_pag%></td-->
					<td align="right"><%=linha_numerar_e%><%'=conta_linha_e%><%'=conta_linha_e=qtd_linhas_pag%></td>					
					<td>&nbsp;<%=dt_data%></td>
					<td>&nbsp;<%=a002_pacnom%> <%'=teste%></td>
					<!--td><%'=a055_codmed%>/<%'=medico_anterior%></td-->
					<td align="right"><%=a002_numpro%></td>
					<%strsql_proc = "SELECT * FROM View_protocolo_procedimento WHERE cd_protocolo = '"&cd_protocolo&"' ORDER BY a003_propri desc"
					Set rs_proc = dbconn.execute(strsql_proc)
					if not rs_proc.EOF Then
						nm_procedimento = rs_proc("nm_procedimento")
							nm_proc_len = len(nm_procedimento)
							if nm_proc_len > 35 then
								nm_procedimento = left(nm_procedimento,35)&"..."
								end if
						a003_propri = rs_proc("a003_propri")
					end if%>
					<td>&nbsp;<%=nm_procedimento%></td>
					<td colspan="2">&nbsp;<%=nm_convenio%></td>
					<%strsql_mat = "SELECT * FROM View_protocolo_material WHERE cd_protocolo="&cd_protocolo&" ORDER BY a010_quantidade desc"
					Set rs_mat = dbconn.execute(strsql_mat)
					'	while not rs_mat.EOF
					'		qtd_material = rs_mat("a010_quantidade")
					'		nm_material = rs_mat("a061_nommat")
					linha_pula = 0%>
					<td align="left" colspan="2">
						<%while not rs_mat.EOF
							qtd_material = rs_mat("a010_quantidade")
							nm_material = rs_mat("a061_nommat")%>
								<%if linha_pula > 0 then
								total_emprestimos = total_emprestimos + qtd_material
								conta_linha_e = conta_linha_e + 1%></td>
															<tr bgcolor="<%=corlinha%>" valign="top" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
																<!--td> &nbsp; <%=conta_linha_e%>/<%=qtd_linhas_pag%></td--><td colspan="7"></td><td ><%end if%>
								<%=zero(qtd_material)%> - <%=nm_material%><br>								
								<%if linha_pula > 0 then%></td></tr><%end if%>
						<%linha_pula = linha_pula + 1
						
						rs_mat.movenext
						wend%></td>				
				</tr>
				
			<%reg_num_e = reg_num_e + 1
			total_emprestimos = total_emprestimos + qtd_material
			
			mostra_registro = 0	
			
			numero = numero + 1
				if numero > 2 then
				numero = 1
				end if
	
			'******************* Mostra os valores totais ***********************
				if linha_numerar_e = emprestimos then
					'reg_num_e = reg_num_e + 4
					conta_linha_e = conta_linha_e + 3%>
					<tr style="height:20px;"><td colspan="8" valign="bottom"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr bgcolor="dadada" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#dadada');" style="font-size:8px; height:20px;" valign="middle">
						<td colspan="4"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>TOTAL DE PROCEDIMENTOS DE EMPRESTIMO:</b> <%=rack_num_e%><%=emprestimos%></td>
						<td align="right"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>TOTAL DE ITENS EMPRESTADOS:</b> <%=total_emprestimos%></td>
						<td align="right"><%'=parcial_emprestimos%></td>
						<td colspan="2" ></td></tr>
					<tr style="height:20px;"><td colspan="8" valign="top"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%end if%>				
			<%'***********************************************************************
	linha_numerar_e = linha_numerar_e + 1
	rack_numerar_e = rack_numerar_e + 1
	conta_linha_e = conta_linha_e + 1%>	
	
	<%'*** RODAPÉ - PULO DE PÁGINAS
		if conta_linha_e = qtd_linhas_pag then
		reg_num_e = 0
		conta_linha_e = 0
		num_pag = num_pag + 1%>
			<tr><td colspan="8">&nbsp;</td></tr>
			<tr id="print_ok"><td colspan="8"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
			<tr style="font-size:10px;" >
				<td colspan="2">&nbsp;</td>
				<td colspan="5" style="page-break-after:always;" align="center">Data de emissão: <b><%=zero(day(now))&"/"&mesdoano(month(now))&"/"&right(year(now),2)&" "&hour(now)&":"&minute(now)%></b></td>
				<td align="right">&nbsp;Pág.:<%=num_pag%>&nbsp;&nbsp;&nbsp;</td>
			</tr>
		<%end if%>
		
	<%medico_anterior = a055_codmed
	rs.movenext
	next%>
		<%'******************* Linhas em branco  ***********************
		'response.write(reg_num_e&"***"&qtd_linhas_pag_)
			while conta_linha_e <= qtd_linhas_pag%>
				<tr bgcolor="ffffff" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'ffffff');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<%'=conta_linha_e&"***"&qtd_linhas_pag%><%'=reg_num_e&"__ "&qtd_linhas_pag&"=="&emprestimos &"__ "& linha_numerar_e%></td></tr>
				<%conta_linha_e = conta_linha_e + 1'reg_num_e = reg_num_e +1
			wend
		'***********************************************************************%>
		

		<%num_pag = num_pag + 1
		
		'end if%>
			
			
			
<%end if%>
		
<!--*******************************************************************************************************************************************************
****************  E.0  Fim do relatório  - Último rodapé   ************************************************************************************************
*********************************************************************************************************************************************************-->

	<%if nao_mostra_ultimo_rodape = "" Then%>
		<tr id="print_ok"><td colspan="8"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<tr style="font-size:10px;">
			<td colspan="2" align="right">&nbsp;</td>
			<td colspan="5" align="center">Data de emissão: <b><%=zero(day(now))&"/"&mesdoano(month(now))&"/"&right(year(now),2)&" "&hour(now)&":"&minute(now)%></b></td>
			<td align="right">&nbsp;Pág.:<%=num_pag%>&nbsp;&nbsp;&nbsp;</td>
		</tr>
	<%end if%>
<%'**************************************************************************************************
'*************************** FIM DO  RELATÓRIO DE PROCEDIMENTOS *************************************
'****************************************************************************************************%>

</table>
<%'end if%>
