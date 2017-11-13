<SCRIPT LANGUAGE="javascript">
function Jsconfirm_print()
{
  if (confirm ("Tem certeza que deseja imprimir?"))
	  {
		//document.location.href('acoes/patrimonio_acao.asp?cd_apaga='+cod1+'&cd_tipo=patrimonio');		
	  }
}
</SCRIPT>
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
		<%strsql_unidade = "SELECT * FROM TBL_unidades Where cd_codigo = '"&cd_unidade&"'"
		Set rs_unidade = dbconn.execute(strsql_unidade)%>
		<%if not rs_unidade.EOF then
			nm_unidade = rs_unidade("nm_unidade")%>	
		<%rs_unidade.movenext
		end if%>
<br id="no_print">
<br id="no_print">		
<table align="center" style="border:1px solid gray; font-size: 13px; border-collapse:collapse;" id="no_print">	
	<tr>
		<td align="center" style="border:1px solid gray; font-size: 15px; border-collapse:collapse;"><b>Relatório de Cirurgia</b></td>
	</tr>
	<tr>
		<td align="center" style="color:red;">
			<br><%=nm_unidade%>
			<br><b><%=mes_selecionado(int(dt_mes))&"/"&ano_sel%></b></td>
	</tr>
	<tr>
		<td align="center"><br><input type="button" onclick="visualizarImpressao();" value="Visualizar" id="inputs"><br>
		<br>
		<br>
		<input type="button" onClick="window.print()" value="IMPRIMIR" id="inputs"></td>
		
	</tr>
	<tr><td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td></tr>
</table>
<table align="center" style="border:0px solid gray; font-size: 13px; border-collapse:collapse;" id="oks_print">
	<%'******************************************************
	'***     Define a quantidade de linhas por página     ***
	'********************************************************
	'*** Use 41 para impressão em folha Carta e 46 para A4 ***
	
		qtd_linhas_pag = 46
	
	'*********************************
	'*** Contagem de procedimentos ***
	'*********************************
	'A.1 - Conta todos os registros normais
		strsql = "SELECT COUNT (a002_numpro) AS pre_conta, dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"'  AND cd_co <> '1' Group by dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla ORDER BY dt_ano, dt_mes"
			Set rs_conta = dbconn.execute(strsql)
			if not rs_conta.EOF Then
				pre_conta = rs_conta("pre_conta")
			end if
		
	'A.2 - Conta todos os emprestimos
		strsql = "SELECT COUNT (a002_numpro) AS pre_conta, dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage = 'E' Group by dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla ORDER BY dt_ano, dt_mes"
			Set rs_conta = dbconn.execute(strsql)
			if not rs_conta.EOF Then
				pre_conta_e = rs_conta("pre_conta")
			end if
		
	'A.3 - Conta os procedimentos realizados no CO.
		strsql = "SELECT COUNT (a002_numpro) AS pre_conta, dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND cd_co = '1' Group by dt_ano, dt_mes,a053_codung, nm_sigla,nm_sigla ORDER BY dt_ano, dt_mes"
			Set rs_conta = dbconn.execute(strsql)
			if not rs_conta.EOF Then
				pre_conta_co = rs_conta("pre_conta")
			end if
		
			'*** Fim das contagens de procedimentos	
			procedimentos = pre_conta - pre_conta_e' - pre_conta_co
			emprestimos = pre_conta_e			
			proc_co = pre_conta_co
			Total_procedimentos = procedimentos + emprestimos + proc_co
			
			'response.write("proc("&procedimentos&")<br>co("&proc_co&")<br>empr:("&emprestimos&")<br>__________<br>Total proc:"&Total_procedimentos&")")
		
				'Regra Edmundo Vasconcelos ******************
					if cd_unidade = 20 then 
						rack_ed = "a056_codrac, a056_desrac,"
					end if
				'*********************************************

'******************************************************************************************************
	'B.1 - Mostra os procedimentos - Normais
'******************************************************************************************************
	
	strsql = "SELECT a002_numpro, a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, "&rack_ed&" a002_tipage FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage <> 'E' AND cd_co <> '1' Group by a002_numpro,a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, "&rack_ed&" a002_tipage ORDER BY "&rack_ed&" a002_datpro"
	Set rs = dbconn.execute(strsql)%>
	
	<%linha_numerar =  1
	rack_numerar = 1
	rack_total = 1
	num_pag = 1
	
	for i = 1 to procedimentos
	
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
					
						strsql_conta_rack = "SELECT COUNT(A002_numpro) AS conta, A056_codrac FROM View_Protocolo_lista WHERE A002_datpro BETWEEN '"&dt_mes&"/01/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND A053_codung = "&cd_unidade&" AND a056_codrac = "&a056_codrac&"  AND cd_co <> '1' GROUP BY A056_codrac"
						Set rs_conta_rack = dbconn.execute(strsql_conta_rack)
						
						if not rs_conta_rack.EOF Then
							rack_num = rs_conta_rack("conta")
						end if
				end if
			
			rs.movenext
			
		end if%>
	
<!-- ***************************************************
***													 ***
*** INICIO DO RELATÓRIO DE CIRURGIAS E PROCEDIMENTOS ***
***													 *** 
******************************************************-->	
		<!-- mostrava o cabeçalho -->	
		<!-- mostrava o rack para o edmundo -->	
	<%if reg_num < 1 then%>
		<tr>
			<td colspan="3" align="left">&nbsp;<img src="../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></td>
			<td colspan="2" align="center"><b>Vd Lap Cirúrgica LTDA</b></td>
		</tr>
		<tr>		
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
				<b>RELATÓRIO DE PROCEDIMENTOS MÉDICOS</b> <br>
				<b><%=nm_unidade%></b><br>
				<b><%=mes_selecionado(int(dt_mes))%>/<%=ano_sel%></b></td>
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<%if cd_unidade = 20 then
		reg_num = reg_num + 1
		numero = 1%>
			<tr><td align="center" colspan="7">Procedimentos Médicos Realizados com <%=a056_desrac%> <%'=rack_num%></td></tr>
			<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<%end if%>
		<tr style="border:0px; solid black; font-size:11px;">
			<td align="right">&nbsp;</td>
			<td align="left"><b>Data</b></td>		
			<td align="left"><b>Paciente</b></td>
			<td align="right"><b>Reg</b>&nbsp;</td>
			<td align="left">&nbsp;<b>Cirurgia</b></td>			
			<td align="left" colspan="2"><b>Convênio</b></td>
			<td align="left"><b>Equipe Médica</b></td>		
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>	
	<%end if%>
		<!-- mostrava o titulo das colunas -->		
		
			<%'Colore as linhas do relatório
				if numero = 1 then
					corlinha = "#dadada"
				else
					corlinha = "#ffffff"
				end if
			'*******************************%>
		
		<%'CORPO DO RELATÓRIO - %>
		<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
			<%reg_num = reg_num + 1%>
			<td align="left">&nbsp;<%=linha_numerar%></td>
			<!--td align="left"><%=reg_num%></td>
			<td><%=rack_numerar%></td>
			<td><%=rack_total%></td-->
			<td align="left"><%=dt_data%></td>
			<td align="left"><%=a002_pacnom%></td>
			<td align="right"><%=a002_pacreg%>&nbsp;</td>
			<td align="left">&nbsp;
				<%'if x = "bc"  then
				strsql_proc = "SELECT * FROM View_protocolo_procedimento WHERE a002_numpro = '"&a002_numpro&"' AND a053_codung = "&cd_unidade&" ORDER BY a003_propri desc"
				Set rs_proc = dbconn.execute(strsql_proc)%>
					<%if not rs_proc.EOF then
						a002_numpro = rs_proc("a002_numpro")
						nm_procedimento = rs_proc("nm_procedimento")
						nm_proc_len = len(nm_procedimento)
							if nm_proc_len > 35 then
							nm_procedimento = left(nm_procedimento,35)&"..."
							end if
							a003_propri = rs_proc("a003_propri")
							if a003_propri <> "S" then%><!--i--><font color="red"><%end if%><b><%=a058_codpro%></b></font><%'=lcase(a003_propri)%><%if a003_propri = "s" then%><!--/i--><%end if%><%=nm_procedimento%><br>							
					<%end if
							rs_proc.close
							set rs_proc = nothing%></td>
			<td align="left" colspan="2"><%=nm_convenio%></td>
			<td align="left"><%=a055_desmed%></td>
		</tr>
		
		
		<%numero = numero + 1
				if numero > 2 then
				numero = 1
				end if
		
		if reg_num = qtd_linhas_pag OR rack_numerar = rack_num OR procedimentos = reg_num OR procedimentos = linha_numerar then 'Salto de página caso alcance 46 registros ou se o alcançar a quantidade de registros por rack
			
			'******************* Mostra os valores parciais *********************
				if rack_numerar = rack_num OR procedimentos = linha_numerar then
					reg_num = reg_num + 3
					total_racks = total_racks&";"&a056_desrac&": "&rack_numerar
						if left(total_racks,1) = ";" Then
							rack_len = len(total_racks)
								total_racks = mid(total_racks,2,rack_len)
						end if%>
					<tr style="height:20px;"><td colspan="8" valign="bottom"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<%if proc_co = "" and cd_unidade <> 20 Then%>
					<tr bgcolor="dadada" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#dadada');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>TOTAL DE PROCEDIMENTOS REALIZADOS:</b> <%if cd_unidade = 20 then%><%=rack_num%><%else%><%=procedimentos%><%end if%></td></tr>
					<%Else%>
					<tr bgcolor="ffffff" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#ffffff');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>Total parcial de procedimentos:</b> <%if cd_unidade = 20 then%><%=rack_num%><%else%><%=procedimentos%><%end if%></td></tr>
					<%end if%>
					<tr style="height:20px;"><td colspan="8" valign="top"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%end if%>
				<%'******************* Mostra o valor total de racks (HEV) *********************
				if cd_unidade = 20 AND rack_numerar = rack_num AND int(procedimentos) = int(rack_total) then
					reg_num = reg_num + 5%>										
					<tr style="height:20px;"><td colspan="8" valign="bottom"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<%rack_array = split(total_racks,";")
					For Each rack_item In rack_array
						reg_num = reg_num + 1%>
					<tr bgcolor="#dadada" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#dadada');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>Procedimentos realizados com <%=rack_item%></b></td></tr>
					<%next%>
					<tr style="height:20px;"><td colspan="8" valign="top"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr style="height:20px;"><td colspan="8" valign="bottom"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr bgcolor="#ffffff" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#ffffff');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>TOTAL DE PROCEDIMENTOS REALIZADOS: <%=procedimentos%></b></td></tr>
					<tr style="height:20px;"><td colspan="8" valign="top"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%end if%>
				<%'******************* Mostra o valor total *********************
				if proc_co = "" AND emprestimos = "" AND procedimentos = linha_numerar then
					reg_num = reg_num + 3%>
					<tr style="height:20px;"><td colspan="8" valign="bottom"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr bgcolor="#dadada" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#dadada');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>TOTAL DE PROCEDIMENTOS REALIZADOS: <%=procedimentos%></b></td></tr>
					<tr style="height:20px;"><td colspan="8" valign="top"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%end if%>
			<%while reg_num+1 <= qtd_linhas_pag%>
				<tr style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<%'=reg_num&"__ "&qtd_linhas_pag%></td></tr>
				<%reg_num = reg_num +1 
					
					
			wend
			'********************************************************************%>
			<tr><td colspan="8">&nbsp;</td></tr>
			<tr id="print_ok"><td colspan="8"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
			<tr style="font-size:10px;" >
				<td colspan="2">&nbsp;</td>
				<td colspan="5" style="page-break-after:always;" align="center">Data de emissão: <b><%=zero(day(now))&"/"&mesdoano(month(now))&"/"&right(year(now),2)&" "&hour(now)&":"&minute(now)%></b></td>
				<td align="right">&nbsp;Pág.:<%=num_pag%>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<%reg_num = 0	'Reinicia a contagem de registros por página.
				if rack_numerar = rack_num then ' Reinicia a contagem de registros por rack para comparar com a contagem de racks do DB.
					rack_numerar = 0
					linha_numerar = 0
					'nao_mostra_ultimo_rodape = 1
				end if
				
		num_pag = num_pag + 1
		nao_mostra_ultimo_rodape = 1
		end if%>
	
		<%'Numeração e mudança das cores das linhas
			linha_numerar = linha_numerar + 1
			rack_numerar = rack_numerar + 1
			rack_total = rack_total + 1
		'********************************************%>
	<%next
	
	total_procedimentos = procedimentos
	total_procs = "total parcial de procedimentos: "&procedimentos%>

<%if proc_co <> "" Then%>
<%'******************************************************************************************************
'********************************************************************************************************
	'B.2 - Mostra os procedimentos - CO
'********************************************************************************************************
'********************************************************************************************************
	
	strsql = "SELECT a002_numpro, a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, "&rack_ed&" a002_tipage = 'E' FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage <> 'E'  AND cd_co = '1' Group by a002_numpro,a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, "&rack_ed&" a002_tipage ORDER BY "&rack_ed&" a002_datpro"
	Set rs = dbconn.execute(strsql)%>
	
	<%linha_numerar_co =  1
	rack_numerar_co = 1
	reg_num_co = 0
	
	
	for i_co = 1 to proc_co
	
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
					
						strsql_conta_rack = "SELECT COUNT(A002_numpro) AS conta, A056_codrac FROM View_Protocolo_lista WHERE A002_datpro BETWEEN '"&dt_mes&"/01/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND A053_codung = "&cd_unidade&" AND a056_codrac = "&a056_codrac&" GROUP BY A056_codrac"
						Set rs_conta_rack = dbconn.execute(strsql_conta_rack)
						
						if not rs_conta_rack.EOF Then
							rack_num = rs_conta_rack("conta")
						end if
				end if%>
			
			<!-- **********************************************************
			***													 		***
			*** INICIO DO RELATÓRIO DE CIRURGIAS E PROCEDIMENTOS - CO	***
			***													 		***
			************************************************************-->
					<!-- mostrava o cabeçalho -->	
					<!-- mostrava o rack para o edmundo -->		
					<!-- mostrava o titulo das colunas -->
				
	<%if reg_num_co < 1 then%>
		<tr>
			<td colspan="3" align="left">&nbsp;<img src="../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></td>
			<td colspan="2" align="center"><b>Vd Lap Cirúrgica LTDA</b></td>
		</tr>
		<tr>		
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
				<b>RELATÓRIO DE PROCEDIMENTOS REALIZADOS NO C.O.</b> <br>
				<b><%=nm_unidade%></b><br>
				<b><%=mes_selecionado(int(dt_mes))%>/<%=ano_sel%></b></td>
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<%'if cd_unidade = 20 then
		if cd_unidade = "x" then%>
		<tr><td align="center" colspan="7"><%=a056_desrac%> - <%=rack_num%></td></tr>
		<%end if%>
		<tr style="border:0px; solid black; font-size:11px;">
			<td align="right">&nbsp;</td>
			<td align="left"><b>Data</b></td>		
			<td align="left"><b>Paciente</b></td>
			<td align="right"><b>Reg</b>&nbsp;</td>
			<td align="left">&nbsp;<b>Cirurgia</b></td>			
			<td align="left" colspan="2"><b>Convênio</b></td>
			<td align="left"><b>Equipe Médica</b></td>		
		</tr>
		<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<%end if%>
					<%'Colore as linhas do relatório
				if numero = 1 then
					corlinha = "#ffffff"
				else
					corlinha = "#dadada"
				end if
			'*******************************%>
		
		<%'CORPO DO RELATÓRIO - CO%>
		<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
			<%reg_num_co = reg_num_co + 1%>
			<td align="left">&nbsp;<%=linha_numerar_co%></td>
			<!--td align="left"><%=reg_num_co%></td>
			<td><%=rack_numerar_co%></td-->
			<td align="left"><%=dt_data%></td>
			<td align="left"><%=a002_pacnom%></td>
			<td align="right"><%=a002_pacreg%>&nbsp;</td>
			<td align="left">&nbsp;
				<%'if x = "bc"  then
				strsql_proc = "SELECT * FROM View_protocolo_procedimento WHERE a002_numpro = '"&a002_numpro&"' AND a053_codung = "&cd_unidade&" ORDER BY a003_propri desc"
				Set rs_proc = dbconn.execute(strsql_proc)%>
					<%if not rs_proc.EOF then
						a002_numpro = rs_proc("a002_numpro")
						nm_procedimento = rs_proc("nm_procedimento")
						nm_proc_len = len(nm_procedimento)
							if nm_proc_len > 35 then
							nm_procedimento = left(nm_procedimento,35)&"..."
							end if
							a003_propri = rs_proc("a003_propri")
							if a003_propri <> "S" then%><!--i--><font color="red"><%end if%><b><%=a058_codpro%></b></font><%'=lcase(a003_propri)%><%if a003_propri = "s" then%><!--/i--><%end if%><%=nm_procedimento%><br>							
					<%end if
					
							rs_proc.close
							set rs_proc = nothing
					'end if%></td>
			<%strsql_mat = "SELECT * FROM View_protocolo_material WHERE a002_numseq = '"&a002_numpro&"' AND a053_codung = "&cd_unidade&" ORDER BY a010_quantidade desc"
			Set rs_mat = dbconn.execute(strsql_mat)%>	
			<td align="left" colspan="2"><%=nm_convenio%></td>			
			<td align="left"><%=a055_desmed%></td>
		</tr>
		<%numero = numero + 1
				if numero > 2 then
				numero = 1
				end if
		
		if reg_num_co = qtd_linhas_pag OR rack_numerar_co = rack_num_co OR proc_co = reg_num_co then 'Salto de página caso alcance 46 registros ou se o alcançar a quantidade de registros por rack
			
			'******************* Mostra os valores parciais ***********************
				if rack_numerar_co = rack_num_co OR proc_co = linha_numerar_co then
					reg_num_co = reg_num_co + 3%>
					<tr style="height:20px;"><td colspan="8" valign="bottom"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr bgcolor="ffffff" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#ffffff');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>Total parcial de procedimentos realizados no CO:</b>  <%=proc_co%></td></tr>
					<tr style="height:20px;"><td colspan="8" valign="top"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%end if%>
				<%if proc_co = linha_numerar_co then
				total_procedimentos = procedimentos + proc_co
					reg_num_co = reg_num_co + 3%>
					<tr style="height:20px;"><td colspan="8" valign="bottom"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr bgcolor="#dadada" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#dadada');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>Total parcial de procedimentos: <%=procedimentos%></b></td></tr>
					<tr bgcolor="#dadada" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#dadada');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>Total parcial de procedimentos no CO: <%=proc_co%></b></td></tr>
					<tr style="height:20px;"><td colspan="8" valign="bottom"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr bgcolor="#dadada" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#dadada');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>TOTAL DE PROCEDIMENTOS REALIZADOS: <%=procedimentos + proc_co%></b></td></tr>
					<tr style="height:20px;"><td colspan="8" valign="top"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%end if%>
			<%while reg_num_co+1 <= qtd_linhas_pag%>
				<tr style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<%'=reg_num&"__ "&qtd_linhas_pag%></td></tr>
				<%reg_num_co = reg_num_co +1 
					
					
			wend
			'***********************************************************************%>
			
			<%reg_num_co = 0	'Reinicia a contagem de registros por página.
				if rack_numerar_co = rack_num then ' Reinicia a contagem de registros por rack para comparar com a contagem de racks do DB.
					rack_numerar_co = 0
					linha_numerar_co = 0
					nao_mostra_ultimo_rodape = 1
				end if%>
			<tr><td colspan="8">&nbsp;</td></tr>
			<tr id="print_ok"><td colspan="8"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
			<tr style="font-size:10px;" >
				<td colspan="2">&nbsp;</td>
				<td colspan="5" style="page-break-after:always;" align="center">Data de emissão: <b><%=zero(day(now))&"/"&mesdoano(month(now))&"/"&right(year(now),2)&" "&hour(now)&":"&minute(now)%></b></td>
				<td align="right">&nbsp;Pág.:<%=num_pag%>&nbsp;&nbsp;&nbsp;</td>
			</tr>
		<%num_pag = num_pag + 1
		
		end if%>
		<%'Numeração e mudança das cores das linhas
			linha_numerar_co = linha_numerar_co + 1
			rack_numerar_co = rack_numerar_co + 1
				
			
			rs.movenext
			
		end if
	next%>	
		
		
		
		
	
<%end if%>

<%if emprestimos <> "" Then%>
<%'******************************************************************************************************
'********************************************************************************************************
	'B.3 - Mostra os procedimentos - Empréstimos
'********************************************************************************************************
'********************************************************************************************************
	
	strsql = "SELECT a002_numpro, a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, "&rack_ed&" a002_tipage = 'E' FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&dt_mes&"/1/"&ano_sel&"' AND '"&dt_mes&"/"&UltimoDiaMes(dt_mes,ano_sel)&"/"&ano_sel&"' AND a053_codung = '"&cd_unidade&"' AND a002_tipage = 'E' Group by a002_numpro,a002_datpro, a002_pacreg, a002_pacnom, a054_codcon, nm_convenio, a055_desmed, "&rack_ed&" a002_tipage ORDER BY "&rack_ed&" a002_datpro"
	Set rs = dbconn.execute(strsql)
	
	linha_numerar_e =  1
	rack_numerar_e = 1
	reg_num_e = 0
	
	for i_e = 1 to emprestimos
	
		if numero = 1 then
			corlinha = "#ffffff"
		else
			corlinha = "#dadada"
		end if
	
		a002_numpro = rs("a002_numpro")
		dt_data = zero(day(rs("a002_datpro")))&"/"&zero(month(rs("a002_datpro")))&"/"&right(year(rs("a002_datpro")),2)
		a002_pacnom = rs("a002_pacnom")
			nm_pac_len = len(a002_pacnom)
				if nm_pac_len > 34 then
					a002_pacnom = left(a002_pacnom,34)&"..."
				end if
		a055_desmed = rs("a055_desmed")%>
	
		<%mostra_registro = 0
		strsql_mat = "SELECT * FROM View_protocolo_material WHERE a002_numseq = '"&a002_numpro&"' AND a053_codung = "&cd_unidade&" ORDER BY a010_quantidade desc"
		Set rs_mat = dbconn.execute(strsql_mat)
			while not rs_mat.EOF
				qtd_material = rs_mat("a010_quantidade")
				nm_material = rs_mat("a061_nommat")%>		
		
			<%if reg_num_e < 1 then%>
			<tr>
				<td colspan="3" align="left">&nbsp;<img src="../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></td>
				<td colspan="2" align="center"><b>Vd Lap Cirúrgica LTDA</b></td>
			</tr>
			<tr>		
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
					<b><%=nm_unidade%></b><br>
					<b><%=mes_selecionado(int(dt_mes))%>/<%=ano_sel%></b></td>
			</tr>
			<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%if cd_unidade = 20 then%>
				<tr><td align="center" colspan="7"><%=a056_desrac%></td></tr>
				<%end if%>
			<tr style="border:0px; solid black; font-size:11px;">
				<td align="right">&nbsp;</td>
				<td align="left"><b>Data</b></td>		
				<td align="left"><b>Paciente</b></td>
				<td align="right"><b>Reg</b>&nbsp;</td>
				<td align="left">&nbsp;<b>Cirurgia</b></td>			
				<td align="left"><b>Qtd.</b></td>
				<td align="left"><b>ITENS</b></td>
				<td align="left"><b>Equipe Médica</b></td>		
			</tr>
			<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
			<%reg_num_e = reg_num_e + 1
			end if%>
			
			<%if mostra_registro = 0 then%>
				<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
					<td align="right"><%=linha_numerar_e%></td>
					<!--td align="center"><%=reg_num_e%></td>
					<!td><%'=rack_numerar_e%></td-->
					<td>&nbsp;<%=dt_data%></td>
					<td>&nbsp;<%=a002_pacnom%></td>
					<td align="right"><%=a002_numpro%></td>
					<%strsql_proc = "SELECT * FROM View_protocolo_procedimento WHERE a002_numpro = '"&a002_numpro&"' AND a053_codung = "&cd_unidade&" ORDER BY a003_propri desc"
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
					<td align="right"><%=qtd_material%></td>
					<td><%=nm_material%></td>
					<td>&nbsp;<%=a055_desmed%></td>
				</tr>
			<%else%>			
				<tr bgcolor="<%=corlinha%>" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'<%=corlinha%>');" style="font-size:8px; height:20px;" valign="middle">
					<td align="right">&nbsp;<%'=linha_numerar_e%></td>
					<!--td align="center"><%'=reg_num_e%></td>
					<!--td>&nbsp;<%'=rack_numerar_e%></td-->
					<td align="right"><%'=dt_data%></td>
					<td>&nbsp;<%'=a002_pacnom%></td>
					<td align="right"><%'=a002_numpro%></td>
					<td>&nbsp;<%'=nm_procedimento%></td>
					<td align="right"><%=qtd_material%></td>
					<td><%=nm_material%></td>
					<td>&nbsp;<%'=a055_desmed%></td>
				</tr>			
			<%end if
			
			reg_num_e = reg_num_e + 1
			
			mostra_registro = mostra_registro + 1
			parcial_emprestimos = parcial_emprestimos + qtd_material
			
			rs_mat.movenext
			wend
			
			mostra_registro = 0	
			
			numero = numero + 1
				if numero > 2 then
				numero = 1
				end if
	
			'******************* Mostra os valores parciais ***********************
				if linha_numerar_e = emprestimos then
					reg_num_e = reg_num_e + 3%>
					<tr style="height:20px;"><td colspan="8" valign="bottom"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr bgcolor="dadada" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'#dadada');" style="font-size:8px; height:20px;" valign="middle">
						<td colspan="4"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>TOTAL DE PROCEDIMENTOS DE EMPRESTIMO:</b> <%=rack_num_e%><%=emprestimos%></td>
						<td align="right"><!-- salto de linha para ajuste do rodapé -->&nbsp;<b>TOTAL DE ITENS EMPRESTADOS:</b> <%=rack_num_e%></td>
						<td align="right"><%=parcial_emprestimos%></td>
						<td colspan="2" ></td></tr>
					<tr style="height:20px;"><td colspan="8" valign="top"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%end if%>				
			<%'***********************************************************************
	linha_numerar_e = linha_numerar_e + 1
	rack_numerar_e = rack_numerar_e + 1
	
	rs.movenext
	next%>
		<%'******************* Linhas em branco  ***********************
		'response.write(reg_num_e&"***"&qtd_linhas_pag_)
			while reg_num_e <= qtd_linhas_pag%>
				<tr bgcolor="ffffff" onmouseover="mOvr(this,'#ffff99');"  onclick="javascript:;" onmouseout="mOut(this,'ffffff');" style="font-size:8px; height:20px;" valign="middle"><td colspan="8"><!-- salto de linha para ajuste do rodapé -->&nbsp;<%'=reg_num_e&"__ "&qtd_linhas_pag&"=="&emprestimos &"__ "& linha_numerar_e%></td></tr>
				<%reg_num_e = reg_num_e +1
			wend
		'***********************************************************************
		
		
		
		'if reg_num_e = qtd_linhas_pag OR emprestimos = reg_num_e then 'Salto de página caso alcance 46 registros ou se o alcançar a quantidade de registros por rack
			reg_num_co = 0	'Reinicia a contagem de registros por página.
				if rack_numerar_co = rack_num then ' Reinicia a contagem de registros por rack para comparar com a contagem de racks do DB.
					rack_numerar_co = 0
					linha_numerar_co = 0
					'nao_mostra_ultimo_rodape = 1
				end if%>
			<tr><td colspan="8">&nbsp;</td></tr>
			<tr id="print_ok"><td colspan="8"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
			<tr style="font-size:10px;" >
				<td colspan="2">&nbsp;</td>
				<td colspan="5" style="page-break-after:always;" align="center">Data de emissão: <b><%=zero(day(now))&"/"&mesdoano(month(now))&"/"&right(year(now),2)&" "&hour(now)&":"&minute(now)%></b></td>
				<td align="right">&nbsp;Pág.:<%=num_pag%>&nbsp;&nbsp;&nbsp;</td>
			</tr>
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
<%end if%>
