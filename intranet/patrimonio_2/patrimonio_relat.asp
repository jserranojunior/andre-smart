<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->

<%
cd_pat_codigo = request("cd_pat_codigo")
strcd_unidade  =  request("cd_unidade")
ordem = request("ordem")
	if ordem = "" Then
	ordem = "nm_tipo, no_patrimonio"
	end if

mensagem = request("mensagem")%>
<br id="no_print">	
<br id="no_print">
<table align="center" id="no_print" style="border:1 solid orange; page-break-after:always;">
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">patrimonio_relat.asp</span></td></tr>
	<%end if%>
	<form action="patrimonio.asp" method="post">
	<input type="hidden" name="tipo" value="relat_int">
	<tr style="border:1 solid orange; background-color:silver;">
		<td colspan="100%" align="center"><B>RELAÇÃO DE PATRIMÔNIOS DIVIDIDOS POR UNIDADE E RACK</B></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp; Unidade:</td>
			<%strsql = "SELECT * FROM TBL_unidades where cd_status <> 0"
			Set rs_uni = dbconn.execute(strsql)%>
	    <td><select name="cd_unidade" class="inputs" onFocus="nextfield ='num_hospital';">
						<option value=""></option>
						<%Do While Not rs_uni.eof
						if int(strcd_unidade) = rs_uni("cd_codigo") then%><%unidade_check = "selected"%><%end if%>
						<option value="<%=rs_uni("cd_codigo")%>" <%=unidade_check%>><%=rs_uni("nm_unidade")%></option>
						<%rs_uni.movenext
						unidade_check = ""
						loop%>
					</select>					
		</td>
	<tr>		
		<td>&nbsp;</td>
		<td><input type="submit" name="ok" value="ok"></td>
		<td><img src="../imagens/ic_print.gif" alt="" width="24" height="24" border="0" onClick="window.print();">&nbsp;&nbsp;<img src="../imagens/ic_print_view.gif" alt="" width="24" height="26" border="0" onclick="visualizarImpressao();">
		<i style="color:red;"><b>ATENÇÃO:</b> Ao impimir, selecione A4 e orientação paisagem</i></td>	
	</tr>
		</form>
	</tr>
</table><br id="no_print">
<%if strcd_unidade <> "" Then%>
<table class="txt" border="0" width="600" align="center" cellpadding="0" cellspacing="0" bgcolor="#C0C0C0" style="border-collapse:collapse;">
	<%num_lista = 1
	linha = 0
	linha_bloc = linha
	qtd_linhas_pg = 27
	qtd_linhas_pg_dif = qtd_linhas_pg
	titulo = ""
	cabeca = 1
	
	estilo_titulo = "border:1px solid gray; border-collapse:collapse; font-size:12px; color:white;"
	estilo_corpo = "border:1px solid gray; border-collapse:collapse; font-size:12px;"
	
	'rack_anterior = 0
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				'strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' order by cd_rack, cd_posicao, nm_equipamento"
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade_comodato = "&strcd_unidade&" AND nm_tipo = 'E' order by cd_rack_comodato, cd_posicao, nm_equipamento_novo"
			  	Set rs_patrimonio = dbconn.execute(strsql)
				
				Do while not rs_patrimonio.EOF
				cd_pat_codigo = rs_patrimonio("cd_codigo")
				no_patrimonio = rs_patrimonio("no_patrimonio")
				cd_equipamento = rs_patrimonio("cd_equipamento")
				nm_equipamento = rs_patrimonio("nm_equipamento_novo")
				
				cd_marca = rs_patrimonio("cd_marca")
				nm_marca = rs_patrimonio("nm_marca")
				cd_ns = rs_patrimonio("cd_ns")
				cd_pn = rs_patrimonio("cd_pn")
				nm_tipo = rs_patrimonio("nm_tipo")
				cd_especialidade = rs_patrimonio("cd_especialidade")
				'cd_unidade = rs_patrimonio("cd_unidade")
				cd_unidade_comodato = rs_patrimonio("cd_unidade_comodato")
					'strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" AND cd_status <> 0"
					strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade_comodato&" AND cd_status <> 0"
					Set rs_uni = dbconn.execute(strsql)
					while not rs_uni.EOF
						nm_unidade = rs_uni("nm_unidade")
					rs_uni.movenext
					wend
					
				nm_sigla = rs_patrimonio("nm_sigla")
				'cd_rack = rs_patrimonio("cd_rack")
				cd_rack = rs_patrimonio("cd_rack_comodato")
					if cd_rack = "null" Then
						rack_anterior = 0
					end if
					
				num_hospital =  rs_patrimonio("num_hospital")
				'cd_rack = rs_patrimonio("cd_rack") '????? Aparece novamente
				num_hospital = rs_patrimonio("num_hospital")
				cd_posicao = rs_patrimonio("cd_posicao")
				dt_data = rs_patrimonio("dt_data")
				
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
					strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"' "
				  	Set rs_espec = dbconn.execute(strsql_espec)
						if not rs_espec.EOF then
							nm_especialidade = rs_espec("nm_especialidade")
						end if
					
					strsql_rack ="SELECT * FROM TBL_rack Where a056_codrac='"&cd_rack&"' "
				  	Set rs_rack = dbconn.execute(strsql_rack)
						if not rs_rack.EOF then
							nm_rack = rs_rack("a056_desrac")
						end if
						
				 '*********************************************
				'*** Conta os equipamentos que compõe o RACK ***
				 '*********************************************
				 'strsql = "SELECT COUNT (no_patrimonio)as conta FROM VIEW_patrimonio_lista WHERE cd_rack = '"&cd_rack&"' and cd_unidade = "&cd_unidade&""
				 'strsql = "SELECT COUNT (no_patrimonio)as conta FROM VIEW_patrimonio_lista WHERE cd_rack = '"&cd_rack&"' and cd_unidade_comodato = "&cd_unidade_comodato&""
				 strsql = "SELECT COUNT (no_patrimonio)as conta FROM VIEW_patrimonio_lista WHERE cd_rack_comodato = '"&cd_rack&"' and cd_unidade_comodato = "&cd_unidade_comodato&""
			  	 Set rs_conta = dbconn.execute(strsql)
				 if not rs_conta.EOF Then
				 	conta_patr = rs_conta("conta")
				 end if%>
	
	<%if saltar_pg = 1 then
		cabeca = 1
		'while not linha = qtd_linhas_pg
		'	linha = linha + 1%>
			<tr><td colspan="100%">&nbsp;<!--adequa as linhas para o rodapé (<%=linha%>)</td>--></tr-->
		<%'wend%>
		<tr>
			<td colspan="7"  align="right"><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%><%'=linha&"-"&qtd_linhas_pg%><%'=linhas_dif%></td>
		</tr>
		<tr><td style="page-break-after:always;">&nbsp;</td></tr>		
	<%linha = 0
	end if%>
	<%if cabeca = 1 then
	estilo_titulo = "border:1px solid gray; border-collapse:collapse; font-size:12px; color:white;"
	estilo_corpo = "border:1px solid gray; border-collapse:collapse; font-size:12px;"%>
		<tr style="border:0px solid gray; border-collapse:collapse; font-size:14px; color:black;">			
			<td colspan="5" align="center" valign="middle">
						<b>HOSPITAL - <%=nm_unidade%> <br>
						Relação de Equipamentos  - Videocirurgia<br>
						Vd Lap <%'=linhas_dif%></td>
			<td colspan="2" align="right" valign="middle"><img src="../imagens/logotipo_vdlap.gif" alt="" width="80" height="40" border="0"></td>
		</tr>
	<%end if
	
	if rack_anterior = "" OR rack_anterior <> cd_rack  Then%>				
		<%linhas_dif = qtd_linhas_pg_dif - conta_patr
		linha_bloc = 0%>
	<tr><td>&nbsp;<%'linha=linha-1%><%'=linha%><br>&nbsp;<%'linha=linha-1%><%'=linha%></td></tr>
	<tr bgcolor="#969696">
		<!--td>&nbsp;<%'linha=linha-1%><%'linhas_parc = linhas_parc + 1%><%'=linha%><%'=linhas_dif%></td-->
		<td align="center" style="<%=estilo_titulo%>"><b>Qtd</b><%'=linha%></td>
		<td align="center" style="<%=estilo_titulo%>"><b>&nbsp;<%=nm_rack%>&nbsp;</b><%'=conta_patr%></td>
		<td align="center" style="<%=estilo_titulo%>"><b>Modelo</b></td>
		<td align="center" style="<%=estilo_titulo%>"><b>Marca</b></td>
		<td align="center" style="<%=estilo_titulo%>"><b>N/S</b></td>
		<td align="center" style="<%=estilo_titulo%>"><b>Nº Patrimonio</b></td>
		<!--td align="center" style="<%'=estilo_titulo%>"><b>Nº Hospital</b></td-->
		<td align="center" style="<%=estilo_titulo%>"><b>Nº Hospital</b></td>
	</tr>
	<tr>
		<td style="<%=estilo_titulo%>">&nbsp;<%'linha=linha-1%><%'=linha%></td>
		<td style="<%=estilo_titulo%>">&nbsp;</td>
		<td style="<%=estilo_titulo%>">&nbsp;</td>
		<td style="<%=estilo_titulo%>">&nbsp;</td>
		<td style="<%=estilo_titulo%>">&nbsp;</td>
		<td align="center" style="<%=estilo_titulo%>" bgcolor="#969696"><b>VD LAP</b></td>
		<td align="center" style="<%=estilo_titulo%>" bgcolor="#969696"><b>CC</b></td>
	</tr>	
	<%end if%>
	
	<%data_agora = month(now)&"/"&day(now)&"/"&year(now)
		strsql = "SELECT * FROM View_patrimonio_emprestimos where cd_situacao >= 4 and cd_patrimonio = '"&cd_pat_codigo&"' and dt_emprestimo <= '"&data_agora&"' and dt_retorno is NULL OR cd_situacao >= 4 and cd_patrimonio = '"&cd_pat_codigo&"' and dt_emprestimo <= '"&data_agora&"' and dt_retorno >= '"&data_agora&"'"
		Set rs_pat_emprestimo = dbconn.execute(strsql)
		if not rs_pat_emprestimo.EOF then
			cd_emprestimo = rs_pat_emprestimo("cd_codigo")
			nm_sigla_des = rs_pat_emprestimo("nm_sigla_des")
		end if
		
		'if cd_emprestimo = "" Then%>
			<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');" style="border:2px solid black;">		
				<!--td style="<%=estilo_corpo%>"--><%'linha=linha+1%><%linha_bloc=linha_bloc+1%> <%'=linha&"."&linhas_dif&"*"&linha_bloc%><!--/td-->
				<td valign="top" align="center" valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;1</td>
				<td valign="top" align="left" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<a href="javascript:viod(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar&jan=1&caminho=patrimonio','Alterações','width=700,height=600,scrollbars=1'); return false;"><%=nm_equipamento%></a> <%'="&nbsp;"&cd_posicao%></td>
				<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=cd_pn%></td>
				<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=nm_marca%></td>
				<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=cd_ns%></td>
				<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar','Alterações','width=700','height=600'); return false;">&nbsp;<%=prefx_pat(no_patrimonio)%></td>
				<!--td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=num_hospital%><%'=rack_anterior&"+"&cd_rack%></td-->
				<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>"><%=num_hospital%>
					<%data_agora = month(now)&"/"&day(now)&"/"&year(now)
					strsql = "SELECT * FROM View_patrimonio_emprestimos where cd_situacao >= 4 and cd_patrimonio = '"&cd_pat_codigo&"' and dt_emprestimo <= '"&data_agora&"' and dt_retorno is NULL OR cd_situacao >= 4 and cd_patrimonio = '"&cd_pat_codigo&"' and dt_emprestimo <= '"&data_agora&"' and dt_retorno >= '"&data_agora&"'"
					Set rs_pat_emprestimo = dbconn.execute(strsql)
					if not rs_pat_emprestimo.EOF then
						cd_emprestimo = rs_pat_emprestimo("cd_codigo")
						nm_sigla_des = rs_pat_emprestimo("nm_sigla_des")
					end if
					
					if cd_emprestimo <> "" Then%><span style="font-style:italic; color:gray;" id="no_print"><%=nm_sigla_des%></span>
					<%end if%></td>
			</tr>
		<%'end if%>
	<%cabeca = 0
	
	'*** diz que se a sobra do limite de linhas for menor que o próximo bloco, salta a página. 
	'if qtd_linhas_pg_dif < conta_patr and linha_bloc = conta_patr then
	if qtd_linhas_pg_dif < conta_patr AND linha_bloc = conta_patr then
		'qtd_linhas_pg_dif = qtd_linhas_pg
		linhas_dif = qtd_linhas_pg
		linha_bloc = 0
		saltar_pg = 1
	else 
		saltar_pg = 0
	end if
	
	qtd_linhas_pg_dif = linhas_dif
	rack_anterior = cd_rack
	titulo = 1
	conta_patr = ""
	nm_especialidade = ""
	nm_sigla_des = ""
	cd_emprestimo = ""
	num_lista = num_lista + 1
	rs_patrimonio.movenext
	loop
	%>
	<%'while not linha = qtd_linhas_pg
			'linha = linha + 1%>
			<tr>
				<td>&nbsp;</td>
				<td colspan="100%">&nbsp;<!--adequa as linhas para o rodapé (<%=linha%>)--></td>
			</tr>
		<%'wend%>
		
		<tr><td colspan="100%>&nbsp;salta a linha</td></tr>
		<tr>
			<td colspan="100%" align="right"><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%><%'=linha&"-"&qtd_linhas_pg%><%'=linhas_dif%></td>
		</tr>
	<!--tr><td><br>&nbsp;</td></tr>
	<tr>
		<td colspan="6">&nbsp;</td>
		<td align="right"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr-->
	<tr>
		<!--td id="ok_print"><img src="../imagens/blackdot.gif" alt="" width="90" height="3" border="0"></td-->	
		<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="215" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="105" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
	</tr>
</table>
<%end if%>