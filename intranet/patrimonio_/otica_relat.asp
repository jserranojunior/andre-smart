<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->

<style type="text/css">
<!--
.relatorio { border:1px solid black; border-collapse:collapse; border-left-color:#000000S; font: 12px verdana, arial, helvetica, sans-serif; color:#000000;}
-->
</style>
<html>
<%
cd_pat_codigo = request("cd_pat_codigo")
strcd_unidade  =  request("cd_unidade")
ordem = request("ordem")
	if ordem = "" Then
	ordem = "nm_tipo, no_patrimonio"
	end if

mensagem = request("mensagem")
%>
<head>
	<title>Patrimonio - óticas</title>
</head>

<body><br id="no_print">
<table align="center" id="no_print">
	
</table><br id="no_print">

<table align="center" id="no_print" style="border:1 solid orange;">
	<form action="patrimonio.asp" method="post">
	<input type="hidden" name="tipo" value="relat_otica">
	
	
	<tr style="border:1 solid orange; background-color:silver;">
		<td colspan="100%" align="center"><B>RELAÇÃO DE ÓTICAS DIVIDIDAS POR UNIDADE</B></td>
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
<table class="txt" border="0" width="600" align="center" cellpadding="0" cellspacing="0" bgcolor="#C0C0C0" >
	<tr id="ok_print"><td rowspan="1000%" id="ok_print"><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td></tr>
	<%'*** parâmetros iniciais para montagem dapágina***
	
	num_lista = 0
	linha = 0
	titulo = 1
	cabeca = 1
	qtd_registros_pg = 1
	conta_reg = 1
	
	linha_bloc = linha
	qtd_linhas_pg = 27
	qtd_total_pag = 37
	
	'qtd_linhas_pg_dif = qtd_linhas_pg
	
	estilo_titulo = "border:1px solid gray; border-collapse:collapse; font-size:12px; color:white;"
	estilo_corpo = "border:1px solid gray; border-collapse:collapse; font-size:12px;"
	estilo_acerto = "border:1px solid white; border-collapse:collapse; font-size:12px;"
	
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				'*** Conta quantos registros por especialidade e grava em variável ***
				strsql = "SELECT     COUNT(cd_especialidade) AS qtd_reg, cd_especialidade FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'O' GROUP BY cd_especialidade, cd_especialidade ORDER BY cd_especialidade"
				Set rs_conta = dbconn.execute(strsql)
					while not rs_conta.EOF 
						espec_qtd = espec_qtd&";"&rs_conta("qtd_reg")&","&rs_conta("cd_especialidade")
					rs_conta.movenext
					wend
				
					if not espec_qtd = "" Then
						espec_qtd = mid(espec_qtd,2,len(espec_qtd))
						espec_array = split(espec_qtd,";")
						'	
							For Each espec_item_junto In espec_array
								qtd_reg = mid(espec_item_junto,1,instr(espec_item_junto,",")-1)
								espec_reg = mid(espec_item_junto,instr(espec_item_junto,",")+1,len(espec_item_junto))
									total_qtd_reg = total_qtd_reg + 3 + int(qtd_reg)
								'response.write("<tr><td colspan=3>"&espec_qtd&" - "&espec_reg&"<br></td></tr>")
								
							'next
							
							'if abs(total_qtd_reg) <= qtd_linhas_pg then
								'response.write("Total é menor ou igual que a qtd de linhas pré-estabelecidas. Coloca tudo em uma pág. só.")
							'	uma_pagina = 1
							'else
								'response.write("Total é maior que a qtd de linhas pré-estabelecidas")
							'	uma_pagina = 2
							'end if
					'end if
				'*************************************************************************
				
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'O' AND cd_especialidade = "&espec_reg&" order by cd_especialidade,cd_posicao,nm_equipamento"
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
				cd_unidade = rs_patrimonio("cd_unidade")
					strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" AND cd_status <> 0"
					Set rs_uni = dbconn.execute(strsql)
					while not rs_uni.EOF
						nm_unidade = rs_uni("nm_unidade")
					rs_uni.movenext
					wend
					
				nm_sigla = rs_patrimonio("nm_sigla")
				cd_rack = rs_patrimonio("cd_rack")
					if cd_rack = "null" Then
						rack_anterior = 0
					end if
					
					
				num_hospital =  rs_patrimonio("num_hospital")
				cd_rack = rs_patrimonio("cd_rack")
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
					
						'if 
						numreg = numreg + 1
						'*********************************************
						'*** Conta os equipamentos que compõe o RACK ***
						'*********************************************
						strsql = "SELECT COUNT (no_patrimonio)as conta FROM VIEW_patrimonio_lista WHERE cd_rack = '"&cd_rack&"' and cd_unidade = "&cd_unidade&""
							 Set rs_conta = dbconn.execute(strsql)
						if not rs_conta.EOF Then
							conta_patr = rs_conta("conta")
						end if%>
								
								<%'*** salto de página ***
								if int(saltar_pg) = int(1) then
									cabeca = 1
									saltar_pg = 0%>
									<%while not linha = qtd_total_pag 'Acerta as linhas até o fim
										linha = linha + 2%>
										<tr><td colspan="100" style="<%=estilo_acerto%>">&nbsp;<%'=linha%></td></tr>
									<%wend%>
									<tr>
										<td colspan="100%" style="page-break-after:always;" align="right"><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%><%'=linha&"-"&qtd_linhas_pg%><%'=linhas_dif%></td>
									</tr>									
									<%linha = 0
								end if%>
								
								<%if cabeca = 1 then
								'estilo_titulo = "border:1px solid gray; border-collapse:collapse; font-size:12px; color:white;"
								'estilo_corpo = "border:1px solid gray; border-collapse:collapse; font-size:12px;"%>
									<tr><td><img src="../imagens/px.gif" alt="" width="1" height="33" border="0"></td></tr>
									<tr style="border:0px solid gray; border-collapse:collapse; font-size:14px; color:black;">			
										<td colspan="6" align="center" valign="middle">
													<b>HOSPITAL - <%=nm_unidade%> <br>
													Relação de Óticas  - Videocirurgia<br>
													Vd Lap <%'=linhas_dif%><%'=espec_qtd%></td>
										<td colspan="2" align="right" valign="middle"><img src="../imagens/logotipo_vdlap.gif" alt="" width="80" height="40" border="0"></td>
									</tr>
								<%end if
								
								if titulo = 1 or cabeca = 1 then%>
									<%linha = linha +1%>
								<tr><td style="<%=estilo_acerto%>">&nbsp;<%'linha=linha-1%><%'=linha%></td></tr>
								<tr bgcolor="#969696">
									<!--td>&nbsp;<%'linha=linha-1%><%'linhas_parc = linhas_parc + 1%><%'=linha%><%'=linhas_dif%></td-->
									<td align="center" style="<%=estilo_titulo%>"><b>&nbsp;</b><%'=linha%></td>
									<td align="center" style="<%=estilo_titulo%>"><b>Qtd</b><%'=linha%></td>
									<td align="center" style="<%=estilo_titulo%>"><b>&nbsp;<%=nm_especialidade%>&nbsp;</b><%'=conta_patr%></td>
									<td align="center" style="<%=estilo_titulo%>"><b>Modelo</b></td>
									<td align="center" style="<%=estilo_titulo%>"><b>Marca</b></td>
									<td align="center" style="<%=estilo_titulo%>"><b>N/S</b></td>
									<td align="center" style="<%=estilo_titulo%>"><b>Nº Patrimonio</b></td>
									<td align="center" style="<%=estilo_titulo%>"><b>Nº Hospital</b></td>
								</tr>
								<tr>
									<%linha = linha +1%>
									<td style="<%=estilo_titulo%>">&nbsp;<%'linha=linha-1%><%'=linha%></td>
									<td style="<%=estilo_titulo%>">&nbsp;</td>
									<td style="<%=estilo_titulo%>">&nbsp;</td>
									<td style="<%=estilo_titulo%>">&nbsp;</td>
									<td style="<%=estilo_titulo%>">&nbsp;</td>
									<td style="<%=estilo_titulo%>">&nbsp;</td>
									<td align="center" style="<%=estilo_titulo%>" bgcolor="#969696"><b>VD LAP</b></td>
									<td align="center" style="<%=estilo_titulo%>" bgcolor="#969696"><b>CC</b></td>
								</tr>
								<%end if%>
								<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');" style="border:2px solid black;">									
									<!--td style="<%=estilo_corpo%>"--><%'linha=linha+1%><%linha_bloc=linha_bloc+1%> <%'=linha&"."&linhas_dif&"*"&linha_bloc%><!--/td-->
									<td valign="top" align="center" valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=conta_reg%></td>
									<td valign="top" align="center" valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;1</td>
									<td valign="top" align="left" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<a href="javascript:viod(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar','Alterações','width=700','height=600'); return false;"><%=nm_equipamento%></a> <%'="&nbsp;"&cd_posicao%></td>
									<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=cd_pn%></td>
									<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=nm_marca%></td>
									<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=cd_ns%></td>
									<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar','Alterações','width=700','height=600'); return false;">&nbsp;<%="0"&no_patrimonio%></td>
									<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=num_hospital%><%'=nm_especialidade%><%'=qtd_registros_pg& "-"&titulo&"-"&numreg&"-"&qtd_reg%> </td>			
								</tr>
								<%linha = linha + 1
								
								cabeca = 0
								conta_reg = conta_reg + 1
								'*** Posicionamento do Título ***
								if numreg = abs(qtd_reg) then
									titulo = 1
									numreg = 0
								else
									titulo = titulo + 1
								end if
									'*** 
									if titulo = abs(qtd_reg) then
										titulo = 0
										'numreg = 0
									end if
										if qtd_registros_pg = abs(qtd_linhas_pg) then
											'titulo = 0
											numreg = 0
											saltar_pg = 1
											qtd_registros_pg = 0
										end if
								'*** conta registro de bloco ***
								if int(conta_reg) > int(qtd_reg) then
									conta_reg = 1 
								end if
								
								'*** Quebra de página ***
								if int(numreg) = int(qtd_linhas_pg) then
									numreg = 0
									saltar_pg = 1
								end if
								
								qtd_registros_pg = qtd_registros_pg + 1
								qtd_linhas_pg_dif = linhas_dif
								especialidade_anterior = nm_especialidade
								
								conta_patr = ""
								num_lista = num_lista + 1
								rs_patrimonio.movenext
								loop
								%>
								<%'while not linha = qtd_linhas_pg
										'linha = linha + 1%>
										<!--tr>
											<td>&nbsp;</td>
											<td colspan="100%">&nbsp;<adequa as linhas para o rodapé (<%=linha%>)></td>
										</tr-->
									<%'wend%>
									
									<!--tr><td colspan="100%>&nbsp;salta a linha</td></tr>
									<tr>
										<td colspan="100%" align="right"><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%><%'=linha&"-"&qtd_linhas_pg%><%'=linhas_dif%></td>
									</tr-->
								<!--tr><td><br>&nbsp;</td></tr>
								<tr>
									<td colspan="6">&nbsp;</td>
									<td align="right"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
								</tr-->
								<%
								next
								end if%>
								<%while not linha = qtd_total_pag 'Acerta as linhas até o fim
										linha = linha + 1%>
										<tr><td colspan="100" style="<%=estilo_acerto%>">&nbsp;<%'=linha%></td></tr>
									<%wend%>
	<tr>
										<td colspan="100%" align="right"><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%><%'=linha&"-"&qtd_linhas_pg%><%'=linhas_dif%></td>
									</tr><tr>
		<!--td id="ok_print"><img src="../imagens/blackdot.gif" alt="" width="90" height="3" border="0"></td-->	
		<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="215" height="2" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="150" height="3" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="90" height="4" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="105" height="5" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="100" height="6" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="100" height="7" border="0"></td>
	</tr><tr><td colspan="100%">&nbsp;</td></tr>


</table>
<%end if%>

</body>
</html>
