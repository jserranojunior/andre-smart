<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<%
cd_pat_codigo = request("cd_pat_codigo")

strcd_unidade  =  request("cd_unidade")
	if strcd_unidade = "0" then
		'strcd_unidade = ""
	end if
	
ordem = request("ordem")
	if ordem = "" Then
	ordem = "nm_tipo, no_patrimonio"
	end if

mensagem = request("mensagem")
rel_mod = request("rel_mod")
%>
<br class="no_print">

<table align="center" class="no_print" style="border:1 solid blue;">
	<form action="patrimonio.asp" method="post">
	<input type="hidden" name="tipo" value="invent_hosp">
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">patrimonio_invent_hosp.asp</span></td></tr>
	<%end if%>
	<tr style="border:1 solid orange; background-color:silver;">
		<td colspan="100%" align="center"><B>INVENTÁRIO DE EQUIPAMENTOS GERAL</B></td>
	</tr>
	<tr>
	    <td>&nbsp; Unidade:</td>
			<%strsql = "SELECT * FROM TBL_unidades where cd_status <> 0 and cd_hospital between 1 and 3 "
			Set rs_uni = dbconn.execute(strsql)%>
	    <td><select name="cd_unidade" class="inputs" onFocus="nextfield ='num_hospital';">
						<option value="0"></option>
						<%While Not rs_uni.eof
							cd_unidade = rs_uni("cd_codigo")
							nm_unidade = rs_uni("nm_unidade")%>
							<option value="<%=cd_unidade%>" <%if int(strcd_unidade) = cd_unidade then response.write("selected")%>><%=nm_unidade%></option>
							<%rs_uni.movenext
							unidade_check = ""
						wend%>
						</select>
						<!--a href="javascript:;" onclick="window.open('adm/unidade_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a-->
		</td>
	</tr>
	<td><td><input type="radio" name="rel_mod" value="1" <%if rel_mod = "" OR rel_mod = 1 then response.write("checked")%>> Mod. 1 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <input type="radio" name="rel_mod" value="2" <%if rel_mod = 2 then response.write("checked")%>> Mod. 2 &nbsp; </td></td>
	<tr>		
		<td><input type="submit" name="ok" value="ok"></td>
		<td align="center"><img src="../imagens/ic_print.gif" alt="" width="24" height="24" border="0" onClick="visualizarImpressao();"> &nbsp; <img src="../imagens/ic_print_view.gif" alt="" width="24" height="26" border="0" onclick="visualizarImpressao();"> &nbsp; &nbsp; &nbsp; &nbsp; </td>		
	</tr>	
	<tr>
		<td valign="top" align="right"><b style="color:red;">Atenção:</b></td>
		<td valign="top"><b style="color:red;">Imprimir com orientação PAISAGEM<br> em folha A4.</b></td>
	</tr>
		</form>
	</tr>
</table><br class="no_print">
<%if strcd_unidade <> "" Then%>

<table class="txt" width="600" align="center" style="border-collapse:collapse; border:0px solid white;">
	<%num_lista = 0
	conta_registro = 1
	if rel_mod = 2 AND strcd_unidade = 2 or rel_mod = 2 AND strcd_unidade = 3 or rel_mod = 2 AND strcd_unidade = 15 then
		qtd_linhas = 37
	else
		qtd_linhas = 40
	end if
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade_comodato = "&strcd_unidade&" AND nm_tipo = 'E' AND nao_constar <> 2 order by cd_posicao, nm_equipamento_novo"
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
				cd_unidade = rs_patrimonio("cd_unidade_comodato")
					strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" AND cd_status <> 0"
					Set rs_uni = dbconn.execute(strsql)
					while not rs_uni.EOF
						nm_unidade = rs_uni("nm_unidade")
					rs_uni.movenext
					wend
					
					
				nm_sigla = rs_patrimonio("nm_sigla")
				cd_rack = rs_patrimonio("cd_rack")
				num_hospital =  rs_patrimonio("num_hospital")
				cd_rack = rs_patrimonio("cd_rack")
				num_hospital = rs_patrimonio("num_hospital")
				cd_posicao = rs_patrimonio("cd_posicao")
				dt_data = rs_patrimonio("dt_data")
				
				cd_preventiva = rs_patrimonio("cd_preventiva")
				
				cd_calibracao = rs_patrimonio("cd_calibracao")
				
				cd_seg_elet = rs_patrimonio("cd_seg_elet")
				
				nao_constar = rs_patrimonio("nao_constar")
				
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
					strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"' "
				  	Set rs_espec = dbconn.execute(strsql_espec)
						if not rs_espec.EOF then
							nm_especialidade = rs_espec("nm_especialidade")
						end if
						
					'*** Verifica quais preventivas são realizadas no hospital ****
					if i = 1 then
						prevs_max  = "MAX(cd_preventiva) AS prevs"
						prevs_cond = "cd_preventiva = 1"
					elseif i = 2 then
						prevs_max  = "MAX(cd_calibracao) AS prevs"
						prevs_cond = "cd_calibracao = 1"
					elseif i = 3 then
						prevs_max  = "MAX(cd_seg_elet) AS prevs"
						prevs_cond = "cd_seg_elet = 1"
					end if
					
					strsql = "SELECT MAX(cd_preventiva) AS prevs FROM VIEW_patrimonio_lista_servicos WHERE cd_preventiva = 1 AND (cd_unidade = "&strcd_unidade&")"
					Set rs_prevs1 = dbconn.execute(strsql)
						if not rs_prevs1.EOF then
							if rs_prevs1("prevs") = 1 then
								str_prev_display = "display:block;"
							else
								str_prev_display = "display:none;"
							end if
						end if
					
					strsql = "SELECT MAX(cd_calibracao) AS calibs FROM VIEW_patrimonio_lista_servicos WHERE cd_calibracao = 1 AND (cd_unidade = "&strcd_unidade&")"
					Set rs_prevs2 = dbconn.execute(strsql)
						if not rs_prevs2.EOF then
							if rs_prevs2("calibs") = 1 then
								str_cal_display = "display:block;"
							else
								str_cal_display = "display:none;"
							end if
						end if
					
					strsql = "SELECT MAX(cd_seg_elet) AS elets FROM VIEW_patrimonio_lista_servicos WHERE cd_seg_elet = 1 AND (cd_unidade = "&strcd_unidade&")"
					Set rs_prevs3 = dbconn.execute(strsql)
						if not rs_prevs3.EOF then
							if rs_prevs3("elets") = 1 then
								str_elet_display = "display:block;"
							else
								str_elet_display = "display:none;"
							end if
						end if
					'***************************************************************
					%>
						
	<%if cabeca = 0 then
	estilo_titulo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:11px; background-color:#003366; color:#FFFFFF; "
	estilo_corpo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:11px;"%>
	
	<%if rel_mod = 2 AND cd_unidade = 2 or rel_mod = 2 AND cd_unidade = 3 or rel_mod = 2 AND cd_unidade = 15 then
		colspan_tit = "6"
		colspan_rod = "9"
	else
		colspan_tit = "5"
		colspan_rod = "8"
	end if%>
	<%if rel_mod = 2 AND cd_unidade = 2 or rel_mod = 2 AND cd_unidade = 3 or rel_mod = 2 AND cd_unidade = 15 then%>
	<tr class="ok_print"><td align="center" valign="middle"><img src="../imagens/px.gif" alt="" width="85" height="15" border="0"></td></tr>
	<%end if%>
	<tr>
		<td>&nbsp;</td>
		<td colspan="2">
		<%if cd_unidade = 2 or cd_unidade = 3 or cd_unidade = 15 then%>
			<img src="../imagens/logotipo_sao_luiz.gif" alt="" width="100" height="30" border="0"><!-- class="no_print"-->
			<!--img src="../imagens/logotipo_sao_luiz_bw.gif" alt="" width="100" height="50" border="0" class="ok_print"--><%end if%></td>
		<td colspan="<%=colspan_tit%>" align="center" valign="middle"style="font-family:arial; font-size:12px; color:black;"><b>INVENTÁRIO DE EQUIPAMENTOS MÉDICOS - VD LAP CIRURGICA</b></td>
		<td colspan="2" align="right">
		<%if cd_unidade = 2 or cd_unidade = 3 or cd_unidade = 15 then%>
			<img src="../imagens/logotipo_rede_d_or.jpg" alt="" width="88" height="35" border="0"><!-- class="no_print"-->
			<!--img src="../imagens/logotipo_vdlap_bw.gif" alt="" width="88" height="43" border="0"><class="ok_print"-->
		<%else%>
			<img src="../imagens/logotipo_vdlap.gif" alt="" width="88" height="43" border="0">
		<%end if%></td>			
	</tr>
	
	<tr><td colspan="100%" align="center" style="font-size:12px;"><%'if cd_unidade = 20 then%><%=nm_unidade%><%'end if%></td></tr>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<%if rel_mod = 2 AND cd_unidade = 2 or rel_mod = 2 AND cd_unidade = 3 or rel_mod = 2 AND cd_unidade = 15 then%>
	<tr><td colspan="8">&nbsp;</td><td colspan="2" align="center" style="<%=estilo_titulo%>"><b>Contrato</b></td><td>&nbsp;</td></tr>
	<%end if%>
	
	<tr>
		<td align="center" valign="middle"><img src="../imagens/px.gif" alt="" width="5" height="1" border="0" class="no_print"><img src="../imagens/px.gif" alt="" width="10" height="1" border="0" class="ok_print"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Item</b><br><img src="../imagens/px.gif" alt="" width="52" height="3" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Controle de<br>Manutenção</b><br><img src="../imagens/px.gif" alt="" width="83" height="1" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Descrição</b><br><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Marca</b><br><img src="../imagens/px.gif" alt="" width="67" height="1" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Modelo</b><br><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Nº Série</b><br><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
	<%if rel_mod = 2 AND cd_unidade = 2 or rel_mod = 2 AND cd_unidade = 3 or rel_mod = 2 AND cd_unidade = 15 then%>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Fornecedor</b><br><img src="../imagens/px.gif" alt="" width="85" height="1" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Sim</b><br><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Não</b><br><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%>"><b>Status</b><br><img src="../imagens/px.gif" alt="" width="85" height="1" border="0"></td>
		<td>&nbsp;</td>
	<%else%>	
		<td align="center" valign="middle" style="<%=estilo_titulo%><%=str_prev_display%>"><b>Preventiva</b><br><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%><%=str_cal_display%>"><b>Calibração</b><br><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
		<td align="center" valign="middle" style="<%=estilo_titulo%><%=str_elet_display%>"><b>Seg. Eletrica</b><br><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
	<%end if%>
	</tr>
	<%end if%>
	<%'******************************************
	'***** Esconde a linha do item emprestado ***
	'********************************************
		data_agora = month(now)&"/"&day(now)&"/"&year(now)
		strsql = "SELECT * FROM TBL_patrimonio_emprestimos where cd_situacao = 4 and cd_patrimonio = '"&cd_pat_codigo&"' and dt_emprestimo <= '"&data_agora&"' and dt_retorno is NULL OR cd_situacao = 4 and cd_patrimonio = '"&cd_pat_codigo&"' and dt_emprestimo <= '"&data_agora&"' and dt_retorno >= '"&data_agora&"'"
				Set rs_pat_emprestimo = dbconn.execute(strsql)
				while not rs_pat_emprestimo.EOF
					cd_emprestimo = rs_pat_emprestimo("cd_codigo")
				rs_pat_emprestimo.movenext
				wend
				
				
		'if cd_emprestimo = "" Then
	'********************************************
	num_lista = num_lista+1%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="center" valign="bottom" style="border:0px solid white; border-collapse:collapse; background-color:white;">&nbsp;<%if cd_emprestimo <> "" Then%><!--X--><%end if%></td>
		<td align="center" valign="top" style="<%=estilo_corpo%>">&nbsp;<%=num_lista%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=prefx_pat(no_patrimonio)%></td>		
		<td valign="top" align= "left"  style="<%=estilo_corpo%>">&nbsp;<a href="javascript:viod(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar&jan=1&caminho=patrimonio','Alterações','width=700,height=600,scrollbars=1'); return false;"><%=nm_equipamento%></a> <%'="&nbsp;- "&cd_posicao%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_marca%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_pn%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_ns%></td>
	<%if rel_mod = 2 AND cd_unidade = 2 or rel_mod = 2 AND cd_unidade = 3 or rel_mod = 2 AND cd_unidade = 15 then%>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;Vd LAp&nbsp;</td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;X&nbsp;</td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;</td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;LOCAÇÃO&nbsp;</td>	
	<%else%>
		<td valign="top" align="center" style="<%=estilo_corpo%><%=str_prev_display%>">&nbsp;<%if cd_preventiva = 1 then response.write("Sim") Else response.write("Não Aplicável")end if%>&nbsp;</td>		
		<td valign="top" align="center" style="<%=estilo_corpo%><%=str_cal_display%>">&nbsp;<%if cd_calibracao = 1 then response.write("Sim") Else response.write("Não Aplicável")end if%>&nbsp;</td>
		<td valign="top" align="center" style="<%=estilo_corpo%><%=str_elet_display%>">&nbsp;<%if cd_seg_elet = 1 then response.write("Sim") Else response.write("Não Aplicável")end if%>&nbsp;</td>
	<%end if%>
	</tr>	
		<%'end if
		%>
	<%conta_registro = conta_registro + 1
	cd_emprestimo = ""
	
	if int(conta_registro) > qtd_linhas then
		
		cabeca = 0
		conta_registro = 1
		pagina = pagina + 1%>
		
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td>&nbsp;</td>
		<td align="center" colspan="<%=colspan_rod%>">&nbsp;<!--Pág.:<%=zero(pagina)%>--></td>
		<td align="center" colspan="1"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>
	<tr><td style="page-break-after:always;">&nbsp;</td></tr>
	<%else
		cabeca = 1
	end if
	
	
	nm_especialidade = ""
	rs_patrimonio.movenext
	loop
	%>
	<%if cd_unidade = 2 or cd_unidade = 3 or cd_unidade = 15 then%>
	<tr>
		<td>&nbsp;</td>
		<td colspan="5"><b style="font-family:arial; font-size:10px;">ITA20211.PQ.006 - ANEXO 2 - INVENTÁRIO DE EQUIPAMENTOS MÉDICOS - Vd Lap</b></td>		
	</tr>
	<%end if%>
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td>&nbsp;</td>
		<td align="center" colspan="<%=colspan_rod%>">&nbsp;<!--Pág.:<%'=zero(pagina+1)%>--><b class="no_print"><%=nm_unidade%></b></td>
		<td align="center" colspan="1"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<%if rel_mod = 2 AND cd_unidade = 2 or rel_mod = 2 AND cd_unidade = 3 or rel_mod = 2 AND cd_unidade = 15 then%>
	<tr><td>&nbsp;</td><td align="center" valign="top" style="<%=estilo_corpo%>" rowspan="3">Status</td><td valign="top" style="<%=estilo_corpo%>">&nbsp;Próprio</td></tr>
	<tr><td valign="top"></td><td valign="top" style="<%=estilo_corpo%>">&nbsp;Comodato</td></tr>
	<tr><td valign="top"></td><td valign="top" style="<%=estilo_corpo%>">&nbsp;Locação</td></tr>
	<%end if%>
</table>

<%end if%>
