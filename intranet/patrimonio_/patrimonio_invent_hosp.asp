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

mensagem = request("mensagem")
%>
<br id="no_print">

<table align="center" id="no_print" style="border:1 solid blue;">
	<form action="patrimonio.asp" method="post">
	<input type="hidden" name="tipo" value="invent_hosp">
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">patrimonio_invent_hosp.asp</span></td></tr>
	<%end if%>
	<tr style="border:1 solid orange; background-color:silver;">
		<td colspan="100%" align="center"><B>INVENT�RIO DE EQUIPAMENTOS GERAL</B></td>
	</tr>
	<tr>
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
						<!--a href="javascript:;" onclick="window.open('adm/unidade_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a-->
		</td>
	</tr>
	<tr>		
		<td><input type="submit" name="ok" value="ok"></td>
		<td align="center"><img src="../imagens/ic_print.gif" alt="" width="24" height="24" border="0" onClick="window.print();"> &nbsp; <img src="../imagens/ic_print_view.gif" alt="" width="24" height="26" border="0" onclick="visualizarImpressao();"> &nbsp; &nbsp; &nbsp; &nbsp; </td>		
	</tr>	
	<tr>
		<td valign="top" align="right"><b style="color:red;">Aten��o:</b></td>
		<td valign="top"><b style="color:red;">Imprimir com orienta��o PAISAGEM<br> em folha A4.</b></td>
	</tr>
		</form>
	</tr>
</table><br id="no_print">
<%if strcd_unidade <> "" Then%>

<table class="txt" width="600" align="center" style="border-collapse:collapse; border:0px solid white;">
	<%num_lista = 0
	conta_registro = 1
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
						end if%>
						
	<%if cabeca = 0 then
	estilo_titulo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:11px; color:black;"
	estilo_corpo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:11px;"%>
	<tr>
		<td>&nbsp;</td>
		<td colspan="2">
		<%if cd_unidade = 2 or cd_unidade = 3 or cd_unidade = 15 then%>
			<img src="../imagens/logotipo_sao_luiz.gif" alt="" width="100" height="30" border="0"><!-- id="no_print"-->
			<!--img src="../imagens/logotipo_sao_luiz_bw.gif" alt="" width="100" height="50" border="0" id="ok_print"--><%end if%></td>
		<td colspan="5" align="center" valign="middle"style="font-family:arial; font-size:12px; color:black;"><b>INVENT�RIO DE EQUIPAMENTOS - VDLAP</b></td>
		<td colspan="2" align="right">
			<img src="../imagens/logotipo_vdlap.gif" alt="" width="88" height="43" border="0"><!-- id="no_print"-->
			<!--img src="../imagens/logotipo_vdlap_bw.gif" alt="" width="88" height="43" border="0"><id="ok_print"--></td>			
	</tr>
	
	<tr><td colspan="100%" align="center" style="font-size:12px;"><%if cd_unidade = 20 then%><%=nm_unidade%><%end if%></td></tr>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr>
		<td align="center" valign="bottom">&nbsp;</td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Item</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Controle de<br>Manuten��o</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Descri��o</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Marca</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Modelo</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>N� S�rie</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Preventiva</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Calibra��o</b></td>
		<%if strcd_unidade = 20 then%>
			<!--td align="center" valign="bottom" style="<%'=estilo_titulo%>"><b>Seg. Eletrica</b></td-->
		<%else%>
			<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Seg. Eletrica</b></td>
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
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%="0"&no_patrimonio%></td>		
		<td valign="top" align= "left"  style="<%=estilo_corpo%>">&nbsp;<a href="javascript:viod(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar&jan=1&caminho=patrimonio','Altera��es','width=700,height=600,scrollbars=1'); return false;"><%=nm_equipamento%></a> <%'="&nbsp;- "&cd_posicao%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_marca%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_pn%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_ns%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%if cd_preventiva = 1 then response.write("Sim") Else response.write("N�o Aplic�vel")end if%>&nbsp;</td>		
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%if cd_calibracao = 1 then response.write("Sim") Else response.write("N�o Aplic�vel")end if%>&nbsp;</td>
		<%if strcd_unidade = 20 then%>
			<!--<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%'if cd_seg_elet = 1 then response.write("Sim") Else response.write("N�o Aplic�vel")end if%>&nbsp;</td>-->
		<%else%>
			<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%if cd_seg_elet = 1 then response.write("Sim") Else response.write("N�o Aplic�vel")end if%>&nbsp;</td>
		<%end if%>		
	</tr>	
		<%'end if
		%>
	<%conta_registro = conta_registro + 1
	cd_emprestimo = ""
	
	if int(conta_registro) > 40 then
		
		cabeca = 0
		conta_registro = 1
		pagina = pagina + 1%>
		
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td>&nbsp;</td>
		<td align="center" colspan="8">&nbsp;<!--P�g.:<%=zero(pagina)%>--></td>
		<td align="center"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
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
		<td colspan="5"><b style="font-family:arial; font-size:10px;">ITA20211.PQ.006 - ANEXO 2 - INVENT�RIO DE EQUIPAMENTO VDLAP</b></td>
	</tr>
	<%end if%>
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td>&nbsp;</td>
		<td align="center" colspan="8">&nbsp;<!--P�g.:<%'=zero(pagina+1)%>--><b id="no_print"><%=nm_unidade%></b></td>
		<td align="center"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>
	<tr>
		<td id="no_print"><img src="../imagens/px.gif" alt="" width="5" height="1" border="0"></td>
		<td id="ok_print"><img src="../imagens/px.gif" alt="" width="10" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="52" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="83" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="67" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>		
	</tr>
	<tr><td colspan="100%">&nbsp;</td></tr>


</table>

<%end if%>
