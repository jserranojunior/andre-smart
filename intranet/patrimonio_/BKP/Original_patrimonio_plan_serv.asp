<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_pat_codigo = request("cd_pat_codigo")

mes_sel = request("mes_sel")
	if mes_sel = "" Then
		mes_sel = year(now)
	end if
	
ano_sel = request("ano_sel")
	if ano_sel = "" Then
		ano_sel = year(now)
	end if

strcd_unidade  =  request("cd_unidade")
ordem = request("ordem")
	if ordem = "" Then
	ordem = "nm_tipo, cd_patrimonio"
	end if

mensagem = request("mensagem")
%>
<head>
	<title>Untitled</title>
</head>

<body>
<br id="no_print">
<table align="center" id="no_print" style="border:1 solid green;">
	<form action="patrimonio.asp" method="post">
	<input type="hidden" name="tipo" value="plan_serv">
	<tr style="border:1 solid orange; background-color:silver;">
		<td colspan="100%" align="center"><B>PLANEJAMENTO DE SERVIÇOS</B></td>
	</tr>
	<tr>
		<td>&nbsp; Ano</td>
		<td><input type="text" name="ano_sel" value="<%=ano_sel%>"></td>
	</tr>
	<tr>
	    <td>&nbsp; Unidade:</td>
			<%strsql = "SELECT * FROM TBL_unidades where cd_status <> 0"
			Set rs_uni = dbconn.execute(strsql)%>
	    <td><select name="cd_unidade" class="inputs" onFocus="nextfield ='num_hospital';">
						<option value=""></option>
						<%Do While Not rs_uni.eof
						if int(cd_unidade) = rs_uni("cd_codigo") then%><%unidade_check = "selected"%><%end if%>
						<option value="<%=rs_uni("cd_codigo")%>" <%=unidade_check%>><%=rs_uni("nm_unidade")%></option>
						<%rs_uni.movenext
						unidade_check = ""
						loop%>
						</select>
						<a href="javascript:;" onclick="window.open('adm/unidade_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a>
		</td>
	<tr>		
		<td><input type="submit" name="ok" value="ok"></td>		
	</tr>	
		</form>
	</tr>
</table><br id="no_print">
<%if strcd_unidade <> "" Then

for i = 1 to 3%>
<table class="txt" border="0" width="600" align="center" style="border-collapse:collapse;>	
				<%condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				'strsql ="up_patrimonio_lista @ordem= '"&ordem&"', @condicao = '"&condicao&"'"
			if i = 1 then
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' AND cd_preventiva = 1 order by cd_posicao,cd_rack, nm_equipamento"
			elseif i = 2 then
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' AND cd_calibracao = 1 order by cd_posicao,cd_rack, nm_equipamento"
			elseif i = 3 then
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' AND cd_seg_elet = 1 order by cd_posicao,cd_rack, nm_equipamento"
			end if
			
			  	Set rs_patrimonio = dbconn.execute(strsql)
				
				Do while not rs_patrimonio.EOF
				cd_pat_codigo = rs_patrimonio("cd_codigo")
				cd_patrimonio = rs_patrimonio("cd_patrimonio")
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
				num_hospital =  rs_patrimonio("num_hospital")
				cd_rack = rs_patrimonio("cd_rack")
				num_hospital = rs_patrimonio("num_hospital")
				dt_data = rs_patrimonio("dt_data")
				
				cd_preventiva = rs_patrimonio("cd_preventiva")
				dt_periodicidade_prev = rs_patrimonio("dt_periodicidade_prev")
				
				cd_calibracao = rs_patrimonio("cd_calibracao")
				dt_periodicidade_cal = rs_patrimonio("dt_periodicidade_cal")
				
				cd_seg_elet = rs_patrimonio("cd_seg_elet")
				dt_periodicidade_elet = rs_patrimonio("dt_periodicidade_elet")
				
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
					strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"' "
				  	Set rs_espec = dbconn.execute(strsql_espec)
						if not rs_espec.EOF then
							nm_especialidade = rs_espec("nm_especialidade")
						end if
				
				if i = 1 then
					tipo_servico = "preventiva"
				elseif i = 2 then
					tipo_servico = "calibraçao"
				elseif i = 3 then
					tipo_servico = "segurança elétrica"
				end if
				
		estilo_titulo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:12px; color:black;"
		estilo_corpo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:12px;"%>
		
	<%if cabeca = 0 then%>
	<tr>		
		<td align="center">&nbsp;</td>
		<td align="center">&nbsp;</td>
		<td  valign="bottom"><img src="../imagens/px.gif" alt="" width="10" height="50" border="0"><br>
			<%if cd_unidade = 2 or cd_unidade = 3 or cd_unidade = 15 then%><img src="../imagens/logotipo_sao_luiz.gif" alt="" width="120" height="37" border="0"><%end if%>
			</td>
		<td colspan="12" align="center" valign="bottom" style="font-family:arial; font-size:12px; color:black;"><b>PLANEJAMENTO DE SERVIÇOS REFERENTE A <%=Ucase(tipo_servico)%> ANO <%=ano_sel%><br>VDLAP</b></td>
		<td colspan="4" align="right" valign="bottom"><img src="../imagens/logotipo_vdlap.gif" alt="" width="100" height="53" border="0"></td>
	</tr>
	<tr><td colspan="100%"><img src="../imagens/px.gif" alt="" width="10" height="25" border="0"></td></tr>
	<tr>
		<td align="center" valign="bottom">&nbsp;</td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>EQUIPAMENTO</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>CONTROLE DE<br>MANUTENÇÃO</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>MARCA</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>MODELO</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>PERIODICIDADE</b></td>		
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>JAN</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>FEV</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>MAR</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>ABR</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>MAI</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>JUN</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>JUL</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>AGO</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>SET</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>OUT</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>NOV</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>DEZ</b></td>
	</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');">
		<td align="center" valign="bottom">&nbsp;</td>
		<td valign="top" align="left" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=nm_equipamento%> <%'=dt_periodicidade_prev%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%="0"&cd_patrimonio%></td>				
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=nm_marca%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=cd_pn%></td>
		<%
		'SELECT     * FROM         TBL_patrimonio_manutencoes WHERE     (cd_patrimonio = 268) AND (cd_suspenso <> 1) AND (tipo_manut = 1) AND (dt_plan_prev >= '1/1/2009') ORDER BY dt_plan_prev
		'"&mes_sel&"/"&ano_sel&"%>
		
	<%if i = 1 then%>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=dt_periodicidade_prev%> <%if dt_periodicidade_prev <> "" Then response.write("meses") end if%></td>
		<%for imes = 1 to 12
			strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = "&i&" AND (dt_plan_prev >= '1/1/"&ano_sel&"') order by dt_plan_prev"
			Set rs_prev = dbconn.execute(strsql)
				if not rs_prev.EOF then
					cd_plan_prev = rs_prev("cd_codigo")
					dt_plan_prev = rs_prev("dt_plan_prev")
					dt_plan_prev_mes = month(dt_plan_prev)
					if int(dt_plan_prev_mes) = imes then
						'marca = dt_plan_prev'"X"
						marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&cd_plan_prev="&cd_plan_prev&"&acao=editar','nome','width=400,height=250'); return false;><img src='../imagens/x1.gif' height=15 border='0'></a>"
					else
						marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&cd_plan_prev="&cd_plan_prev&"','nome','width=400,height=250'); return false;><img src='../imagens/x0.gif' height=15 border='0'></a>"
					end if%>
				<%'rs_prev.movenext
				end if%>
		<td style="<%=estilo_corpo%>"><%=marca%></td>
		<%marca = ""
		next%>
		
	<%elseif i = 2 then%>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=dt_periodicidade_cal%> <%if dt_periodicidade_cal <> "" Then response.write("meses") end if%></td>
		<%for imes = 1 to 12
			strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = "&i&" AND (dt_plan_prev >= '1/1/"&ano_sel&"') order by dt_plan_prev"
			Set rs_prev = dbconn.execute(strsql)
				if not rs_prev.EOF then
					dt_plan_prev = rs_prev("dt_plan_prev")
					dt_plan_prev_mes = month(dt_plan_prev)
					if int(dt_plan_prev_mes) = imes then
						'marca = dt_plan_prev'"X"
						marca = "<img src='../imagens/x1.gif' height=15>"
					else
						marca = "<img src='../imagens/x0.gif' height=15>"
					end if%>
				<%'rs_prev.movenext
				end if%>
		<td style="<%=estilo_corpo%>"><%=marca%></td>
		<%marca = ""
		next%>
		
	<%elseif i = 3 then%>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=dt_periodicidade_elet%> <%if dt_periodicidade_elet <> "" Then response.write("meses") end if%></td>
		<%for imes = 1 to 12
			strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = "&i&" AND (dt_plan_prev >= '1/1/"&ano_sel&"') order by dt_plan_prev"
			Set rs_prev = dbconn.execute(strsql)
				if not rs_prev.EOF then
					dt_plan_prev = rs_prev("dt_plan_prev")
					dt_plan_prev_mes = month(dt_plan_prev)
					if int(dt_plan_prev_mes) = imes then
						'marca = dt_plan_prev'"X"
						marca = "<img src='../imagens/x1.gif' height=15>"
					else
						marca = "<img src='../imagens/x0.gif' height=15>"
					end if%>
				<%'rs_prev.movenext
				end if%>
		<td style="<%=estilo_corpo%>"><%=marca%></td>
		<%marca = ""
		next%>
	<%end if%>
	</tr>
		
	<%dt_plan_prev = ""
	cabeca = 1
	nm_especialidade = ""
	num_lista = num_lista + 1
	rs_patrimonio.movenext
	loop
	%>
	<%if cd_unidade = 2 or cd_unidade = 3 or cd_unidade = 15 then%>
	<tr>
		<td align="center" valign="bottom">&nbsp;</td>
		<td colspan="5">ITA20211.PQ.006 - ANEXO 3 - PLANEJAMENTO DE SERVIÇOS - VDLAP<p>OBS. As datas marcadas são referentes a <%=tipo_servico%></p></td>
	</tr>
	<%end if%>
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td align="center" colspan="15"></td>
		<td align="center" colspan="3"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>
	<tr>
		<td id="ok_print"><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
		<td id="no_print">&nbsp;</td>
			
		<td><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
	</tr><tr><td colspan="100%">&nbsp;</td></tr>
	


</table><%'if i <= 2 then%><!--p style="page-break-after:always;"></p--><%'end if%>
<div style="page-break-after:always;">&nbsp;</div>
<%cabeca = 0
next
	
end if%>

</body>
</html>
