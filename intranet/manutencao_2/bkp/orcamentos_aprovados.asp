<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<html>
<head>
	<title>Orçamentos Aprovados</title>
</head>
<!-- #include file="../css/geral.htm"-->
<body>
<%tipo = request("tipo")%>

<table align="center" border="1" style="border-collapse:collapse;">
<tr><td align="center" colspan="8" style="background-color:gray; color:white; font-size:14px;"><b>Listagem de O.S com orçamento aprovado</b></td></tr>
<tr style="background-color:silver; font-size:12px;">
	<td align="center">Nº<br><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
	<td align="center">Favorecidos<br><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
	<td align="center">O.S.<br><img src="../imagens/px.gif" alt="" width="55" height="1" border="0"></td>
	<td align="center">Item<br><img src="../imagens/px.gif" alt="" width="230" height="1" border="0"></td>
	<td align="center">Nº Orç.<br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
	<td align="center">Tipo<br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
	<td align="center">Data Aprov<br><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
	<td align="center">Valor<br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
</tr>
<%num_linha = 1
total_valor = 0
strsql = "SELECT * FROM VIEW_os_lista2 WHERE fecha <> 1 AND cd_gestao BETWEEN 2 AND 3 AND cd_pagto is null ORDER BY  nm_fornecedor, dt_resposta_orc, num_os_assist, cd_gestao"
Set rs = dbconn.execute(strsql)

while not rs.EOF
	cd_os = rs("cd_codigo")
	cd_fornecedor = rs("cd_fornecedor")
	nm_fornecedor = rs("nm_fornecedor")
	cd_equipamento = rs("cd_equipamento")
	nm_equipamento_novo = rs("nm_equipamento_novo")
	cd_unidade = rs("cd_unidade_os")
	num_os = rs("num_os")
	cd_valor_orc = rs("cd_valor_orc")
		
	cd_gestao = rs("cd_gestao")
	num_os_assist = rs("num_os_assist")
	dt_resposta_orc = rs("dt_resposta_orc")
		
		if isnumeric(cd_valor_orc) Then total_valor = total_valor + cd_valor_orc%>
	<%if ult_num_assist <> num_os_assist then%>
		<tr style="background-color:silver;">
			<td colspan="5">&nbsp;</td>
			<td align="center"><b>Total Orç.</b></td>
			<td align="right" colspan="2"><b>
				<%if tipo = "orc_aprov" then%><%'=total_orcamento%><%=formatnumber(total_orcamento,2)%>
				<%elseif tipo = "financ" then%><a href="javascript:void(0);" return false;" onclick="window.open('../financeiro/diario_cad3.asp?cd_origem=<%=2&"."&cd_os%>&cd_forn=<%=cd_fornecedor%>&num_os_assist=<%=ult_num_assist%>&cd_valor_orc=<%=total_orcamento%>&cd_tipo_orc=<%=cd_gestao%>&cd_unidade=<%=cd_unidade%>&campo=cd_codigo&visual=1&jan=1','pagamento<%response.write(replace(ult_num_assist,"/",""))%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=formatnumber(total_orcamento,2)%><%end if%></a></b></td>
		</tr>
		<tr><td colspan="8">&nbsp;</td></tr>
	<%total_orcamento = 0
	end if%>
	<tr>
		<td align="center"><%=num_linha%></td>
		<td><b><%=nm_fornecedor%></b></td>
		<td align="center">
			<%if tipo = "orc_aprov" then%><a href="javascript:void(0);" return false;" onclick="window.open('../manutencao_2/manutencao_andamento2.asp?cd_codigo=<%=cd_os%>&campo=cd_codigo&visual=1&jan=1', 'janela_os<%=cd_os%>', 'width=650, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=zero(cd_unidade)&"."&num_os%></a>
			<%elseif tipo = "financ" then%><a href="javascript:void(0);" return false;" onclick="window.open('../manutencao_2/manutencao_ver_janela2.asp?cd_codigo=<%=cd_os%>&campo=cd_codigo&visual=1&jan=1', 'janela_os<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=zero(cd_unidade)&"."&num_os%></a><%end if%></td>
		
		<td><%=nm_equipamento_novo%></td>
		<td align="center"><%=num_os_assist%></td>
		<td><%if cd_gestao = 2 then%>Compra<%else%>Manutenção<%end if%></td>
		<td align="center"><%if dt_resposta_orc <> "" Then%><%=zero(day(dt_resposta_orc))&"/"&zero(month(dt_resposta_orc))&"/"&right(year(dt_resposta_orc),2)%><%else%>&nbsp;<%end if%></td>
		<td align="right">
			<%if isnumeric(cd_valor_orc) Then%>
				<!--a href="javascript:void(0);" return false;" onclick="window.open('diario_cad3.asp?cd_origem=<%=2&"."&cd_os%>&cd_forn=<%=cd_fornecedor%>&cd_equip=<%=cd_equipamento%>&cd_valor_orc=<%=cd_valor_orc%>&cd_tipo_orc=<%=cd_gestao%>&cd_unidade=<%=cd_unidade%>&campo=cd_codigo&visual=1&jan=1', 'pagamento_<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"--><%=formatnumber(cd_valor_orc,2)%><!--/a-->
			<%else%>&nbsp;<%end if%></td>
	</tr>	
<%
if isnumeric(cd_valor_orc) Then total_orcamento = abs(total_orcamento) + abs(cd_valor_orc)
ult_num_assist = num_os_assist
num_linha = num_linha + 1
rs.movenext
wend%>
	<tr>
		<td colspan="5">&nbsp;</td>
		<td align="center"><b>Total Orç.</b></td>
		<td align="right" colspan="2"><b><%if tipo = "financ" then%><a href="javascript:void(0);" return false;" onclick="window.open('diario_cad3.asp?cd_origem=<%=2&"."&cd_os%>&cd_forn=<%=cd_fornecedor%>&cd_equip=<%=cd_equipamento%>&cd_valor_orc=<%=cd_valor_orc%>&cd_tipo_orc=<%=cd_gestao%>&cd_unidade=<%=cd_unidade%>&campo=cd_codigo&visual=1&jan=1', 'pagamento_<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=formatnumber(total_orcamento,2)%></a><%else%><%=formatnumber(total_orcamento,2)%><%end if%></b></td>
	</tr>
	<tr><td colspan="8">&nbsp;</td></tr>
	<tr style="background-color:gray;">
		<td colspan="5">&nbsp;</td>
		<td align="center"><b>Total</b></td>
		<td align="right" style="font-size:11px;" colspan="2"><b><%if total_valor <> "" Then response.write(formatnumber(total_valor,2))%></b></td>
	</tr>
</table>
</body>
</html>
