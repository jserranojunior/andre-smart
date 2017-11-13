<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<html>
<head>
	<title>Orçamentos Aprovados</title>
</head>
<!-- #include file="../css/geral.htm"-->
<body>
<%tipo = request("tipo")
cd_orcamento = request("cd_orcamento")%>

<table align="center" border="1" style="border-collapse:collapse;">
<tr><td align="center" colspan="7" style="background-color:gray; color:white; font-size:14px;"><b>Orçamento Aprovado</b></td></tr>
<tr style="background-color:silver; font-size:12px;">
	<td>&nbsp;</td>
	<td align="center">Fornecedor<br><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
	<td align="center">Nº Orç.<br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
	<td align="center">Data Aprov<br><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
	<td align="center">Valor<br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
	<td align="center" colspan="2">Parc<br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
</tr>

<%num_linha = 1
total_valor = 0

strsql ="SELECT * FROM View_manutencao_orcamento WHERE cd_codigo = '"&cd_orcamento&"' ORDER BY dt_orcamento"
		Set rs_orc = dbconn.execute(strsql)
			if not rs_orc.EOF then
				nm_fornecedor = rs_orc("nm_fornecedor")
				nr_orcamento = rs_orc("nr_orcamento")
				dt_orcamento = rs_orc("dt_orcamento")
					if dt_orcamento <> "" Then	
						dia_orc = zero(day(dt_orcamento))
						mes_orc = zero(month(dt_orcamento))
						ano_orc = year(dt_orcamento)
						dt_orcamento = dia_orc&"/"&mes_orc&"/"&ano_orc
					end if
					
				nm_valor = rs_orc("nm_valor")
					'total_nm_valor = total_nm_valor + nm_valor
					if nm_valor = "" then nm_valor = 0
					
				nr_parcela = rs_orc("nr_parcela")
					if nr_parcela > 1 then
						nr_valor_parcela = nm_valor / nr_parcela
					else
						nr_valor_parcela = 0
					end if%>
					<tr>
						<td>&nbsp;</td>
						<td><b><%=nm_fornecedor%></b></td>
						<td align="center"><b><%=nr_orcamento%></b></td>
						<td align="center"><b><%=dt_orcamento%></b></td>
						<td align="center"><b><%if nm_valor <> "" Then response.write(formatNumber(nm_valor,2))%></b></td>
						<%if nr_parcela > 1 then%>
							<td><b><%=nr_parcela%></b></td>
							<td><b><%if nr_valor_parcela <> "" Then response.write(formatnumber(nr_valor_parcela,2))%></b></td>
						<%else%>
							<td colspan="2">&nbsp;</td>
						<%end if%>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr style="background-color:silver;">
						<td>&nbsp;</td>
						<td align="center" colspan="2">Item</td>
						<td align="center">OS</td>
						<td align="center">Tipo</td>
						<td align="center" colspan="2">Valor</td>
					</tr>
					
					<%
					
			'rs_orc.movenext
			end if
		filtro = "orc_cad_upd"
		nm_botao = "Altera"
		

strsql = "SELECT * FROM VIEW_os_lista2 WHERE num_os_assist = '"&nr_orcamento&"' ORDER BY dt_resposta_orc"
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
	if isnumeric(cd_valor_orc) Then valor_total_orcamento = valor_total_orcamento + abs(cd_valor_orc)
	tipo = "orc_aprov"
	cd_gestao = rs("cd_gestao")
	num_os_assist = rs("num_os_assist")
	dt_resposta_orc = rs("dt_resposta_orc")
		
		if isnumeric(cd_valor_orc) Then total_valor = total_valor + cd_valor_orc%>
	
	<tr>
		<td align="center"><%=num_linha%></td>
		<td colspan="2"><%=nm_equipamento_novo%></td>
		<td align="center">
			<%if tipo = "orc_aprov" then%><a href="javascript:void(0);" return false;" onclick="window.open('../manutencao_2/manutencao_andamento2.asp?cd_codigo=<%=cd_os%>&campo=cd_codigo&visual=1&jan=1', 'janela_os<%=cd_os%>', 'width=650, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=zero(cd_unidade)&"."&num_os%></a>
			<%elseif tipo = "financ" then%><a href="javascript:void(0);" return false;" onclick="window.open('../manutencao_2/manutencao_ver_janela2.asp?cd_codigo=<%=cd_os%>&campo=cd_codigo&visual=1&jan=1', 'janela_os<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=zero(cd_unidade)&"."&num_os%></a><%end if%></td>
		<td><%if cd_gestao = 2 then%>Compra<%else%>Manutenção<%end if%></td>
		<td align="right" colspan="2">
			<%if isnumeric(cd_valor_orc) Then%>
				<!--a href="javascript:void(0);" return false;" onclick="window.open('diario_cad3.asp?cd_origem=<%=2&"."&cd_os%>&cd_forn=<%=cd_fornecedor%>&cd_equip=<%=cd_equipamento%>&cd_valor_orc=<%=cd_valor_orc%>&cd_tipo_orc=<%=cd_gestao%>&cd_unidade=<%=cd_unidade%>&campo=cd_codigo&visual=1&jan=1', 'pagamento_<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"-->
					<%=formatnumber(cd_valor_orc,2)%>
				<!--/a-->
			<%else%>
				&nbsp;
			<%end if%></td>
				
	</tr>	
<%
nm_valor = 0
ult_num_assist = num_os_assist
num_linha = num_linha + 1
rs.movenext
wend%>
	
	<tr><td colspan="7">&nbsp;</td></tr>
	<tr style="background-color:gray;">
		<td colspan="4">&nbsp;</td>
		<td align="center"><b>Total</b></td>
		<td align="right" style="font-size:11px;" colspan="2"><b><%if valor_total_orcamento <> "" Then response.write(formatnumber(valor_total_orcamento,2))%><%'if valor_total_orcamento <> "" Then response.write(formatnumber(valor_total_orcamento,2))%></b></td>
	</tr>
</table>
</body>
</html>
