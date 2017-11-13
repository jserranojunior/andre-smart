<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<%
cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)

dt_mes = request("mes_sel")
dt_ano = request("ano_sel")
	dia_final = ultimodiames(dt_mes,dt_ano)
	
	mes_anterior = dt_mes - 1
	ano_anterior = dt_ano
		if mes_anterior < 1 then
			mes_anterior = 12
			ano_anterior = dt_ano - 1
		end if
		dia_final_anterior = ultimodiames(mes_anterior,ano_anterior)

cd_diario = request("cd_diario")

if cd_diario = "" Then
	acao = "inserir"
else
	acao = "editar"
end if
%>
<html>
<head>
	<title>Financeiro - Diário</title>
</head>
<!--#include file="../js/ajax.js"-->
<!--#include file="../ferramentas/js/ferramentas.js"-->
<body ="window.close();">

<%cor_item1 = "#8f96e7"
cor_item2 = "#eaeaea"%>
<table border="1" align="center" style="font-size:12px; font-family:arial; border-collapse:collapse;">
	<tr style="background-color:<%=cor_item1%>; font-size:16px; color:white;">
		<td colspan="5">&nbsp;Observação <%'=cd_diario& "-" &dt_mes&"/1/"&dt_ano&" - "&dt_mes&"/"&dia_final&"/"&dt_ano%></td>		
	</tr>
	<%xsql = "up_financeiro_diario3_lista_individual @cd_diario='"&cd_diario&"', @dt_i='"&dt_mes&"/1/"&dt_ano&"', @dt_f='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i_a='"&mes_anterior&"/1/"&ano_anterior&"', @dt_f_a='"&mes_anterior&"/"&dia_final_anterior&"/"&ano_anterior&"'"
	Set rs_diario = dbconn.execute(xsql)
		if NOT rs_diario.EOF Then
			
			cd_valor = rs_diario("cd_codigo")
			cd_fornecedor = rs_diario("cd_fornecedor")
			nm_fornecedor = rs_diario("nm_fornecedor")
			nm_pagamento = rs_diario("nm_pagamento")
			cd_tipo_valor = rs_diario("cd_tipo_valor")
'			nm_tipo = rs_diario("nm_tipo")
			pre_ordem = rs_diario("pre_ordem")
			cd_valor = rs_diario("cd_valor")
			nm_valor = rs_diario("nm_valor")
			cd_qtd_parcelas = rs_diario("cd_qtd_parcelas")
			cd_area = rs_diario("cd_area")
			nm_area = rs_diario("nm_area")
			cd_centrocusto = rs_diario("cd_centrocusto")
			nm_centrocusto = rs_diario("nm_centrocusto")
			cd_conta = rs_diario("cd_conta")
			nm_conta = rs_diario("nm_conta")
			cd_unidade = rs_diario("cd_unidade")
			nm_unidade = rs_diario("nm_unidade")
			nm_sigla = rs_diario("nm_sigla")
			
			nm_valor = rs_diario("nm_valor")
			dt_vencimento = rs_diario("dt_vencimento")
			cd_parcela = rs_diario("cd_parcela")
			nm_obs = rs_diario("nm_obs")
		end if
	
	%>
	
	
	<tr>
		<td colspan="2" style="background-color:<%=cor_item2%>;">&nbsp;Favorecido</td>
		<td colspan="3"><b>&nbsp;<%=nm_fornecedor%></b></td>
	</tr>
	<tr>
		<td colspan="2" style="background-color:<%=cor_item2%>;">&nbsp;Histórico</td>
		<td colspan="3"><b>&nbsp;<%=nm_pagamento%></b></td>
	</tr>
	<tr>
		<td colspan="2" style="background-color:<%=cor_item2%>;">&nbsp;Tipo</td>
		<td colspan="3"><b>&nbsp;<%=nm_tipo%> &nbsp; &nbsp; &nbsp; <%if cd_tipo_valor = 3 then response.write(zero(cd_qtd_parcelas)&"x")%></b></td>
	</tr>
	
	<tr>
		<td><img src="../imagens/blackdot.gif" alt="" width="20" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
	</tr>
	<tr>
		<td colspan="2" style="background-color:<%=cor_item%>;">&nbsp;Área</td>
		<td style="background-color:#<%=cor_item%>;">&nbsp;C.Custos</td>
		<td style="background-color:#<%=cor_item%>;">&nbsp;Conta</td>
		<td style="background-color:#<%=cor_item%>;">&nbsp;Unidade</td>
		
	</tr>
	<tr>
		<td colspan="2"><b>&nbsp;<%=nm_area%></b></td>
		<td><b>&nbsp;<%=nm_centrocusto%></b></td>
		<td><b>&nbsp;<%=nm_conta%></b></td>
		<td><b>&nbsp;<%=nm_sigla%></b></td>
	</tr>
	<tr><td colspan="5"><img src="../imagens/blackdot.gif" alt="" width="450" height="3" border="0"></td></tr>
	<tr style="background-color:<%=cor_item1%>; font-size:12px; color:white;">
		<td><b>&nbsp;</b></td>
		<td align="center">Data Vencimento</td>
		<td align="center">Valor</td>
		<td align="center">Qtd. Parcelas</td>
		<td align="center">Parcela</td>
	</tr>
		
	<tr>
		<td align="right">&nbsp;</td>
		<td align="center"><%=zero(day(dt_vencimento))&"/"&left(mes_selecionado(month(dt_vencimento)),3)&"/"&year(dt_vencimento)%></td>
		<td align="right"><%=formatNumber(nm_valor)%>&nbsp;</td>
		<td align="center"><b><%=cd_qtd_parcelas%></b></td>
		<td align="center"><b><%=cd_parcela%></b></td>
	</tr>
	<tr><td colspan="5"><img src="../imagens/blackdot.gif" alt="" width="450" height="3" border="0"></td></tr>
	
	<form action="acoes/acoes3.asp" name="observacoes" id="observacoes">
	<input type="hidden" name="acao" value="edita_obs">
	<input type="hidden" name="cd_valor" value="<%=cd_valor%>">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="data_atual" value="<%=data_atual%>">
	
	<tr>
		<td colspan="2" align="center" style="background-color:<%=cor_item2%>;"><b>Observações</b></td>
		<td colspan="3"><textarea cols="40" rows="2" name="nm_obs"><%if nm_obs <> "" Then response.write(nm_obs)%></textarea></td>
	
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
		<td colspan="3"><input type="submit" name="OK" value=" OK "></td>
	</tr>
	</form>
</table>
</body>
</html>
