<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<%
cd_user = session("cd_codigo")
pasta_loc = "financeiro\"
arquivo_loc = "diario_lista_pag3.asp"

data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)
dt_mes = request("mes_sel")
dt_ano = request("ano_sel")
	if dt_mes = "" Then dt_mes = month(now)
	if dt_ano = "" Then dt_ano = year(now)

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
<body onLoad="diario.cd_fornecedor.focus()">
<!--#include file="../includes/arquivo_loc.asp"-->


<table border="1" align="center" style="font-size:12px; font-family:arial; border-collapse:collapse;">
	<tr>
		<td><img src="../imagens/blackdot.gif" alt="" width="20" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
	</tr>
	<tr style="background-color:green; font-size:16px; color:white;">
		<td colspan="5">&nbsp;Listagem dos últimos pagamentos - v3 <%'=cd_diario& "-" &dt_mes&"/1/"&dt_ano&" - "&dt_mes&"/"&dia_final&"/"&dt_ano%></td>		
	</tr>
	
	<%xsql = "up_financeiro_diario3_lista_individual @cd_diario='"&cd_diario&"', @dt_i='"&dt_mes&"/1/"&dt_ano&"', @dt_f='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i_a='"&mes_anterior&"/1/"&ano_anterior&"', @dt_f_a='"&mes_anterior&"/"&dia_final_anterior&"/"&ano_anterior&"'"
	Set rs_diario = dbconn.execute(xsql)
		if NOT rs_diario.EOF Then
			
			cd_fornecedor = rs_diario("cd_fornecedor")
			nm_fornecedor = rs_diario("nm_fornecedor")
			nm_pagamento = rs_diario("nm_pagamento")
			cd_tipo_valor = rs_diario("cd_tipo_valor")
			'nm_tipo = rs_diario("nm_tipo")
			nm_tipo = rs_diario("nm_tipo_pagamento")
			pre_ordem = rs_diario("pre_ordem")
			cd_valor = rs_diario("cd_valor")
			nm_valor = rs_diario("nm_valor")
			cd_qtd_parcelas = rs_diario("cd_qtd_parcelas")
				if cd_qtd_parcelas = 1 then
					cd_qtd_parcelas = 12
				end if
			cd_area = rs_diario("cd_area")
			nm_area = rs_diario("nm_area")
			cd_centrocusto = rs_diario("cd_centrocusto")
			nm_centrocusto = rs_diario("nm_centrocusto")
			cd_conta = rs_diario("cd_conta")
			nm_conta = rs_diario("nm_conta")
			cd_unidade = rs_diario("cd_unidade")
			nm_unidade = rs_diario("nm_unidade")
			nm_sigla = rs_diario("nm_sigla")
			dt_inicial = rs_diario("dt_inicial")
				dt_inicial = zero(day(dt_inicial))&"/"&zero(month(dt_inicial))&"/"&right(year(dt_inicial),2)
			dt_final = rs_diario("dt_final")
				if not ISNULL(dt_final) then 
					dt_final = "Fim: "&zero(day(dt_final))&"/"&zero(month(dt_final))&"/"&right(year(dt_final),2)
				else
					dt_final = "Conta ativa"
				end if
		end if
	
	%>
	
	
	<tr>
		<td colspan="2" style="background-color:#eaeaea;">&nbsp;Favorecido</td>
		<td colspan="3"><b>&nbsp;<%=nm_fornecedor%></b></td>
	</tr>
	<tr>
		<td colspan="2" style="background-color:#eaeaea;">&nbsp;Descr. / Histórico</td>
		<td colspan="3"><b>&nbsp;<%=nm_pagamento%></b></td>
	</tr>
	<tr>
		<td colspan="2" style="background-color:#eaeaea;">&nbsp;Tipo</td>
		<td colspan="3"><b>&nbsp;<%=nm_tipo%> &nbsp; &nbsp; &nbsp; <%if cd_tipo_valor = 3 then response.write(zero(cd_qtd_parcelas)&"x")%></b></td>
	</tr>
	<tr>
		<td colspan="2" style="background-color:#eaeaea;">Vigência</td>
		<td align="left"><b>Inicio: <%=dt_inicial%></b></td>
		<td align="left"><b><%=dt_final%></b></td>
		<td>&nbsp;</td>
	</tr>
</table>	
<br>
<table border="1" align="center" style="font-size:12px; font-family:arial; border-collapse:collapse;">
	
	<tr>
		<td><img src="../imagens/blackdot.gif" alt="" width="20" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="105" height="3" border="0"></td>
	</tr>
	<tr>
		<td colspan="2" style="background-color:#eaeaea;">&nbsp;Área</td>
		<td style="background-color:#eaeaea;">&nbsp;C.Custos</td>
		<td style="background-color:#eaeaea;">&nbsp;Conta</td>
		<td style="background-color:#eaeaea;">&nbsp;Unidade</td>
		
	</tr>
	<tr>
		<td colspan="2"><b>&nbsp;<%=nm_area%></b></td>
		<td><b>&nbsp;<%=nm_centrocusto%></b></td>
		<td><b>&nbsp;<%=nm_conta%></b></td>
		<td><b>&nbsp;<%=nm_sigla%></b></td>
	</tr>
	<tr><td colspan="5"><img src="../imagens/blackdot.gif" alt="" width="450" height="3" border="0"></td></tr>
	<tr style="background-color:green; font-size:12px; color:white;">
		<td><b>&nbsp;</b></td>
		<td align="center">Data Vencimento</td>
		<td align="center">Valor</td>
		<td colspan="2">Descrição</td>
	</tr>
		<%'response.write(cd_qtd_parcelas)
		if cd_qtd_parcelas = "" Then
			cd_qtd_parcelas = 1
		end if 
		
		strsql = "SELECT TOP "&cd_qtd_parcelas&" * FROM TBL_financeiro_valores3 WHERE cd_diario = "&cd_diario&" AND dt_vencimento <= '"&month(now)&"/"&ultimodiames(month(now),year(now))&"/"&year(now)&"' ORDER BY dt_vencimento DESC"
		Set rs_hist_val = dbconn.execute(strsql)
		While NOT rs_hist_val.EOF
			conta_vezes = conta_vezes + 1
			
			nm_pagamento = rs_hist_val("nm_pagamento")
			nm_valor = rs_hist_val("nm_valor")
			dt_vencimento = rs_hist_val("dt_vencimento")
			cd_parcela = rs_hist_val("cd_parcela")
				if conta_vezes = 1 then
					total_parcelas_pagas = cd_parcela
					maior_valor_pago = nm_valor
					dt_vencimento_maior = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&year(dt_vencimento)
					menor_valor_pago = nm_valor
					dt_vencimento_menor = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&year(dt_vencimento)
				end if
			
			soma_valores = soma_valores + abs(nm_valor)
			media_valores = soma_valores/conta_vezes
			
				if maior_valor_pago < nm_valor then
					maior_valor_pago = nm_valor
					dt_vencimento_maior = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&year(dt_vencimento)
				end if
			
				if menor_valor_pago > nm_valor then
					menor_valor_pago = nm_valor
					dt_vencimento_menor = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&year(dt_vencimento)
				end if
				
				'diferenca_valores = maior_valor_pago - menor_valor_pago
				
				if conta_vezes = 1 then
					valor_ultimo_mes = nm_valor
				elseif conta_vezes = 2 then
					valor_penultimo_mes = nm_valor
				end if%>
		<tr>
			<td align="right"><%if cd_tipo_valor = 3 then response.write(cd_parcela) else response.write("&nbsp;") end if%></td>
			<td align="center"><%=zero(day(dt_vencimento))&"/"&left(mes_selecionado(month(dt_vencimento)),3)&"/"&year(dt_vencimento)%></td>
			<td align="right"><%=formatNumber(nm_valor)%></td>
			<td align="left" colspan="2"><%=left(nm_pagamento,30)%><%if len(nm_pagamento) > 30 Then response.write("...")%></td>
			
		</tr>
			
		<%rs_hist_val.movenext
		wend
		
			diferenca_ultimos_valores = valor_ultimo_mes - valor_penultimo_mes
					status_variacao = instr(1,diferenca_ultimos_valores,"-",1)
						if cd_tipo_valor <> 1  AND valor_penultimo_mes <> "" then
							if diferenca_ultimos_valores > 0 then
								variacao_valores = formatnumber((diferenca_ultimos_valores/valor_penultimo_mes)*100,2)
							else
								variacao_valores = "x"'formatnumber((diferenca_ultimos_valores/valor_penultimo_mes)*100,2)
							end if
						else
							valor_penultimo_mes = 0
							if diferenca_ultimos_valores > 0 then
								variacao_valores = formatnumber((diferenca_ultimos_valores/valor_ultimo_mes)*100,2)
							ELSE
								variacao_valores = "y"
							end if
						end if%>
		
		<%if cd_tipo_valor = 3 then '*** Valor parcelado *** 
		if nm_valor > 0 then total_pago = total_parcelas_pagas * abs(nm_valor)
				qtd_parcelas_abertas = cd_qtd_parcelas - total_parcelas_pagas
			saldo = qtd_parcelas_abertas * abs(nm_valor)
			total_final = total_pago+saldo%>
			<tr>
				<td colspan="5"><img src="../imagens/blackdot.gif" alt="" width="450" height="3" border="0"></td>
			</tr>
			<tr style="background-color:#eaeaea;">
				<td colspan="2" align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2" align="center" style="background-color:#eaeaea;"><b>Total Pago</b></td>
				<td align="right"><%if total_pago <> "" Then response.write(formatnumber(total_pago))%>&nbsp;</td>
				<td align="center"><%=total_parcelas_pagas%> parcela<%if total_parcelas_pagas > 1 Then%>s<%end if%></td>
				<td align="center"><%'="<span style=color:gray;>"&(total_pago)&"%</span>"%><%="<span style=color:gray;>"&formatnumber((total_pago/total_final)*100,2)&"%</span>"%> </td>
			</tr>
			<tr>
				<td colspan="2" align="center" style="background-color:#eaeaea;"><b>Saldo</b></td>
				<td align="right"><%if saldo <> "" Then response.write(formatnumber(saldo))%>&nbsp;</td>
				<td align="center"><%=qtd_parcelas_abertas%> parcela<%if qtd_parcelas_abertas > 1 Then%>s<%end if%></td>
				<td align="center"><%'="<span style=color:gray;>"&(saldo)*100&"%</span>"%><%="<span style=color:gray;>"&formatnumber((saldo/total_final)*100,2)&"%</span>"%></td>
			</tr>
			<tr>
				<td colspan="5"><img src="../imagens/blackdot.gif" alt="" width="450" height="1" border="0"></td>
			</tr>
			<tr>
				<td colspan="2" align="center" style="background-color:silver;"><b>Total Final</b></td>
				<td align="right"><b><%if total_final <> "" Then response.write(formatnumber(total_final))%>&nbsp;</b></td>
				<td align="center"><%=cd_qtd_parcelas%> parcela<%if cd_qtd_parcelas > 1 Then%>s<%end if%></td>
				<td>&nbsp;</td>
			</tr>
		<%elseif cd_tipo_valor = 2 then '*** Valores Fixos ***%>	
			<tr>
				<td colspan="5"><img src="../imagens/blackdot.gif" alt="" width="450" height="3" border="0"></td>
			</tr>
			<tr>
				<td colspan="2" align="center" style="background-color:#eaeaea;"><b>Maior Valor</b></td>
				<td align="right"><%if maior_valor_pago <> "" Then response.write(formatnumber(maior_valor_pago))%>&nbsp;</td>
				<td align="center"><%=dt_vencimento_maior%></td>
				
			</tr>
			<tr>
				<td colspan="2" align="center" style="background-color:#eaeaea;"><b>Menor Valor</b></td>
				<td align="right"><%if menor_valor_pago <> "" Then response.write(formatnumber(menor_valor_pago))%>&nbsp;</td>
				<td align="center"><%=dt_vencimento_menor%></td>
			</tr>
			<tr>
				<td colspan="5"><img src="../imagens/blackdot.gif" alt="" width="450" height="1" border="0"></td>
			</tr>
			<tr>
				<td colspan="2" align="center" style="background-color:#eaeaea;"><b>Média</b></td>
				<td align="right"><b><%if media_valores <> "" Then response.write(formatnumber(media_valores))%></b>&nbsp;</td>
				<!--td colspan="2">&nbsp;</td-->
			</tr>
			<tr><td colspan="5">&nbsp;</td></tr>
			<tr style="background-color:#eaeaea;">
				
				<td colspan="2" align="center">Penultimo mês</td>
				<td align="center">Último mês</td>
				<td align="center">Diferença</td>
				<td align="center">Variação</td>
				<!--td align="center" rowspan="2"><%'=formatnumber(diferenca_valores,2)&" <br><span style=color:gray;>("&formatnumber((diferenca_valores/maior_valor_pago)*100,3)&"%)</span>"%></td-->
			</tr>
			<tr>
				<%if valor_ultimo_mes < valor_penultimo_mes then
					cor_variacao = "green"
					seta = "ic_indicador_down.gif"
				elseif valor_ultimo_mes = valor_penultimo_mes then
					cor_variacao = "gray"
					seta = "ic_indicador_ok.gif"
				elseif valor_ultimo_mes > valor_penultimo_mes then
					cor_variacao = "red"
					seta = "ic_indicador_up.gif"
				end if%>
				<td align="center"><img src="../imagens/<%=seta%>" alt="" width="15" height="13" border="0"></td>
				<td align="center"><%if valor_penultimo_mes <> 0 Then response.write(formatnumber(valor_penultimo_mes,2))%></td>
				<td align="center"><%if valor_ultimo_mes <> 0 Then response.write(formatnumber(valor_ultimo_mes,2))%></td>
				<td align="center"><%=formatnumber(diferenca_ultimos_valores,2)%></td>
				<td align="center" style="color:<%=cor_variacao%>;"><b><%=variacao_valores&"%"%></b>&nbsp;</td>
			</tr>
			
		<%end if%>
	</tr>
</table>
</body>
</html>
