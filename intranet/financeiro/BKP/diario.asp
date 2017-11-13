
<%'Dim dt_mes, dt_ano

dt_ano = request("dt_ano")
dt_mes = request("dt_mes")
dia_final = ultimodiames(dt_mes,dt_ano)

	if dt_mes = "" OR dt_ano = "" then
		dt_mes = month(now)
		dt_ano = year(now)
	end if
	
	mes_anterior = dt_mes - 1
	ano_anterior = dt_ano
		if mes_anterior < 1 then
			mes_anterior = 12
			ano_anterior = dt_ano - 1
		end if
	mes_posterior = dt_mes + 1
	ano_posterior = dt_ano
		if mes_posterior > 12 then
			mes_posterior = 1
			ano_posterior = dt_ano + 1
		end if
		
cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)%>

<br class="no_print">
<table cellspacing="1" cellpadding="1" border="0" class="no_print">
	<form action="../financeiro.asp" name="seleciona_período" id="seleciona_período">
	<input type="hidden" name="tipo" value="diario">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="data_atual" value="<%=data_atual%>">
	
	<tr>
		<td align="center" colspan="2">CONTAS A PAGAR</td>
	</tr>
	<tr>
		<td><b>M&ecirc;s</b></td>
		<td><b>Ano</b></td>
	</tr>
	<tr>
		
		<td><select name="dt_mes">
					<option value=""></option>
					<%for imes = 1 to 12%>
						<option value="<%=imes%>" <%if abs(imes)=abs(dt_mes) then response.write("SELECTED")%>><%=mes_selecionado(imes)%></option>
					<%next%>
			</select></td>	
		<td><input type="text" name="dt_ano" maxlength="4" size="4" value="<%=dt_ano%>">&nbsp;<input type="submit" name="OK" value="OK" width="4" height="5"></td>
	</tr>
	</form>
	<tr>
		<td align="center" colspan="2"><a href="financeiro.asp?tipo=diario&dt_mes=<%=mes_anterior%>&dt_ano=<%=ano_anterior%>"><< Anterior </a>&nbsp; &nbsp; &nbsp; <a href="financeiro.asp?tipo=diario&dt_mes=<%=month(now)%>&dt_ano=<%=year(now)%>">Atual</a> &nbsp; &nbsp; &nbsp; <a href="financeiro.asp?tipo=diario&dt_mes=<%=mes_posterior%>&dt_ano=<%=ano_posterior%>">Próximo>></a></td>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2"><a href="javascript:void(0);" onClick="window.open('financeiro/diario_cad.asp?mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','Cadastro','width=500,height=480,scrollbars=1')" return false;"><b>(+)</b> Incluir</a> &nbsp;</td>
	</tr>
	</tr>
</table>
<br class="no_print">
<!-- Monta a tabela separando as semanas em blocos -->
<table cellspacing="1" cellpadding="1" border="1" style="1">
	
	<tr>
		<td align="center" colspan="10">Diário de <%=mes_selecionado(dt_mes)%> de <%=dt_ano%></td>
	</tr>
	<tr>
		<td align="center"><b>Área</b></td>
		<td align="center"><b>C.C</b></td>
		<td align="center"><b>Conta</b></td>
		<td align="center"><b>Unidade</b></td>
		<td align="center"><b>Data</b></td>
		<td><b>Tipo</b></td>
		<td><b>Histórico</b></td>
		<!--td><b>Anterior</b></td-->
		<td align="center"><b>Atual</b></td>
		<td align="center"><b>Extra</b></td>
		<!--td><b>Pagamento</b></td-->
	</tr>
	<%strsql = "Select * From view_financeiro_diario Where dt_vencimento between '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"' OR cd_tipo_valor = 2 AND year(dt_vencimento) <= '"&dt_ano&"' order by day(dt_vencimento)"
		Set rs = dbconn.execute(strsql)
		
		while not rs.EOF
			nm_area = rs("nm_area")
			nm_cc = rs("nm_cc")
			nm_centrocusto = rs("nm_centrocusto")
			nm_conta = rs("nm_conta")
			nm_unidade = rs("nm_unidade")
			nm_sigla = rs("nm_sigla")
			dt_vencimento = rs("dt_vencimento")
				dt_vencimento = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&year(dt_vencimento)
			cd_fornecedor = rs("cd_cedente")
			nm_fornecedor = rs("nm_fornecedor")
			nm_despesa = rs("nm_despesa")
			nm_valor = rs("nm_valor")
			cd_tipo_valor = rs("cd_tipo_valor")
			nm_tipo_pag_abrv = rs("nm_tipo_pag_abrv")
			cd_parcela = rs("cd_parcela")
			cd_qtd_parcelas = rs("cd_qtd_parcelas")
			
				if cd_tipo_valor = 1 then
					nm_valor_extra = nm_valor
					nm_valor = 0 
				else
					nm_valor = nm_valor
					nm_valor_extra = 0
				end if
		%>
	<tr>
		<td align="center">&nbsp;<%=nm_area%></td>
		<td align="center">&nbsp;<%=nm_cc%></td>
		<td align="center">&nbsp;<%=nm_conta%></td>
		<td align="center">&nbsp;<%=nm_sigla%></td>
		<td align="center">&nbsp;<%=dt_vencimento%></td>
		<td align="center">&nbsp;<%=nm_tipo_pag_abrv%></td>
		<td>&nbsp;<%if cd_fornecedor > 0 then response.write("<b>"&nm_fornecedor&"</b>-")%><%=nm_despesa%>&nbsp;
										<%if cd_qtd_parcelas > 1 then
											response.write("<span style='color:red;'><b>"&cd_parcela&"/"&cd_qtd_parcelas&"</b></span>")
										end if%></td>
		<!--td>&nbsp;<%%>-</td-->
		<td align="right">&nbsp;<%if nm_valor <> "0" then 
									response.write(FormatNumber(nm_valor,2))
								end if%></td>
		<td align="right">&nbsp;<%if nm_valor_extra <> "0" then 
									response.write(FormatNumber(nm_valor_extra,2))
								end if%></td>
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr>	
	<%soma_valores = abs(soma_valores) + nm_valor
	soma_extra = abs(soma_extra) + nm_valor_extra
	
		rs.movenext
		wend%>
	<tr>
		<td><img src="../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="70" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="500" height="1" border="0"></td>
		<!--td>&nbsp;<%%>-</td-->
		<td align="right"><%=formatNumber(soma_valores,2)%></td>
		<td align="right"><%=formatNumber(soma_extra,2)%></td>
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr>
	<%total_valores = soma_valores + soma_extra%>
	<tr>
		<td colspan="7">&nbsp;</td>
		<!--td>&nbsp;<%%>-</td-->
		<td align="center" colspan="2">&nbsp;<b><%=FormatNumber(total_valores,2)%></b></td>
		
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr>
	<%total = abs(total) + nm_valor%>
	
</table>