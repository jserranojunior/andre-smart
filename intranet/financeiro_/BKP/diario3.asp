
<%'Dim dt_mes, dt_ano
ancora = request("ancora")

dt_ano = request("dt_ano")
dt_mes = request("dt_mes")

dt_ano = request("ano_sel")
dt_mes = request("mes_sel")
	if dt_mes = "" OR dt_ano = "" then
		dt_mes = month(now)
		dt_ano = year(now)
	end if
		dia_final = ultimodiames(dt_mes,dt_ano)
		
		mes_anterior = dt_mes - 1
		ano_anterior = dt_ano
			if mes_anterior < 1 then
				mes_anterior = 12
				ano_anterior = dt_ano - 1
			end if
			dia_final_anterior = ultimodiames(mes_anterior,ano_anterior)
			'response.write("/"&mes_anterior&"/"&ano_anterior)
		
		mes_posterior = dt_mes + 1
		ano_posterior = dt_ano
			if mes_posterior > 12 then
				mes_posterior = 1
				ano_posterior = dt_ano + 1
			end if
			dia_final_posterior = ultimodiames(mes_posterior,ano_posterior)



cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)
	
	'if  year(now)&zero(month(now))&zero(day(now)) > dt_ano&zero(dt_mes)&dia_final then
	if  year(now)&zero(month(now)) > dt_ano&zero(dt_mes) then
		ordem = " dt_vencimento "'
	elseif year(now)&zero(month(now)) = dt_ano&zero(dt_mes) then
		ordem = " pre_ordem "
	elseif year(now)&zero(month(now)) < dt_ano&zero(dt_mes) then
		ordem = " dia_vencimento_padrao "
		'ordem = " pre_ordem "
	end if
	
qtd_semanas = datediff("ww","1/"&dt_mes&"/"&dt_ano,dia_final&"/"&dt_mes&"/"&dt_ano)
qtd_dias = 7*qtd_semanas
dif_semanas = dia_final - qtd_dias
pri_dia = 1%>
		

<br class="no_print">
<table cellspacing="1" cellpadding="1" border="0" class="no_print" style="background-color:#b8d8da; border:black solid 2px;" align="center">
	<form action="../financeiro.asp" name="seleciona_período" id="seleciona_período">
	<input type="hidden" name="tipo" value="diario3">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="data_atual" value="<%=data_atual%>">
	
	<tr>
		<td align="center" colspan="2">CONTAS A PAGAR - ver3 - <%=ordem%></td>
	</tr>
	<tr>
		<td><b>M&ecirc;s</b></td>
		<td><b>Ano</b></td>
	</tr>
	<tr>
		
		<td><select name="mes_sel">
					<option value=""></option>
					<%for imes = 1 to 12%>
						<option value="<%=imes%>" <%if abs(imes)=abs(dt_mes) then response.write("SELECTED")%>><%=mes_selecionado(imes)%></option>
					<%next%>
			</select></td>	
		<td><input type="text" name="ano_sel" maxlength="4" size="4" value="<%=dt_ano%>">&nbsp;<input type="submit" name="OK" value="OK" width="4" height="5"></td>
	</tr>
	</form>
	<tr>
		<td align="center" colspan="2"><a href="financeiro.asp?tipo=diario3&mes_sel=<%=mes_anterior%>&ano_sel=<%=ano_anterior%>"><< Anterior </a>&nbsp; &nbsp; &nbsp; <a href="financeiro.asp?tipo=diario3&mes_sel=<%=month(now)%>&ano_sel=<%=year(now)%>">Atual</a> &nbsp; &nbsp; &nbsp; <a href="financeiro.asp?tipo=diario3&mes_sel=<%=mes_posterior%>&ano_sel=<%=ano_posterior%>">Próximo>></a></td>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2"><a href="javascript:void(0);" onClick="window.open('financeiro/diario_cad3.asp?mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=inserir','Cadastro','width=500,height=480,scrollbars=1')" return false;"><b>(+)</b> Incluir</a> &nbsp;</td>
	</tr>
	</tr>
</table>
<br class="no_print">
<!-- Monta a tabela separando as semanas em blocos -->
<table cellspacing="1" cellpadding="1" border="1"  style="border-collapse:collapse; border:black solid 2px;">
	
	<tr>
		<td align="center" colspan="12" style="font-size:15px; color:red;"><b>Diário de <%=mes_selecionado(dt_mes)%> de <%=dt_ano%></b></td>
	</tr>
	
	<tr style="background-color:gray; font-size:11px; color:white;">
		<!--td align="center">&nbsp;</td>
		<td align="center"><b>Área</b></td>
		<td align="center"><b>C.C</b></td>
		<td align="center"><b>Conta</b></td>
		<td align="center"><b>Unidade</b></td>
		<td><b>Tipo</b></td>
		<td><b>Favorecido - Histórico</b></td-->
		<td colspan="8">&nbsp;</td>
		<!--td align="center" colspan="2"><b>Anterior</b></td-->
		<td align="center" colspan="4"><b>Atual</b></td>
		<!--td><b>&nbsp;</b></td-->
		<!--td><b>Pagamento</b></td-->
	</tr>
	<tr style="background-color:gray; font-size:11px; color:white;">
		<td align="center">&nbsp;</td>
		<td align="center"><b>Área</b></td>
		<td align="center"><b>C.C</b></td>
		<td align="center"><b>Conta</b></td>
		<td align="center"><b>Unidade</b></td>
		<td align="center"><b>Tipo</b></td>
		<!--td align="center"><b>Pos.</b></td-->
		<td><b>Favorecido - Histórico</b></td>
		<!--td align="center"><b>venc</b></td-->
		<td align="center"><b>Valor</b></td>
		<td align="center"><b>Venc.</b></td>	
		<td align="center"><b>Atual</b></td>
		<td align="center"><b>Extra</b></td>
		<td><b>&nbsp;</b></td>
		<!--td><b>Pagamento</b></td-->
	</tr>
	<%qtd_dias_semana = 7
	primeiro_dia_semana = 2
	cabecalho_semana = 0
	primeiro_dia = "1"'&dt_mes&"/"&dt_ano
		pri_dia = weekday("1/"&dt_mes&"/"&dt_ano)
			if pri_dia < 6 then 
				dif_ultimo_dia = 7 - pri_dia
			else
				dif_ultimo_dia = 7
			end if
			
	ultimo_dia = 1 + dif_ultimo_dia + 1
			
		'response.write(left(DiaSemanaExtenso(primeiro_dia),3))
		'response.write(primeiro_dia&" ("&dif_ultimo_dia&") ")
		'response.write(ultimo_dia&"<br>")
		
		for i_dia = 1 to qtd_dias 
			'if primeiro_dia_mes = 1 then
				'ultimo_dia_semana = weekday(i_dia&"/"&dt_mes&"/"&dt_ano)+1
				ultimo_dia_semana = 2 + qtd_dias
			'else
			'	ultimo_dia_semana = weekday("1/"&dt_mes&"/"&dt_ano)-1
			'end if
		next
		'response.write("-"&(ultimo_dia_semana))
	
	num_linha = 1
	cor = 1
	cor_linha = "#FFFFFF"
	cor_linha2 = "#e9e9e9"
	linha_separadora = 0
	nm_valor = 0
	nm_valor_extra = 0
	nm_valor_anterior = 0
	
	xsql = "up_financeiro_diario3_lista_geral @dt_i='"&dt_mes&"/1/"&dt_ano&"', @dt_f='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i_a='"&mes_anterior&"/1/"&ano_anterior&"', @dt_f_a='"&mes_anterior&"/"&dia_final_anterior&"/"&ano_anterior&"', @ordem="&ordem&""
		Set rs = dbconn.execute(xsql)
		
		while not rs.EOF
			cd_diario = rs("cd_codigo")
			dt_inicial = rs("dt_inicial")
			pre_ordem = rs("pre_ordem")
			dia_vencimento_padrao = rs("dia_vencimento_padrao")
			dt_final = rs("dt_final")
			cd_fornecedor = rs("cd_fornecedor")
			nm_fornecedor = rs("nm_fornecedor")
			nm_descricao = rs("nm_descricao")
			cd_area = rs("cd_area")
			nm_area = rs("nm_area")
			cd_centrocusto = rs("cd_centrocusto")
			nm_cc = rs("nm_cc")
			nm_centrocusto = rs("nm_centrocusto")
			cd_conta = rs("cd_conta")
			nm_conta = rs("nm_conta")
			cd_unidade = rs("cd_unidade")
			nm_unidade = rs("nm_unidade")
			nm_sigla = rs("nm_sigla")
			cd_tipo_valor = rs("cd_tipo_valor")
			nm_tipo_abrv = rs("nm_tipo_abrv")
			cd_parcela = rs("cd_parcela")
			cd_qtd_parcelas = rs("cd_qtd_parcelas")
			
			nm_pagamento = rs("nm_pagamento")
				if nm_pagamento <> "" Then
					nm_pagamento = nm_pagamento
				else
					nm_pagamento = nm_descricao
				end if
			dt_vencimento = rs("dt_vencimento")
				dt_vencimento = zero(day(dt_vencimento))'&"/"&mesdoano(month(dt_vencimento))&"/"&year(dt_vencimento)
			cd_valor = rs("cd_valor")
			nm_valor = rs("nm_valor")
				if not isnumeric(nm_valor) then
					nm_valor = "0"
				end if
				if cd_tipo_valor = 1 then
					nm_valor_extra = nm_valor
					nm_valor = "0"
					soma_semana = soma_semana + nm_valor_extra
				else
					nm_valor = nm_valor
					nm_valor_extra = "0"
					soma_semana = soma_semana + nm_valor
				end if
				
				dt_vencimento_anterior = day(rs("dt_vencimento_anterior"))'&"/"&month(rs("dt_vencimento_anterior"))
				nm_valor_anterior = rs("nm_valor_anterior")
				
					'dt_ano_anterior = dt_ano
					'dt_mes_anterior = dt_mes - 1
					'	if dt_mes_anterior < 1 then
					'		 dt_mes_anterior = 12
					'		 dt_ano_anterior = dt_ano - 1
					'	end if
					'		strsql = "Select * From TBL_financeiro_valores3 Where cd_diario='"&cd_diario&"' AND dt_vencimento between'"&month(dt_mes_anterior)&"/1/"&year(dt_ano_anterior)&"' AND '"&dt_mes_anterior&"/"&ultimoDiaMes(dt_mes_anterior,dt_ano_anterior)&"/"&(dt_ano_anterior)&"'"
					'		'strsql = "Select * From TBL_financeiro_valores3 Where cd_diario='"&cd_diario&"' "
					'		Set rs_valor = dbconn.execute(strsql)
					'			while not rs_valor.EOF 
					'				cd_diario_anterior = rs_valor("cd_diario")
					'				dt_vencimento_anterior = day(rs_valor("dt_vencimento"))
					'				nm_valor_anterior = rs_valor("nm_valor")
					'			rs_valor.movenext
					'			wend
	'if dt_vencimento <> "" and linha_separadora = 0 Then%>
		<!--tr><td colspan="12">&nbsp;</td></tr>-->
	<%if year(now)&zero(month(now)) > dt_ano&zero(dt_mes) then 'hoje maior data sel
			pos = dt_vencimento
			cabeca = dt_vencimento
		elseif year(now)&zero(month(now)) = dt_ano&zero(dt_mes) then 'hoje Igual data sel
			pos = pre_ordem
			cabeca = pre_ordem
		elseif year(now)&zero(month(now)) < dt_ano&zero(dt_mes) then ' hoje Menor data sel
			pos = dia_vencimento_padrao
			cabeca = dia_vencimento_padrao
	end if%>
	<%if int(ultimo_dia) <= int(cabeca) then
	'response.write(primeiro_dia&"~"&ultimo_dia&"<"&cabeca&":"&cabecalho_semana)
		'while not ultimo_dia
		'*** avança para a semana seguinte ***
		primeiro_dia = ultimo_dia + 1
		ultimo_dia = ultimo_dia + 7
		cabecalho_semana = 0
		
	end if%>	
	<%if cabecalho_semana = 0 AND int(cabeca) >= int(primeiro_dia) AND  int(cabeca) <= int(ultimo_dia) then
		if ultimo_dia > dia_final then ultimo_dia = dia_final%>
		<!--tr><td colspan="10">&nbsp;</td><td><%=soma_semana%></td></tr-->
			<%'if cabecalho_semana <> 0 then soma_semana = 0%>	
			<%soma_semana = 0%>		
				<tr>
					<td colspan="7" align="center" style="font-size:12px; background-color:#b8d8da;"><b>Semana do dia <%=primeiro_dia%> <%="("&left(DiaSemanaExtenso(dia_sema),3)&primeiro_dia&" - "&ultimo_dia&")"%><%'=pri_dia&" - "&dia_sema%></b></td>
					<td colspan="2" align="center" style="color:silver; font-size:10px; background-color:#b8d8da;" onClick="window.open('financeiro/diario_cad3.asp?mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=inserir','Cadastro','width=500,height=480,scrollbars=1')" return false;" style="color:silver;"><b>(+)</b> Incluir</a></td>
					<td align="right"style="font-size:10px; background-color:#b8d8da;"><%xsql = "up_financeiro_valores_soma @dt_dia_i="&primeiro_dia&", @dt_dia_f="&ultimo_dia&", @dt_mes="&dt_mes&", @dt_ano="&dt_ano&", @cd_tipo_valor=' <> 1'"
						Set rs_valores = dbconn.execute(xsql)
						if NOT rs_valores.EOF Then
							soma_semana = rs_valores("soma_semana")
							if soma_semana <> "" Then response.write("<b>"&formatNumber(soma_semana)&"</b>")
						end if%></td>
					<td align="right" style="font-size:10px; background-color:#b8d8da;"><%xsql = "up_financeiro_valores_soma @dt_dia_i="&primeiro_dia&", @dt_dia_f="&ultimo_dia&", @dt_mes="&dt_mes&", @dt_ano="&dt_ano&", @cd_tipo_valor=' = 1'"
						Set rs_valores = dbconn.execute(xsql)
						if NOT rs_valores.EOF Then
							soma_semana_extra = rs_valores("soma_semana")
							if soma_semana_extra <> "" Then response.write("<b>"&formatNumber(soma_semana_extra)&"</b>")
						end if%></td>
					<td align="right" style="font-size:10px; background-color:#b8d8da;">
						<%'total_semana = soma_semana + soma_semana_extra
						if total_semana <> "" Then response.write("<b>"&formatNumber(total_semana)&"</b>")%></td>
				</tr>
	<%cabecalho_semana = 1
	end if%>
	<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
		<td align="center">&nbsp;<a href="#<%=num_linha%>"><%=num_linha%></a></td>
		<td align="center">&nbsp;<%=nm_area%></td>
		<td align="center">&nbsp;<%=nm_cc%></td>
		<td align="center">&nbsp;<%=nm_conta%></td>
		<td align="center">&nbsp;<%=nm_sigla%></td>
		<td align="center">&nbsp;<%=nm_tipo_abrv%> <%'="*"&dia_vencimento_padrao%></td>
		<!--td align="center" style="color:gray; background-color:<%'=cor_linha2%>;">&nbsp;<%'if cabeca <> "" then 
									''response.write(zero(pos))
									'response.write(cabeca)
								'end if%></td-->
		<td onClick="window.open('financeiro/diario_lista_pag3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','Visualizar','width=500,height=<%if cd_tipo_valor=1 then%>240<%elseif cd_tipo_valor=2 then%>480<%elseif cd_tipo_valor=3 then%>440<%end if%>,scrollbars=1')" return false;">&nbsp;
									<%'=cd_fornecedor&" * "%><%if cd_fornecedor > 0 then response.write("<b>"&nm_fornecedor&"</b>-")%><%'=nm_descricao%><%=nm_pagamento%>&nbsp;
									<%if cd_qtd_parcelas > 1 then
										response.write("<span style='color:red;'><b>"&cd_parcela&"/"&cd_qtd_parcelas&"</b></span>")
									end if%><%'=primeiro_dia&"("&cabeca&")"&ultimo_dia&" : "&cabecalho_semana%></td>
		<!--td align="center" style="color:gray; background-color:<%'=cor_linha2%>;">&nbsp;<%'if dt_vencimento_anterior <> "" then 
									'response.write(zero(dt_vencimento_anterior))
								'end if%></td-->
		<td align="right" style="color:gray; background-color:<%=cor_linha2%>;">&nbsp;<%if nm_valor_anterior <> "0" then 
									response.write(FormatNumber(nm_valor_anterior,2))
								end if%></td>
		<td align="center">&nbsp;<%=dt_vencimento%></td>
		<td align="right" <%if cd_tipo_valor <> 1 Then%>onClick="window.open('financeiro/diario_cad3.asp?cd_diario=<%=cd_diario%>&dia_sel=<%=dt_vencimento_anterior%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=<%if cd_valor <> "" Then%>editar_valor<%else%>inserir_valor<%end if%>','Cadastro','width=500,height=250,scrollbars=1')" return false;<%end if%>><%if nm_valor <> "0" then 
									response.write(FormatNumber(nm_valor,2))
									'response.write(nm_valor)
								end if%></td>
		<td align="right" <%if cd_tipo_valor = 1 Then%>onClick="window.open('financeiro/diario_cad3.asp?cd_diario=<%=cd_diario%>&dia_sel=<%=dt_vencimento_anterior%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=<%if nm_valor_extra = "0" Then%>inserir_valor<%else%>editar_valor<%end if%>','Cadastro','width=500,height=250,scrollbars=1')" return false;<%end if%>><%if nm_valor_extra > "0" then 
									response.write(FormatNumber(nm_valor_extra,2))
								end if%></td>
		<td>&nbsp;<img src="../imagens/ic_editar.gif" alt="" width="13" height="14" border="0" onClick="window.open('financeiro/diario_cad3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=editar','Cadastro','width=500,height=480,scrollbars=1')" return false;"></td-->
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr>	
	<%dt_vencimento_anterior = "0"
	nm_valor_anterior = 0
	soma_valores = abs(soma_valores) + nm_valor
	soma_extra = abs(soma_extra) + nm_valor_extra
	num_linha = num_linha + 1
	
	if cor > 0 then
		cor_linha = "#d7d7d7"
		cor_linha2 = "#d7d7d7"
		cor = 0
	else
		cor_linha = "#FFFFFF"
		cor_linha2 = "#e9e9e9"
		cor = 1
	end if
	
		rs.movenext
		wend%>
	<tr><td colspan="12">&nbsp;</td></tr>
	<tr>
		<td><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
		<!--td><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td-->
		<td><img src="../imagens/px.gif" alt="" width="400" height="1" border="0"></td>
		<!--td colspan="2"><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td-->
		
		<td>&nbsp;<%%></td>
		<td>&nbsp;<%%></td>
		<td align="right"><b><%if soma_valores <> "" Then response.write(formatNumber(soma_valores,2))%></b></td>
		<td align="right"><b><%if soma_extra <> "" Then response.write(formatNumber(soma_extra,2))%></b></td>
		<td>&nbsp;<%%></td>
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr>
	<%total_valores = soma_valores + soma_extra%>
	<tr>
		<td colspan="9">&nbsp;</td>
		<!--td>&nbsp;<%%>-</td-->
		<!--td>&nbsp;<%%></td-->
		<td align="center" colspan="3">&nbsp;<b><%if total_valores <> "" Then response.write(FormatNumber(total_valores,2))%></b></td>
		<!--td>&nbsp;<%%></td-->
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr>
	<%total = abs(total) + nm_valor%>	
</table>
<br class="no_print">
<table cellspacing="1" cellpadding="1" border="1" class="no_print" style="border-collapse:collapse; border:black solid 2px;" align="center">
	<tr style="background-color:#b8d8da;">
		<td>&nbsp;</td>
		<td><b>Receita</b></td>
		<td><b>N° de cirurgias</b></td>
		<td><b>Liquido</b></td>
		<td><b>Data pag.</b></td>
		<td><b>Recebido</b></td>		
	</tr>
	<%
	mes_faturamento = dt_mes - 2
		ano_faturamento = dt_ano
			if mes_faturamento < 1 then
				if mes_faturamento = 0 then
					mes_faturamento = 12
				elseif mes_faturamento < 0 then
					mes_faturamento = 11
				end if
				ano_faturamento = dt_ano - 1
			end if
			dia_final_faturamento = ultimodiames(mes_faturamento,ano_faturamento)
	
	num_linha = 1
	strsql = "SELECT COUNT (a002_numpro) AS conta , dt_ano, dt_mes,a053_codung, nm_sigla, nm_unidade, dia_recebimento FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_faturamento&"/1/"&ano_faturamento&"' AND '"&mes_faturamento&"/"&dia_final_faturamento&"/"&ano_faturamento&"' AND a002_tipage <> 'E' Group by dt_ano, dt_mes,a053_codung, nm_sigla, nm_unidade, dia_recebimento order by conta desc"
	Set rs = dbconn.execute(strsql)
	
	while not rs.EOF
	conta = rs("conta")
	dt_ano = rs("dt_ano")
	dt_mes = rs("dt_mes")
	dia_recebimento = rs("dia_recebimento")
	a053_codung = rs("a053_codung")
	nm_sigla = rs("nm_sigla")
	nm_unidade = rs("nm_unidade")
		strsql = "SELECT * FROM TBL_financeiro_diario_receita Where cd_unidade = '"&a053_codung&"' AND dt_faturamento <= '"&mes_faturamento&"/"&dia_final_faturamento&"/"&ano_faturamento&"' order by dt_faturamento desc"
		Set rs_receita = dbconn.execute(strsql)
		if not rs_receita.EOF then
			valor_cirurgia = rs_receita("valor_cirurgia")
			qtd_cirurgia_manual = rs_receita("qtd_cirurgia_manual")
			valor_final_manual = rs_receita("valor_final_manual")
			
		rs_receita.movenext
		end if
		
		if valor_cirurgia <> "" then
			valor_final = formatnumber(conta * valor_cirurgia)
		end if
		if valor_cirurgia <> "" AND qtd_cirurgia_manual = "" AND valor_final_manual = "" then
		'	valor_final = formatnumber(conta * valor_cirurgia)
		'	m = 1
		'elseif  valor_cirurgia <> "" AND qtd_cirurgia_manual <>  "" AND valor_final_manual = "" then
		'	valor_final = formatnumber(qtd_cirurgia_manual * valor_cirurgia)
		'	m = 2
		'elseif  valor_cirurgia <> "" AND qtd_cirurgia_manual <>  "" AND valor_final_manual <> "" then
		'	valor_final = valor_final_manual
		'	m = 3
		'elseif  valor_cirurgia <> "" AND qtd_cirurgia_manual =  "" AND valor_final_manual <> "" then
		'	valor_final = valor_final_manual
		'	m = 4
		'else
		'	response.write(valor_cirurgia&"/"&qtd_cirurgia_manual&"/"&valor_final_manual)
		'	m = 5
		end if
		%>
		<tr>
			<td><%=num_linha%></td>
			<td><%=nm_unidade%></td>
			<td align="right"><%=conta%></td>
			<td align="right">&nbsp;<%=valor_final%> <%=m%></td>
			<td align="right"><%=dia_recebimento%></td>
			<td align="right">&nbsp;</td>
		</tr>
	<%num_linha = num_linha + 1
	
	valor_cirurgia = ""
	valor_final = ""
	rs.movenext
	wend%>
</table>