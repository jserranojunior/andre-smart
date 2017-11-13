
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
	end if%>
<script language="javascript">
	function adiciona_pgto(a,lista_pgto,check1,check2){
		 lista_pgto = lista_pgto+','+a
		document.diario.total_pgto.value=(lista_pgto);
		var el1=document.getElementById(check1);
			el1.style.display=(el1.style.display!='none'?'none':'');
		var el2=document.getElementById(check2);
			el2.style.display=(el2.style.display!='none'?'none':'');		
		}
	
	function subtrai_pgto(b,lista_pgto,check2,check1){
		lista_pgto = lista_pgto.replace(','+b,'');
		document.diario.total_pgto.value=(lista_pgto);
		var el2=document.getElementById(check2);
			el2.style.display=(el2.style.display!='none'?'none':'');
		var el1=document.getElementById(check1);
			el1.style.display=(el1.style.display!='none'?'none':'');
		}
	
	function monta_pgto(lista_pgto){
		mywindow = window.open("financeiro/diario_pagamentos3.asp?lista_pagto="+lista_pgto+"&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>", "Emissao_Cheques", "location=0,status=0,width=700,height=400,scrollbars=1");  
		mywindow.moveTo(0, 0); 
}	

</script>		

<br class="no_print">
<table align="center" style="border:white solid 2px; border-collapse:collapse;">
	<tr><td valign="top">
<table cellspacing="1" cellpadding="1" border="0" class="no_print" style="background-color:#b8d8da; border:black solid 2px;" align="center">
	<form action="../financeiro.asp" name="seleciona_periodo" id="seleciona_período">
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
</table>
<br class="no_print">
<!-- Monta a tabela separando as semanas em blocos -->
	</td>
	<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
	<td valign="top">
		<table cellspacing="1" cellpadding="1" border="0" class="no_print" style="background-color:#b8d8da; border:black solid 2px;" align="center">
			<tr>
				<td align="center" colspan="5" style="color:white;"><b>LEGENDA</b></td>
			</tr>
			<tr>
				<%cor_50pos = "red"
				cor_50neg = "#0057ae"%>
				<td bgcolor="<%=cor_50pos%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>> 50% <</td>
				<td bgcolor="<%=cor_50neg%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>&nbsp;</td>
				<td><img src="../imagens/check_ok.gif" alt="" width="25" height="12" border="0"></td>
				<td>confirmado</td>
			</tr>
			<tr>
				<%cor_25pos = "orange"
				cor_25neg = "#0080ff"%>
				<td bgcolor="<%=cor_25pos%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>> 25% <</td>
				<td bgcolor="<%=cor_25neg%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>&nbsp;</td>
				<td><img src="../imagens/check_aguarda.gif" alt="" width="25" height="12" border="0"></td>
				<td>cheque emitido</td>
			</tr>
			<tr>
				<%cor_10pos = "yellow"
				cor_10neg = "#5badff"%>
				<td bgcolor="<%=cor_10pos%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>> 10% <</td>
				<td bgcolor="<%=cor_10neg%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>&nbsp;</td>
				<td><img src="../imagens/check_ativo.gif" alt="" width="25" height="12" border="0"></td>
				<td>selecionado</td>
			</tr>
			<tr>
				<%cor_5pos = "green"
				cor_5neg = "#a6d2ff"%>
				<td bgcolor="<%=cor_5pos%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>> 5% <</td>
				<td bgcolor="<%=cor_5neg%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>&nbsp;</td>
				<td><img src="../imagens/check_inativo.gif" alt="" width="25" height="12" border="0"></td>
				<td>pronto</td>
			</tr>
			<tr>
				<%cor_0pos = "white"
				cor_0neg = "white"%>
				<td bgcolor="<%=cor_0pos%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>>= 0% <=</td>
				<td bgcolor="<%=cor_0neg%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
				<td>&nbsp;</td>
				<td align="center"><img src="../imagens/check_proibido.gif" alt="" width="12" height="12" border="0"></td>
				<td>não há valor</td>
			</tr>
		</table>
	</td></tr>
</table>
<table style="border-collapse:collapse; border:black solid 1px;" cellspacing="1">
	<form action="../financeiro.asp" name="diario" id="diario" method="post">
	<input type="hidden" name="total_pgto" value="">
	<tr><td align="center" colspan="13"><input type="button" name="pagtos" value="Emitir cheque" onclick="monta_pgto(document.diario.total_pgto.value);"></td></tr>
	<tr>
		<td align="center" colspan="13" style="font-size:15px; color:red;"><b>Diário de <%=mes_selecionado(dt_mes)%> de <%=dt_ano%></b></td>
	</tr>
	
	<tr style="background-color:gray; font-size:11px; color:white;">
		<!--td align="center">&nbsp;</td>
		<td align="center"><b>Área</b></td>
		<td align="center"><b>C.C</b></td>
		<td align="center"><b>Conta</b></td>
		<td align="center"><b>Unidade</b></td>
		<td><b>Tipo</b></td>
		<td><b>Favorecido - Histórico</b></td-->
		<td colspan="9">&nbsp;</td>
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
		<td align="center"><b>U.N.</b></td>
		<td align="center"><b>Tipo</b></td>
		<td align="center"><b>Pg</b></td>
		<td colspan="2"><b>Favorecido - Histórico</b></td>
		<!--td align="center"><b>venc</b></td-->
		<!--td align="center"><b>&nbsp;</b></td-->
		<td align="center"><b>Valor</b></td>
		<td align="center"><b>Venc.</b></td>	
		<td align="center"><b>Atual</b></td>
		<td align="center"><b>Extra</b></td>
		<!--td><b>Pagamento</b></td-->
	</tr>
	<%'qtd_dias_semana = 7
	'primeiro_dia_semana = 1
	cabecalho_semana = 0
	'*** Primeira verificação dos dias da semana
	primeiro_dia = "1"'&dt_mes&"/"&dt_ano
		pri_dia = weekday("1/"&dt_mes&"/"&dt_ano)
	'		if pri_dia < 6 then 
				dif_ultimo_dia = 7 - pri_dia
	''		else
	'			dif_ultimo_dia = 6
	'		end if
	'		
	ultimo_dia = dif_ultimo_dia + 1
	'********************************************
		
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
			nm_obs = rs("nm_obs")
			
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
					'if nm_valor_anterior > 0 OR  nm_valor > 0 Then
						'if cd_valor <> "" Then
							
							'status_variacao = instr(1,diferenca_valores,"-",1)
							'variacao_valores = formatnumber((diferenca_valores/nm_valor_anterior)*100,2)
							if nm_valor <> 0 AND nm_valor_anterior <> 0 AND cd_valor <> 0 then
								diferenca_valores = nm_valor - nm_valor_anterior
								'variacao_valores = formatnumber(diferenca_valores,2)
								'variacao_valores = diferenca_valores&":"&formatnumber((diferenca_valores/nm_valor_anterior)*100,2)
								variacao_valores = formatnumber((diferenca_valores/nm_valor_anterior)*100,2)
							end if
						'end if
				cd_conta_banco = rs("cd_conta_banco")
				nm_cheque = rs("nm_cheque")
				cd_confirma_pagto = rs("cd_confirma_pagto")
				cd_pagto_agendado = rs("cd_pagto_agendado")
				dt_pagto = rs("dt_pagto")
				
					
						
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
			
	end if
		
		if cabeca > dia_final then
			cabeca = dia_final
		end if
		
		'*** segunda verificação dos dias da semana ***
		dia_semana = weekday(cabeca&"/"&dt_mes&"/"&dt_ano)
		ultimo_dia = 7 - dia_semana%>
		
	<%'if int(dia_semana) > then
	'response.write(primeiro_dia&"~"&ultimo_dia&"<"&cabeca&":"&cabecalho_semana)
		'while not ultimo_dia
		'*** avança para a semana seguinte ***
	'	primeiro_dia = ultimo_dia + 1
	'	ultimo_dia = ultimo_dia + 7
	'	cabecalho_semana = 0
		
	'end if%>	
	<%'if cabecalho_semana = 0 AND int(cabeca) >= int(primeiro_dia) AND  int(cabeca) <= int(ultimo_dia) 
	if cabecalho_semana = a then
		if ultimo_dia > dia_final then ultimo_dia = dia_final%>
		<!--tr><td colspan="10">&nbsp;</td><td><%=soma_semana%></td></tr-->
			<%'if cabecalho_semana <> 0 then soma_semana = 0%>	
			<%soma_semana = 0%>		
				<tr>
					<td colspan="8" align="center" style="font-size:12px; background-color:#b8d8da;"><%=pri_dia&"->"&dia_semana&"("&ultimo_dia%>) <b>Semana do dia <%=primeiro_dia%> <%'="("&left(DiaSemanaExtenso(dia_sema),3)&primeiro_dia&" - "&ultimo_dia&")"%><%'=pri_dia&" - "&dia_sema%></b></td>
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
						<%'total_soma_semana = soma_semana + soma_semana_extra
						'if total_soma_semana <> "" Then response.write("<b>"&formatNumber(total_soma_semana)&"</b>")%></td>
				</tr>
	<%total_soma_semana = 0
	cabecalho_semana = 1
	end if%>
	<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
		<td align="center"><a href="#<%=num_linha%>"><%=num_linha%></a></td>
		<td align="center"><%=nm_area%></td>
		<td align="center"><%=nm_cc%></td>
		<td align="center"><%=nm_conta%></td>
		<td align="center"><%=nm_sigla%></td>
		<td align="center"><%=nm_tipo_abrv%> <%'="*"&dia_vencimento_padrao%></td>
		<!--td align="center" style="color:gray; background-color:<%'=cor_linha2%>;">&nbsp;<%'if cabeca <> "" then 
									''response.write(zero(pos))
									'response.write(cabeca)
								'end if%></td-->
		<td align="center"><%check1 = "check_1_"&cd_valor
		check2 = "check_2_"&cd_valor
		if not IsNumeric(cd_valor) Then%>
			<img src="../../imagens/check_proibido.gif" alt="Inclua o valor" width="12" height="12" border="0" onClick="window.open('financeiro/diario_cad3.asp?cd_diario=<%=cd_diario%>&dia_sel=<%=dt_vencimento_anterior%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=inserir_valor','Cadastro','width=500,height=250,scrollbars=1')" return false;>
		<%elseif cd_confirma_pagto = int(1) AND cd_pagto_agendado = int(1) then%>
			<img src="../../imagens/check_ok.gif" alt="Pago" width="25" height="12" border="0">
		<%elseif cd_pagto_agendado = int(1) then%>
			<img src="../../imagens/check_aguarda.gif" alt="lista cheque(<%=nm_cheque%>)" width="25" height="12" border="0" onClick="window.open('financeiro/diario_pagamentos3.asp?cd_conta_banco=<%=cd_conta_banco%>&nm_cheque=<%=nm_cheque%>','lista_cheque','width=700,height=500,scrollbars=1')"></td>
		<%else%>	
			<img src="../../imagens/check_inativo.gif" alt="Inclui" width="25" height="12" border="0" onClick="adiciona_pgto('<%=cd_valor%>',document.diario.total_pgto.value,'<%=check1%>','<%=check2%>')" id="<%=check1%>" name="<%=check1%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>>
			<img src="../../imagens/check_ativo.gif" alt="Exclui" width="25" height="12" border="0" onClick="subtrai_pgto('<%=cd_valor%>',document.diario.total_pgto.value,'<%=check2%>','<%=check1%>')" id="<%=check2%>" name="<%=check2%>" <%'if ordem_var <> "" Then%>style="display:none;"<%'end if%>></td>
		<%end if%>
		<td><span onClick="window.open('financeiro/diario_lista_pag3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','Visualizar','width=500,height=<%if cd_tipo_valor=1 then%>240<%elseif cd_tipo_valor=2 then%>480<%elseif cd_tipo_valor=3 then%>460<%end if%>,scrollbars=1')" return false;" style="border-style:none;">
					<%if cd_fornecedor > 0 then response.write("<b>"&nm_fornecedor&"</b>")%>
					<%if nm_fornecedor <> "" and nm_pagamento <> "" Then response.write(" - ")%>
						<%=nm_pagamento%>&nbsp;</span>						
							<%if cd_qtd_parcelas > 1 then
								response.write("<span style='color:red;'><b>"&cd_parcela&"/"&cd_qtd_parcelas&"</b></span>")
							end if%>
							<%if nm_obs <> "" then%><img src="../imagens/ic_obs3.gif" alt="" width="10" height="10" border="0" onClick="window.open('financeiro/diario_lista_obs3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','Visualizar','width=500,height=260,scrollbars=1')" return false;"><%end if%>		</td>
		<!--td align="center" style="color:gray; background-color:<%'=cor_linha2%>;">&nbsp;<%'if dt_vencimento_anterior <> "" then 
									'response.write(zero(dt_vencimento_anterior))
								'end if%></td-->
		
		<td><img src="../imagens/ic_editar.gif" alt="Editar" width="13" height="14" border="0" onClick="window.open('financeiro/diario_cad3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=editar','Cadastro','width=500,height=480,scrollbars=1')" return false;"></td>
		<td align="right" style="color:gray; background-color:<%=cor_linha2%>;">&nbsp;<%if nm_valor_anterior <> "0" then 
									response.write(FormatNumber(nm_valor_anterior,2))
								end if%></td>
		<td align="center"><%=dt_vencimento%></td>
		<td align="right" <%if cd_tipo_valor <> 1 Then%>onClick="window.open('financeiro/diario_cad3.asp?cd_diario=<%=cd_diario%>&dia_sel=<%=dt_vencimento_anterior%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=<%if cd_valor <> "" Then%>editar_valor<%else%>inserir_valor<%end if%>','Cadastro','width=500,height=250,scrollbars=1')" return false;<%end if%>><%if nm_valor <> "0" then 
									response.write(FormatNumber(nm_valor,2))
									'response.write(nm_valor)
								end if%></td>
		<td align="right" <%if cd_tipo_valor = 1 Then%>onClick="window.open('financeiro/diario_cad3.asp?cd_diario=<%=cd_diario%>&dia_sel=<%=dt_vencimento_anterior%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=<%if nm_valor_extra = "0" Then%>inserir_valor<%else%>editar_valor<%end if%>','Cadastro','width=500,height=250,scrollbars=1')" return false;<%end if%>><%if nm_valor_extra <> "0" then 
									response.write(FormatNumber(nm_valor_extra,2))
								end if%></td>
		<%'if nm_valor < nm_valor_anterior then
		'if 	cor_variacao = "green"
			'seta = "ic_indicador_down.gif"
			'indicador = "<img src='../imagens/"&seta&"' alt='"&variacao_valores&"%' width='15' height='13' border='0'>"
		'elseif nm_valor = nm_valor_anterior then
			'cor_variacao = "gray"
			'seta = "ic_indicador_ok.gif"
			'indicador = "<img src='../imagens/"&seta&"' alt='"&variacao_valores&"%' width='12' height='12' border='0'>"
		'elseif nm_valor > nm_valor_anterior then
			'cor_variacao = "red"
			'seta = "ic_indicador_up.gif"
			'indicador = "<img src='../imagens/"&seta&"' alt='"&variacao_valores&"%' width='15' height='13' border='0'>"
		'end if
		if variacao_valores = "" then
			seta = "px.gif"
			cor_variacao = ""'cor_0pos
		elseif int(variacao_valores) > 50 then
			seta = "ic_indicador_up.gif"
			cor_variacao = cor_50pos
		elseif int(variacao_valores) > 25 then
			seta = "ic_indicador_up.gif"
			cor_variacao = cor_25pos
		elseif int(variacao_valores) > 10 then
			seta = "ic_indicador_up.gif"
			cor_variacao = cor_10pos
		elseif int(variacao_valores) >= 5 then
			seta = "ic_indicador_up.gif"
			cor_variacao = cor_5pos
		elseif int(variacao_valores) < -50 then
			seta = "ic_indicador_down.gif"
			cor_variacao = cor_50neg
		elseif int(variacao_valores) < -25 then
			seta = "ic_indicador_down.gif"
			cor_variacao = cor_25neg
		elseif int(variacao_valores) < -10 then
			seta = "ic_indicador_down.gif"
			cor_variacao = cor_10neg
		elseif int(variacao_valores) <= -5 then
			seta = "ic_indicador_down.gif"
			cor_variacao = cor_5neg
		
		
		
		'elseif variacao_valores > -5 and variacao_valores < 5 then
		'	seta = "ic_indicador_ok.gif"
		end if%>
		<td align="center" bgcolor="<%=cor_variacao%>">&nbsp;<img src="../imagens/px.gif<%'=seta%>" alt="<%=variacao_valores%>%" width="15" height="13" border="0"></td>
	</tr>	
	<%
	
	soma_valores = abs(soma_valores) + nm_valor
	soma_extra = abs(soma_extra) + nm_valor_extra
	num_linha = num_linha + 1
	dt_vencimento_anterior = "0"
	nm_valor_anterior = 0
	nm_valor = 0
	nm_fornecedor = ""
	nm_pagamento = ""
	cd_valor = ""
	diferenca_valores = 0
	variacao_valores = ""
	seta = "px.gif"
	cor_variacao = ""
	
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
		<td colspan="11"><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
		<!--td><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="400" height="1" border="0"></td-->
		<!--td colspan="2"><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td-->
		
		<!--td>&nbsp;<%%></td>
		<td>&nbsp;<%%></td-->
		<td align="right"><b><%if soma_valores <> "" Then response.write(formatNumber(soma_valores,2))%></b></td>
		<td align="right"><b><%if soma_extra <> "" Then response.write(formatNumber(soma_extra,2))%></b></td>
		<td>&nbsp;<%%></td>
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr>
	<%total_valores = soma_valores + soma_extra%>
	<tr>
		<td colspan="11">&nbsp;</td>
		<!--td>&nbsp;<%%>-</td-->
		<!--td>&nbsp;<%%></td-->
		<td align="center" colspan="3" bgcolor="Silver" style="font-size:15px;">&nbsp;<b><%if total_valores <> "" Then response.write(FormatNumber(total_valores,2))%></b></td>
		<!--td>&nbsp;<%%></td-->
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr>
	<%total = abs(total) + nm_valor%>
	<tr>
		<td><img src="../imagens/blackdot.gif" alt="" width="25" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="75" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="30" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="370" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="15" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="60" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>	
	</tr>
	</form>

</table>