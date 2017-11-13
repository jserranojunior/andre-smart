
<%dt_ano = request("ano_sel")
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

'*** Filtros ***
subtotal = request("subtotal")
str_area = request("str_area")
	if str_area  = "" Then str_area = 0


cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)
filtro_ordens = request("filtro_ordens")
	if filtro_ordens = "" Then filtro_ordens = 0

	'if  year(now)&zero(month(now))&zero(day(now)) > dt_ano&zero(dt_mes)&dia_final then
	if  year(now)&zero(month(now)) > dt_ano&zero(dt_mes) then
		ordem = " dt_vencimento "'
	elseif year(now)&zero(month(now)) = dt_ano&zero(dt_mes) then
		ordem = " pre_ordem "
	elseif year(now)&zero(month(now)) < dt_ano&zero(dt_mes) then
		'ordem = " dia_vencimento_padrao "
		ordem = " pre_ordem "
	end if
	
		if filtro_ordens = "1" Then
			ordem = " nm_fornecedor"'&ordem
			subtotal = 0
		elseif filtro_ordens = "2" Then
			ordem = " nm_tipo_pag_abrv"'&ordem
			subtotal = 1
		else
			subtotal = 0
		end if

	%>
<script language="javascript">
	function manipulacao_valor(cd_diario,dt_vencimento_anterior,dt_mes,dt_ano,acao,user){
		shipinfo = document.getElementById('seleciona_periodo');
		hoje = new Date
		mes_hoje = hoje.getMonth()+1;
		ano_hoje = hoje.getFullYear();
			if (mes_hoje < 10 ) {
				mes_hoje = '0'+ mes_hoje;}
				//ano_hoje = hoje.getFullYear();}				
			data_hoje = ano_hoje +''+ (mes_hoje);
		 	data_informada = dt_ano +''+ dt_mes;
			/*if (data_informada<data_hoje){
				window.alert ("O mês selecionado já está fechado.");  return (false);}
			else{	*/
				window.open('financeiro/diario_cad3.asp?cd_diario='+cd_diario+'&dia_sel='+dt_vencimento_anterior+'&mes_sel='+dt_mes+'&ano_sel='+dt_ano+'&acao='+acao+'&orient=valor','Cadastro','width=500,height=300,scrollbars=1');
				/*}*/
	}
	
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
		mywindow = window.open("financeiro/diario_pagamentos3.asp?lista_pagto="+lista_pgto+"&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>", "Emissao_Cheques", "location=0,status=1,width=700,height=400,scrollbars=1, resizable=1");  
		mywindow.moveTo(0, 0);
		}
		
	function filtro(){
		str_area = document.diario.str_area.value;
		//alert(str_area);
		document.forms["diario"].submit();
		}
		
	function filtro_ordem(){
		str_area = document.diario.filtro_ordens.value;
		//alert("teste");
		document.forms["diario"].submit();
		}
	
	


	
</script>
<!--br class="no_print"-->
<table>
	<tr class="no_print">
		<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
		<td>
			<a href="javascript:void(0);" return false;" onClick="window.open('financeiro/diario_emissao_cheques3.asp?tipo=emissao_cheques&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','cheques_e','location=0,status=0,width=620,height=600,scrollbars=1')">Cheques emitidos</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('financeiro/diario_emissao_cheques3.asp?tipo=transferencias&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','Transf','location=0,status=0,width=620,height=600,scrollbars=1')">Transfer&ecirc;ncias</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('financeiro/fatura_consolidado3.asp','','location=0,status=0,width=530,height=665,scrollbars=1,resizable=1')">Fatura Consolidado</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('financeiro/fatura3.asp','','location=0,status=0,width=820,height=700,scrollbars=1,resizable=1')">Emissão de faturas</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('financeiro/diario_relat_areas3.asp','','location=0,status=0,width=820,height=700,scrollbars=1,resizable=1')">Relatório de gestão</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('../manutencao_2/cad_orc_aprov.asp?tipo=financ','','location=0,status=0,width=800,height=700,scrollbars=1,resizable=1')">Orçamentos aprovados</a>&nbsp;-&nbsp; 
				<%strsql = "SELECT     COUNT(nr_orcamento) AS qtd_orc FROM TBL_manutencao_orcamento WHERE cd_pagto IS NULL"
				Set rs = dbconn.execute(strsql)
				if not rs.EOF then
					response.write("<i style='color:red;'><b>"&rs("qtd_orc")&"</b></i>")
				else
					response.write("<i>0</i>")
				end if%><br />
		</td>
		<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
		<td>
			<table border="0" cellspacing="1" cellpadding="1" class="no_print" style="background-color:#b8d8da; border:black solid 2px;" align="center">
				<form action="../financeiro.asp" name="seleciona_periodo" id="seleciona_período">
				<input type="hidden" name="tipo" value="diario3">
				<input type="hidden" name="cd_user" value="<%=cd_user%>">
				<input type="hidden" name="data_atual" value="<%=data_atual%>">
				<input type="hidden" name="str_area" value="<%if str_area <> "0" Then%><%=str_area%><%else%>0<%end if%>">
				<input type="hidden" name="filtro_ordens" value="<%=filtro_ordens%>">
				<input type="hidden" name="subtotal" value="<%=subtotal%>">
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
					<td align="center" colspan="2"><a href="financeiro.asp?tipo=diario3&mes_sel=<%=mes_anterior%>&ano_sel=<%=ano_anterior%>&str_area=<%=str_area%>&filtro_ordens=<%=filtro_ordens%>"><< Anterior </a>&nbsp; &nbsp; &nbsp; <a href="financeiro.asp?tipo=diario3&mes_sel=<%=month(now)%>&ano_sel=<%=year(now)%>&str_area=<%=str_area%>&filtro_ordens=<%=filtro_ordens%>">Atual</a> &nbsp; &nbsp; &nbsp; <a href="financeiro.asp?tipo=diario3&mes_sel=<%=mes_posterior%>&ano_sel=<%=ano_posterior%>&str_area=<%=str_area%>&filtro_ordens=<%=filtro_ordens%>">Próximo>></a></td>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2"><%'if dt_ano&zero(dt_mes) >= year(now)&zero(month(now)) then%><a href="javascript:void(0);" onClick="window.open('financeiro/diario_cad3.asp?mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=inserir','Cadastro','width=500,height=480,scrollbars=1,status=1')" return false;"><b>(+)</b> Incluir</a><%'end if%> &nbsp;</td>
				</tr>	
			</table>
		</td>
		<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
		<td>
			<table cellspacing="1" cellpadding="1" border="0" class="no_print" style="background-color:#b8d8da; border:black solid 2px;" align="center">
				<tr>
					<td align="center" colspan="5" style="color:white;"><b>LEGENDA</b></td>
				</tr>
				<tr>
					<%cor_pos = "red"
					'cor_50neg = "#0057ae"%>
					<td align="center"><img src="../imagens/pagto-cheque.gif" alt="" width="12" height="12" border="0"></td>
					<td>cheque emitido</td>
					<td>&nbsp;</td>
					<td bgcolor="<%=cor_pos%>"><img src="../imagens/px.gif" alt="" width="10" height="10" border="0"></td>
					<td>> &nbsp;7%</td>
				</tr>
				<tr>
					<%'cor_0 = "#f2f88f"
					cor_neg = "#0057ae"%>
					
					<td align="center"><img src="../imagens/pagto-transf.gif" alt="" width="12" height="12" border="0"></td>
					<td>transferência</td>
					<td>&nbsp;</td>
					<td bgcolor="<%=cor_neg%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
					<!--td bgcolor="<%'=cor_0%>"><img src="../imagens/px.gif" alt="" width="10" height="10" border="0"></td-->
					<td>< -7%</td>
					<!--td>+/-7%</td-->
					
				</tr>
				<tr>
					<%'cor_neg = "#0057ae"%>
					<td><img src="../imagens/check_ativo.gif" alt="" width="15" height="12" border="0"></td>
					<td>selecionado</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<!--td bgcolor="<%=cor_neg%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td-->
					<!--td>< -7%</td-->
				</tr>
				<tr>
					<%'cor_5neg = "5badff"
					'cor_5neg = "#a6d2ff"%>
					<td align="center"><img src="../imagens/pagto-inativo.gif" alt="" width="12" height="12" border="0"></td>
					<td> &nbsp; -</td>
					<td>&nbsp;</td>
					<td bgcolor="<%'=cor_5neg%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<%'cor_10neg = "0057ae"
					'cor_0neg = "white"%>
					<td align="center"><img src="../imagens/check_proibido.gif" alt="" width="12" height="12" border="0"></td>
					<td>não há valor</td>
					<td>&nbsp;</td>
					<td bgcolor="<%'=cor_10neg%>"><img src="../imagens/px.gif" alt="" width="12" height="12" border="0"></td>
					<td>&nbsp;</td>
				</tr>
			</table>
		</td>
		<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
	</tr>
</table>
<br class="no_print">
<table align="left" style="border-collapse:collapse; border:black solid 1px; height:20px; overflow:scroll;" cellspacing="1" >
	<form action="../financeiro.asp" name="diario" id="diario" method="post">
	<input type="hidden" name="total_pgto" value="">
	<input type="hidden" name="tipo" value="diario3">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="data_atual" value="<%=data_atual%>">
	<input type="hidden" name="mes_sel" value="<%=dt_mes%>">
	<input type="hidden" name="ano_sel" value="<%=dt_ano%>">
	<tr class="no_print">
		<td align="center" colspan="13"><input type="button" name="pagtos" value="Emitir cheques / transferências" onclick="monta_pgto(document.diario.total_pgto.value);" class="no_print"></td>
	</tr>
	<tr>
		<td align="center" colspan="13" style="font-size:15px; color:red;"><b>Diário de <%=mes_selecionado(dt_mes)%> de <%=dt_ano%></b></td>
	</tr>
	
	<tr class="no_print">
		<td colspan="1">&nbsp;</td>
		<td>
			<%xsql = "SELECT * FROM TBL_financeiro_area"
				Set rs = dbconn.execute(xsql)%>
				<select name="str_area" onChange="filtro()"; return false;>
					<option value="0">todos</option>					
					<%while not rs.EOF
						strcd_area = rs("cd_codigo")
						strnm_area = rs("nm_area")%>
						<option value="<%=strcd_area%>" <%if abs(strcd_area) = abs(str_area) Then response.write("SELECTED")%>><%=strnm_area%></option>
					<%rs.movenext
					wend%>
				</td>
				<td align="right" colspan="4"><%if str_area > 0 Then response.write("<b style='color:red;'>*** Filtro Ativado ***</b>")%></td>
				<td>&nbsp;</td>
				<td>ORDEM: <select name="filtro_ordens" onChange="filtro_ordem()"; return false;>
							<option value="0" <%if filtro_ordens = 0 then response.write("SELECTED")%>>&nbsp;</option>
							<option value="1" <%if filtro_ordens = 1 then response.write("SELECTED")%>>Alfabética</option>
							<option value="2" <%if filtro_ordens = 2 then response.write("SELECTED")%>>Tipo</option-->
						</select></td>
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
		<td align="center"><b>Anterior</b></td>
		<td align="center"><b>Venc.</b></td>	
		<td align="center"><b>Atual</b></td>
		<td align="center"><b>Extra</b></td>
		<!--td align="center"><b>Pago</b></td-->
		<td><b>&nbsp;</b></td-->
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
			nm_conta_abrv = rs("nm_conta_abrv")
			cd_unidade = rs("cd_unidade")
			nm_unidade = rs("nm_unidade")
			nm_sigla = rs("nm_sigla")
			cd_tipo_valor = rs("cd_tipo_valor")
			'cd_tipo_valor = rs("nm_tipo_pagamento")
			nm_tipo_abrv = rs("nm_tipo_pag_abrv")
			cd_parcela = rs("cd_parcela")
			cd_qtd_parcelas = rs("cd_qtd_parcelas")
			nm_obs = rs("nm_obs")
			cd_suspenso = rs("cd_suspenso")
			
			nm_pagamento = rs("nm_pagamento")
				if nm_pagamento <> "" Then
					nm_pagamento = nm_pagamento
				else
					nm_pagamento = nm_descricao
				end if
			dia_vencimento_padrao = rs("dia_vencimento_padrao")
			dt_vencimento = rs("dt_vencimento")
				dt_vencimento = zero(day(dt_vencimento))'&"/"&mesdoano(month(dt_vencimento))&"/"&year(dt_vencimento)
			cd_valor = rs("cd_valor")
			nm_valor = rs("nm_valor")
				if not isnumeric(nm_valor) then
					nm_valor = "0"
				end if
				
				if cd_tipo_valor = 1 then
					nm_valor_atual = "0"
					nm_valor_pendente = "0"
					nm_valor_extra = nm_valor
					'soma_semana = soma_semana + nm_valor_extra
				elseif cd_tipo_valor = 4 then
					nm_valor_atual = "0"
					nm_valor_pendente = nm_valor
					nm_valor_extra = "0"
					'soma_semana = soma_semana - nm_valor
				else
					nm_valor_atual = nm_valor
					nm_valor_pendente = "0"
					nm_valor_extra = "0"
					
					'soma_semana = soma_semana + nm_valor
				end if
				
				
				dt_vencimento_anterior = day(rs("dt_vencimento_anterior"))'&"/"&month(rs("dt_vencimento_anterior"))
				data_vencimento_anterior = rs("data_vencimento_anterior")'&"/"&month(rs("dt_vencimento_anterior"))
				nm_valor_anterior = rs("nm_valor_anterior")
					'if nm_valor_anterior > 0 OR  nm_valor > 0 Then
						'if cd_valor <> "" Then
							
							'status_variacao = instr(1,diferenca_valores,"-",1)
							'variacao_valores = formatnumber((diferenca_valores/nm_valor_anterior)*100,2)
							if nm_valor_atual <> 0 AND nm_valor_anterior <> 0 AND cd_valor <> 0 then
								diferenca_valores = nm_valor_atual - nm_valor_anterior
								'variacao_valores = formatnumber(diferenca_valores,2)
								'variacao_valores = diferenca_valores&":"&formatnumber((diferenca_valores/nm_valor_anterior)*100,2)
								variacao_valores = formatnumber((diferenca_valores/nm_valor_anterior)*100,2)
							end if
						'end if
				cd_modo_pagamento = rs("cd_modo_pagamento")
				cd_pagamento = rs("cd_cheque")
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
		'dia_semana = weekday(cabeca&"/"&dt_mes&"/"&dt_ano)
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
						<%=subtotal_valor%><%'total_soma_semana = soma_semana + soma_semana_extra
						'if total_soma_semana <> "" Then response.write("<b>"&formatNumber(total_soma_semana)&"</b>")%></td>
				</tr>
	<%total_soma_semana = 0
	cabecalho_semana = 1
	end if%>
	
	<%if abs(str_area) = "0" OR abs(str_area) = abs(cd_area)  Then%>
		
		<%if subtotal = 1 and ultimo_tipo <> cd_tipo_valor and num_linha > 1 then%>
			<tr>
				<td colspan="10" style="border-bottom:1px solid black;">&nbsp;</td>
				<td colspan="4" style="border-bottom:1px solid black;"><b>Subtotal: <%if isnumeric(subtotal_valor) then response.write(formatnumber(subtotal_valor,2))%></b></td>				
			</tr>
		<%subtotal_valor = 0
		end if
		
		ultimo_tipo = cd_tipo_valor
			
			%>
		
		<%if nm_valor <> 0 then
			'if nm_valor <> "" Then
				subtotal_valor = subtotal_valor + nm_valor
			'else
			'	subtotal_valor = subtotal_valor + nm_valor_anterior
			'end if
		else
			subtotal_valor = subtotal_valor + nm_valor_anterior
		end if
			%>
		
	
	<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
		<td align="center"><a href="#<%=num_linha%>"><%=num_linha%></a></td>
		<td align="center"><%=nm_area%></td>
		<td align="center"><%=nm_cc%></td>
		<td align="center"><%=nm_conta_abrv%></td>
		<td align="center"><%=nm_sigla%></td>
		<td align="center" <%if cd_tipo_valor = 4 then%>style="background-color:yellow;"<%end if%>><%=nm_tipo_abrv%> <%'="*"&dia_vencimento_padrao%></td>
		<td align="center"><%'=cd_modo_pagamento%> <%check1 = "check_1_"&cd_valor
		check2 = "check_2_"&cd_valor
		'if nm_valor = "0" Then 
		
			if cd_valor <> "" Then 
				str_acao = "editar_valor"
			else
				str_acao="inserir_valor"
			end if
		
		if cd_suspenso = 1 Then%>
			<img src="../../imagens/pagto-suspenso.gif" alt="Pagamento suspenso&#13;para o mês atual" width="12" height="12" border="0">
		<%elseif not IsNumeric(cd_valor) Then%>
			<img src="../../imagens/check_proibido.gif" alt="Inclua o valor" width="12" height="12" border="0" onClick="manipulacao_valor('<%=cd_diario%>','<%=dt_vencimento_anterior%>','<%=dt_mes%>','<%=dt_ano%>','<%=str_acao%>','<%=cd_user%>')"; return false;>
		<%elseif cd_confirma_pagto = int(1) AND cd_pagto_agendado = int(1) then%>
			<img src="../../imagens/check_ok.gif" alt="Pago" width="25" height="12" border="0">
		<%elseif cd_pagto_agendado = int(1) then%>
			<!--img src="../../imagens/check_aguarda.gif" alt="lista cheque(<%=nm_cheque%>)" width="25" height="12" border="0" onClick="window.open('financeiro/diario_pagamentos3.asp?cd_conta_banco=<%=cd_conta_banco%>&nm_cheque=<%=nm_cheque%>','lista_cheque','width=700,height=500,scrollbars=1')"-->
			<%if cd_modo_pagamento = "1" then%>
				<img src="../../imagens/pagto-cheque.gif" alt="lista cheque (<%=nm_cheque%>)" width="12" height="12" border="0" onClick="window.open('financeiro/diario_pagamentos3.asp?cd_pagamento=<%=cd_pagamento%>','lista_cheque','width=700,height=500,scrollbars=1')">
			<%elseif cd_modo_pagamento = "2" then%>
				<img src="../../imagens/pagto-transf.gif" alt="Transferência (<%=zero(day(dt_pagto))&"/"&left(mes_selecionado(month(dt_pagto)),3)&")"%>" width="12" height="12" border="0" onClick="window.open('financeiro/diario_pagamentos3.asp?cd_pagamento=<%=cd_pagamento%>','lista_transf','width=700,height=500,scrollbars=1')">
			<%end if%>
			</td>
		<%else%>	
			<!--img src="../../imagens/check_inativo.gif" alt="Inclui" width="25" height="12" border="0" onClick="adiciona_pgto('<%=cd_valor%>',document.diario.total_pgto.value,'<%=check1%>','<%=check2%>')" id="<%=check1%>" name="<%=check1%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>-->
			<img src="../../imagens/pagto-inativo.gif" alt="Inclui" width="12" height="12" border="0" onClick="adiciona_pgto('<%=cd_valor%>',document.diario.total_pgto.value,'<%=check1%>','<%=check2%>')" id="<%=check1%>" name="<%=check1%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>>
			<img src="../../imagens/check_ativo_2.gif" alt="Exclui" width="15" height="12" border="0" onClick="subtrai_pgto('<%=cd_valor%>',document.diario.total_pgto.value,'<%=check2%>','<%=check1%>')" id="<%=check2%>" name="<%=check2%>" <%'if ordem_var <> "" Then%>style="display:none;"<%'end if%>></td>
		<%end if%>
		<td><span onClick="window.open('financeiro/diario_lista_pag3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','Visualizar','width=500,height=<%if cd_tipo_valor=1 then%>240<%elseif cd_tipo_valor=2 then%>480<%elseif cd_tipo_valor=3 then%>460<%end if%>,scrollbars=1')" return false;" style="border-style:none;">
					<%if cd_fornecedor > 0 then response.write("<b>"&nm_fornecedor&"</b>")%>
					<%if nm_fornecedor <> "" and nm_pagamento <> "" Then response.write(" - ")%>
						<%=nm_pagamento%>&nbsp;</span>						
							<%if cd_qtd_parcelas > 1 then
								response.write("<span style='color:red;'><b>"&cd_parcela&"/"&cd_qtd_parcelas&"</b></span>")
							end if%>
							<%if nm_obs <> "" then%><img src="../imagens/ic_obs3.gif" alt="" width="10" height="10" border="0" onClick="window.open('financeiro/diario_lista_obs3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','Visualizar','width=530,height=270,scrollbars=1')" return false;"><%end if%></td>
		<!--td align="center" style="color:gray; background-color:<%'=cor_linha2%>;">&nbsp;<%'if dt_vencimento_anterior <> "" then 
									'response.write(zero(dt_vencimento_anterior))
								'end if%></td-->
		
		<td><img src="../imagens/ic_editar.gif" alt="Editar" width="13" height="14" border="0" onClick="window.open('financeiro/diario_cad3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&acao=editar&orient=reg','Cadastro','width=500,height=480,scrollbars=1')" return false;"><%'=%></td>
<!-- *************** ANTERIOR ***************-->		
		<td align="right" style="color:gray; background-color:<%=cor_linha2%>;"><%if nm_valor_anterior <> "0" then 
								response.write(FormatNumber(nm_valor_anterior,2))
								soma_valores_anteriores = soma_valores_anteriores + nm_valor_anterior
							else
								nm_valor_anterior = 0
							end if%><%'=" - "&data_vencimento_anterior%></td>
							
<!-- *************** DIA VENCIMENTO ***************-->
		<td align="center"><%if not isnumeric(dt_vencimento) Then
								response.write(zero(dia_vencimento_padrao)) 
							else
								response.write(dt_vencimento)
							end if%></td>

<!-- *************** VALOR FIXO ***************-->							
		<td align="right" <%if cd_tipo_valor <> 1 Then%> onClick="manipulacao_valor('<%=cd_diario%>','<%=dt_vencimento_anterior%>','<%=dt_mes%>','<%=dt_ano%>','<%=str_acao%>','<%=cd_user%>')"; return false;<%end if%>>
							<%if cd_suspenso = 1 and cd_tipo_valor <> 1 then
								response.write("---")
							elseif nm_valor_atual <> "0" then 
								response.write(FormatNumber(nm_valor_atual,2))
							elseif nm_valor_pendente <> "0" then 
								response.write("<span style='color:red;'><i>"&FormatNumber(nm_valor_pendente,2)&"</i></span>")
							end if%></td>							
							
<!-- *************** VALOR EXTRA ***************-->
		<td align="right" <%if cd_tipo_valor = 1 Then%>onClick="manipulacao_valor('<%=cd_diario%>','<%=dt_vencimento_anterior%>','<%=dt_mes%>','<%=dt_ano%>','<%=str_acao%>','<%=cd_user%>')"; return false;<%end if%>>
								<%if cd_suspenso = 1 and cd_tipo_valor = 1 then
									response.write("---")
								elseif nm_valor_extra <> "0" then 
									response.write(FormatNumber(nm_valor_extra,2))
								'elseif nm_valor_pendente <> "0" then 
								'	response.write("<span style='color=silver;'>"&FormatNumber(nm_valor_pendente,2)&"</span>")
								end if%></td>
		
							
<!-- *************** XXXXXXXXXX ***************-->							
		<!--td align="right"-->
			<%if cd_modo_pagamento > 0  then
				total_pago = total_pago + nm_valor_atual + nm_valor_extra
				'response.write(" * "&cd_modo_pagamento)
				'response.write(total_pago)
				'response.write("------------")
				
			elseif isnull(cd_modo_pagamento) OR cd_modo_pagamento < 1 then
				
				total_a_pagar = total_a_pagar + nm_valor_atual + nm_valor_extra
					if nm_valor=0 AND nm_valor_extra=0 AND isnull(cd_suspenso) then
						total_aberto = total_aberto + nm_valor_anterior
						'response.write("<span style='color:red;'>"&nm_valor_anterior&" + "&total_aberto&"</span>")
					else
						total_falta_pagto = total_falta_pagto + nm_valor_atual + nm_valor_extra
						'response.write("<span style='color:blue;'>"&total_falta_pagto&"</span>")
						
					end if
			end if%><!--/td-->
		<%'******* Cor padrão ********
			cor_variacao = cor_neg
		
		if variacao_valores = "" then
		'	seta = "px.gif"
			cor_variacao = ""'cor_0
		elseif cdbl(variacao_valores) > 7 then
		'	seta = "ic_indicador_up.gif"
			cor_variacao = cor_pos
		elseif cdbl(variacao_valores) > -7  then
		'	seta = "ic_indicador_up.gif"
			cor_variacao = ""'cor_0
		end if%>
		<td align="center" bgcolor="<%=cor_variacao%>">&nbsp;<img src="../imagens/px.gif<%'=seta%>" title="<%=variacao_valores&"%&#13;"&nm_obs%>" width="10" height="10" border="0"><%'=subtotal_valor%></td>
	</tr>	
	<%'end if
	
	
	soma_valores = abs(soma_valores) + nm_valor_atual - nm_valor_pendente
	soma_extra = abs(soma_extra) + nm_valor_extra
	soma_pendencias = abs(soma_pendencias) + nm_valor_pendente
	
	num_linha = num_linha + 1
	dt_vencimento_anterior = "0"
	nm_valor = 0
	'nm_valor_anterior = 0
	nm_valor_atual = 0
	nm_valor_pendente = 0
	nm_fornecedor = ""
	nm_pagamento = ""
	cd_valor = ""
	cd_suspenso = ""
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
	
	end if
		rs.movenext
		wend%>
	<%if subtotal = 1 then%>
			<tr>
				<td colspan="10" style="border-bottom:1px solid black;">&nbsp;</td>
				<td colspan="4" style="border-bottom:1px solid black;"><b>Subtotal: <%if isnumeric(subtotal_valor) then response.write(formatnumber(subtotal_valor,2))%></b></td>				
			</tr>
	<%end if%>
	<tr><td colspan="14" bgcolor="#808080"><img src="../imagens/px.gif" alt="" width="1" height="1" border="0"></td></tr>
	<tr>
		<td colspan="9"><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
		<!--td><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="400" height="1" border="0"></td-->
		<!--td colspan="2"><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td-->
		
		<td align="right" bgcolor="#ffffff"><b><%'if soma_valores <> "" Then response.write(formatNumber(soma_valores_anteriores,2))%></b></td>
		<td>&nbsp;<%'=soma_pendencias%></td>
		<td align="right" bgcolor="#c0c0c0"><b><%if soma_valores <> "" Then response.write(formatNumber(soma_valores,2))%></b></td>
		<td align="right" bgcolor="#c0c0c0"><b><%if soma_extra <> "" Then response.write(formatNumber(soma_extra,2))%></b></td>
		<td>&nbsp;</td>
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr>
	<%total_valores = soma_valores + soma_extra%>
	<!--tr>
		<td colspan="11" bgcolor="#c0c0c0">&nbsp;</td>
		<td align="center" colspan="2" bgcolor="Silver" style="font-size:15px;">&nbsp;<b><%'if total_valores <> "" Then response.write(FormatNumber(total_valores,2))%></b></td>
		<td>&nbsp;</td>
		<!--td>&nbsp;<%%>Pagamento</td-->
	</tr-->
	<tr><td align="right" colspan="13"><img src="../imagens/blackdot.gif" alt="" width="235" height="1" border="0"></td></tr>
	<!--tr>
		<td colspan="8" bgcolor="#ffffff">&nbsp;</td>
		<td align="center" rowspan="6"><img src="../imagens/blackdot.gif" alt="" width="1" height="80" border="0"></td>
		<td colspan="2" align="right">Total Lançado</td>
		<td align="right" colspan="2"><%'if total_falta_pagto <> "" Then response.write(formatNumber(total_valores,2))%></td>
	</tr-->
	<tr>
		<td colspan="8" bgcolor="#ffffff">&nbsp;</td>
		<td align="center" colspan="2" bgcolor="Silver" style="font-size:15px;">&nbsp;<b>TOTAL DO M&Ecirc;S</b></td>
		<td colspan="1" bgcolor="#ffffff">&nbsp;</td>
		<td align="center" colspan="2" bgcolor="Silver" style="font-size:15px;">&nbsp;<b><%if total_valores <> "" Then response.write(FormatNumber(total_valores,2))%></b></td>
		<td>&nbsp;</td>
	</tr>
	<tr><td align="right" colspan="13"><img src="../imagens/blackdot.gif" alt="" width="235" height="1" border="0"></td></tr>
	<tr>
		<td colspan="1" bgcolor="#ffffff">&nbsp;</td>
		<td colspan="6" bgcolor="#ffffff">&nbsp;<%if soma_pendencias <> "" Then%>Total de Pendencias: <%=formatNumber(soma_pendencias,2)%><%end if%></td>
		<td colspan="1" bgcolor="#ffffff">&nbsp;</td>
		<td colspan="2" align="right">Total pago</td>
		<td align="right" colspan="2"><%if total_pago <> "" Then response.write(formatNumber(total_pago,2))%></td>
	</tr>
	<!--tr>
		<td colspan="8" bgcolor="#ffffff">&nbsp;</td>
		<td colspan="2" align="right">Aguarda pagamento</td>
		<td align="right" colspan="2"><%if total_falta_pagto <> "" Then response.write(formatNumber(total_falta_pagto,2))%></td>
	</tr-->
	<tr><td align="right" colspan="13"><img src="../imagens/blackdot.gif" alt="" width="235" height="1" border="0"></td></tr>
	<tr>
		<td colspan="8" bgcolor="#ffffff">&nbsp;</td>
		<td align="right" colspan="2">Estim. não Lançado</td>
		<td align="right" colspan="2"><%if total_aberto <> "" Then%>
										<%=formatNumber(total_aberto,2)%>
									<%else%>
										<%="000"%>
									<%end if%></td>
	</tr>
	<%total_geral = total_pago + total_falta_pagto + total_aberto%>
	<!--tr>
		<td colspan="8" bgcolor="#ffffff">&nbsp;</td>
		<td align="center" colspan="2" bgcolor="Silver" style="font-size:15px;">&nbsp;<b>TOTAL GERAL</b></td>
		<td align="center" colspan="2" bgcolor="Silver" style="font-size:15px;">&nbsp;<b><%'if total_geral <> "" Then response.write(FormatNumber(total_geral,2))%></b></td>
		<td>&nbsp;</td>
	</tr-->
	<%total_fluxo = total_valores + total_aberto - total_pago%>
	<tr>
		<td colspan="8" bgcolor="#ffffff">&nbsp;</td>
		<td align="right" colspan="2">Necessidade de fluxo de caixa</td>
		<td align="right" colspan="2"><%if total_fluxo <> "" Then%>
										<%=formatNumber(total_fluxo,2)%>
									<%else%>
										<%="0,00"%>
									<%end if%></td>
	</tr>
	<tr>
		<td colspan="8" bgcolor="#ffffff">&nbsp;</td>
		<td align="center" colspan="2" bgcolor="Silver" style="font-size:15px;">&nbsp;<b>TOTAL ESTIMADO</b></td>
		<td colspan="1" bgcolor="#ffffff">&nbsp;</td><%total_valores_estimados = total_valores+total_aberto%>
		<td align="center" colspan="2" bgcolor="Silver" style="font-size:15px;">&nbsp;<b><%if total_valores_estimados <> "" Then%><%=FormatNumber(total_valores_estimados,2)%><%end if%></b></td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td><img src="../imagens/blackdot.gif" alt="" width="15" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="45" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="30" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="340" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="15" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="55" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="60" height="1" border="0"></td>
		<!--td><img src="../imagens/px.gif" alt="" width="55" height="1" border="0"></td-->
		<td><img src="../imagens/blackdot.gif" alt="" width="10" height="1" border="0"></td>
	</tr>
	</form>

</table>
