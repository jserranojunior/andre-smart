
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/numero_extenso.asp"-->
<!--#include file="../includes/util.asp"-->
<!--#include file="../css/geral.htm"-->

<%cd_user = session("cd_codigo")
pasta_loc = "financeiro\"
arquivo_loc = "fatura3.asp"

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
		
		mes_posterior = dt_mes + 1
		ano_posterior = dt_ano
			if mes_posterior > 12 then
				mes_posterior = 1
				ano_posterior = dt_ano + 1
			end if
			dia_final_posterior = ultimodiames(mes_posterior,ano_posterior)
			
novo_venc_dia = request("novo_venc_dia")
novo_venc_mes = request("novo_venc_mes")
novo_venc_ano = request("novo_venc_ano")

cd_unidade = request("cd_unidade")
	



data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)%>
<!--#include file="../includes/arquivo_loc.asp"-->
	<table align="center" class="no_print" border="0" style="border:2px solid black; border-collapse:collapse;">
	<form action="../financeiro/fatura3.asp" name="seleciona_periodo" id="seleciona_período">
		<input type="hidden" name="tipo" value="fatura">
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		<tr style="background-color:gray; color:white; font-size:14px;">
			<td align="center" colspan="4">Emissão de Faturas</td>
		</tr>
		<tr bgcolor="#c0c0c0">
			<td>&nbsp;<b>M&ecirc;s de vencimento</b></td>
			<td><b>Ano</b></td>
			<td rowspan="4"><img src="../imagens/px.gif" alt="" width="30" height="1" border="0">&nbsp;</td>
			<td rowspan="4" valign="bottom"><img src="../imagens/ic_print.gif" alt="imprimir" width="24" height="24" border="0" onclick="visualizarImpressao();" class="no_print"></td>
		</tr>
		<tr bgcolor="#c0c0c0">
			<td>&nbsp;<select name="mes_sel">
					<option value=""></option>
					<%for imes = 1 to 12%>
						<option value="<%=imes%>" <%if abs(imes)=abs(dt_mes) then response.write("SELECTED")%>><%=mes_selecionado(imes)%></option>
					<%next%>
				</select></td>
			<td><input type="text" name="ano_sel" maxlength="4" size="4" value="<%=dt_ano%>"></td>
		</tr>
		<tr bgcolor="#c0c0c0">
			<td colspan="2">&nbsp;<select name="cd_unidade">
					<option value="0">Selecione uma unidade</option>
					<%xsql = "SELECT * FROM TBL_unidades where cd_status = 1 and cd_hospital = 1"
					Set rs = dbconn.execute(xsql)%>
					<%while not rs.EOF
						strcd_unidade = rs("cd_codigo")
						strnm_unidade = rs("nm_unidade")
						strnm_sigla = rs("nm_sigla")%>
						<option value="<%=strcd_unidade%>" <%if abs(strcd_unidade) = abs(cd_unidade) then response.write("selected")%> ><%=strnm_unidade%></option>
					<%rs.movenext
					wend%>
				</select>

			</td>
		</tr>
		<tr bgcolor="#c0c0c0">
			<td colspan="2">&nbsp;<input type="submit" name="OK" value="Seleciona o mês de faturamento" width="4" height="5" style="background-color:orange;"></td>
		</tr>
		</form>
	</table>
	<br class="no_print">
	<br class="no_print">
	<br class="no_print">
<%if cd_unidade <> 0 Then%>
	<table align="center" border="1" style="border-collapse:collapse;" class="no_print" width="100">
		<!--form action="acoes/acoes3.asp" method="post" name="fatura" id="fatura"-->
	<form action="fatura3.asp" method="post" name="fatura" id="fatura">
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		<input type="hidden" name="acao" value="fatura_2">
		<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
		<input type="hidden" name="dt_faturamento" value="<%=dt_mes&"/1/"&dt_ano%>">
		<input type="hidden" name="name="dia_vencimento_<%=strcd_unidade%>"" value="<%=dia_vencimento%>">
		<input type="hidden" name="mes_sel" value="<%=dt_mes%>">
		<input type="hidden" name="ano_sel" value="<%=dt_ano%>">
		<!--tr>
			<td align="center"> &nbsp; &nbsp; &nbsp; &nbsp; </td>
			<td align="center"><b>Unidades</b></td>
		</tr-->
			<%num_linha = 1
			cab = 0
			'*** LISTA AS UNIDADES ***
			if cd_unidade = "0" OR cd_unidade = "" Then
				xsql = "SELECT * FROM TBL_unidades where cd_status = 1 and cd_hospital = 1"
			else
				xsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" AND cd_status = 1 and cd_hospital = 1"
			end if
			
			Set rs = dbconn.execute(xsql)%>
				<%while not rs.EOF
					strcd_unidade = rs("cd_codigo")
					strnm_unidade = rs("nm_unidade")
					strnm_sigla = rs("nm_sigla")
					
						'*** PREFERENCIAS DA UNIDADE/FATURAMENTO ***
						xsql = "SELECT * FROM TBL_unidade_faturamento_preferencias where cd_unidade = "&strcd_unidade&""
						Set rs_pref = dbconn.execute(xsql)
						
						if not rs_pref.EOF Then
							nm_dias_faturamento = rs_pref("nm_dias_faturamento")
							nm_dias_faturamento2 = rs_pref("nm_dias_faturamento2")
							
								nm_dia_vencimento1 = rs_pref("nm_dia_vencimento1")
								nm_dia_vencimento2 = rs_pref("nm_dia_vencimento2")
							cd_valor_fixo = rs_pref("cd_valor_fixo")
						end if%>
					<tr>
						<!--td align="center" valign="top" rowspan="1"><%=zero(num_linha)%><br><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td-->
						<td valign="top" rowspan="1"><a href="javascript:void(0);" onClick="window.open('../ferramentas/unidade_cad.asp?cd_unidade=<%=strcd_unidade%>&status=1&acao=editar&jan=1','Altera_nome','width=590,height=440,location=0,status=0,scrollbars=1')" return false;"><b style="font-size:10px;"><%=strnm_unidade%></b></a><br>
						Dias p/ faturamento: <b><%=nm_dias_faturamento%></b> 	<%if nm_dias_faturamento2 <> "" Then response.write(" e <b>"&nm_dias_faturamento2)%></b><br>
						Dia p/ vencimento: <b><%=nm_dia_vencimento1%></b> <%if nm_dia_vencimento2 <> "" Then response.write(" e <b>"&nm_dia_vencimento2)%></b><br>
						Referente à:
						
						
							<%if nm_dias_faturamento > 30 then
								dia_procedimentos = nm_dias_faturamento - 30
								mes_procedimentos = dt_mes - 2
								ano_procedimentos = dt_ano
									if mes_procedimentos = 0 then
										mes_procedimentos = 12
										ano_procedimentos = dt_ano - 1
									elseif mes_procedimentos < 0 then
										mes_procedimentos = 11
										ano_procedimentos = dt_ano - 1
									end if
									
									dia_final_procedimentos = ultimodiames(mes_procedimentos,ano_procedimentos)
							else
								dia_procedimentos = nm_dias_faturamento
								mes_procedimentos = dt_mes - 1
								ano_procedimentos = dt_ano
									if mes_procedimentos = 0 then
										mes_procedimentos = 12
										ano_procedimentos = dt_ano - 1
									elseif mes_procedimentos < 0 then
										mes_procedimentos = 11
										ano_procedimentos = dt_ano - 1
									end if
									
									dia_final_procedimentos = ultimodiames(mes_procedimentos,ano_procedimentos)
							end if%>
							
							<%'=dia_recebimento&"/"&mes_recebimento&"/"&ano_recebimento%>							
							<b><%=mes_selecionado(mes_procedimentos)&"/"&ano_procedimentos%></b><br>
															
						<img src="../imagens/px.gif" alt="" width="175" height="1" border="0"></td>
						<td>
							<table align="left" border="1" style="border-collapse:collapse;border:1px solid gray;">
								<%'*** TIPO DE PROCEDIMENTO ***
								xsql = "SELECT * FROM View_unidades_procedimento_tipo where cd_unidade = "&strcd_unidade&""
								SET rs_proc = dbconn.execute(xsql)
									while not rs_proc.EOF
										cd_procedimento_tipo = rs_proc("cd_procedimento_tipo")
										nm_procedimento_tipo = rs_proc("nm_procedimento_tipo_abrev")
										cd_rack_unid = rs_proc("cd_rack")
										
											if cab = 0 then%>
											<td align="center"><b>&nbsp;</b><br><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
											<td align="center"><b>Num</b><br><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
											<td align="center"><b>Valor</b><br><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
											<td align="center"><b>Total</b><br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
											<td align="center"><b>Venc.</b><br><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
											<td align="center" bgcolor="#FFFF80"><b>Novo Venc.</b><br><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
											<td align="center" bgcolor="#FFFF80">&nbsp;</td>
											<%end if%>
											<input type="hidden" name="cd_tipo_proc_<%=strcd_unidade%>" value="<%=cd_procedimento_tipo%>">
											<%'*** INFORMAÇÕES SOBRE A FATURA ***
											'* Procedimentos *
											strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE (cd_unidade = "&strcd_unidade&") AND (cd_tipo_proc = "&cd_procedimento_tipo&") AND (dt_faturamento = '"&mes_procedimentos&"/1/"&ano_procedimentos&"')"
											SET rs_qtd_proc = dbconn.execute(strsql)
												if not rs_qtd_proc.EOF then
													cd_fatura = rs_qtd_proc("cd_codigo")
													'nm_qtd_procedimentos = rs_qtd_proc("nm_qtd_procedimentos")
													dt_recebimento = rs_qtd_proc("dt_recebimento")
												end if%>
											<tr>
												<td><%=nm_procedimento_tipo%></td>
												<td align="center">
													<%if cd_valor_fixo = 1 then
														total_procs = 1%>
														<b>Fixo</b>
														<input type="hidden" name="qtd_proc_<%=strcd_unidade%>_<%=cd_procedimento_tipo%>" value="1"><%'=cd_fatura%>
													<%else%>
														<%if isnull(cd_rack_unid) then
															cond_rack = " "
														else
															cond_rack = " and a056_codrac="&cd_rack_unid&" "
														end if
														
														'**** MOSTRA A QUANTIDADE DE CIRURGIAS ******************************
														sql_proc = "SELECT COUNT (a002_numpro) AS conta ,a053_codung, nm_sigla,nm_sigla,a056_codrac FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_procedimentos&"/1/"&ano_procedimentos&"' AND '"&mes_procedimentos&"/"&dia_final_procedimentos&"/"&ano_procedimentos&"' AND a053_codung = "&strcd_unidade&" "&cond_rack&" AND a002_tipage <> 'E' AND cd_cancelado is null Group by a053_codung, nm_sigla,nm_sigla,a056_codrac"
														SET rs_procs = dbconn.execute(sql_proc)
															while not rs_procs.EOF
																qtd_proc_auto = rs_procs("conta")
																cd_rack = rs_procs("a056_codrac")
																
																total_procs = total_procs + qtd_proc_auto
															
															rs_procs.movenext
															wend%>
														<!--input type="text" name="qtd_proc_<%'=strcd_unidade%>_<%'=cd_procedimento_tipo%>" size="3" maxlength="5" value="<%'=total_procs%><%'=nm_qtd_procedimentos%>"--><b><%=total_procs%></b>
													<%end if%>
													<input type="hidden" name="cd_fatura_<%=strcd_unidade%>_<%=cd_procedimento_tipo%>" value="<%=cd_fatura%>"></td>
												<td align="right">
													<%'*** VALOR DO PROCEDIMENTO *********************************************
														strsql = "SELECT * FROM TBL_unidades_valores_procedimentos WHERE (cd_unidade = "&strcd_unidade&") AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' AND cd_procedimento_tipo = "&cd_procedimento_tipo&"  order by dt_atualizacao desc"
														SET rs_valor = dbconn.execute(strsql)
															'while not rs_proc.EOF
															if not rs_valor.EOF then
																nm_valor = rs_valor("nm_valor")%>
															<%=formatnumber(nm_valor)%>
															<%end if%>
													</td>
													<%'if nm_qtd_procedimentos <> "" then nm_total_unidade = nm_qtd_procedimentos * nm_valor
													if total_procs <> "" then nm_total_unidade =  total_procs * nm_valor%>
													<td align="right">&nbsp;<%=formatnumber(nm_total_unidade)%></td>
													
													<td align="center">
													<%if cab = 0 then%>											
															<%if nm_dias_faturamento > 30 then
																dia_recebimento = nm_dias_faturamento - 30
																mes_recebimento = dt_mes + 1
																ano_recebimento = dt_ano
																	if mes_recebimento > 12 then
																		mes_recebimento = 1
																		ano_recebimento = dt_ano + 1
																	end if
															else
																dia_recebimento = nm_dias_faturamento
																mes_recebimento = dt_mes
																ano_recebimento = dt_ano
															end if%>
															<%if nm_dias_faturamento2 > 30 then
																dia_recebimento2 = nm_dias_faturamento2 - 30
																mes_recebimento2 = dt_mes + 1
																ano_recebimento2 = dt_ano
																	if mes_recebimento2 > 12 then
																		mes_recebimento2 = 1
																		ano_recebimento2 = dt_ano + 1
																	end if
															else
																dia_recebimento2 = nm_dias_faturamento2
																mes_recebimento2 = dt_mes
																ano_recebimento2 = dt_ano
															end if%>
															
															<%=zero(nm_dia_vencimento1)&"/"&zero(dt_mes)%>
															<input type="hidden" name="dt_recebimento_<%=strcd_unidade&"_"&cab%>" value="<%=dt_mes&"/"&nm_dia_vencimento1&"/"&ano_recebimento%>">												
														<%else%>
															-
														<%end if%></td>
								<%dia_novo = request("dia_novo")
								mes_novo = request("mes_novo")
								ano_novo = request("ano_novo")
								complemento = request("complemento")
									%>
														<!--td align="center">
															<input type="text" name="dia_novo" maxlength="2" size="2" value="<%=dia_novo%>">
															<input type="text" name="mes_novo" maxlength="2" size="2" value="<%=mes_novo%>">
															<input type="text" name="ano_novo" maxlength="4" size="4" value="<%=ano_novo%>"></td>
														<td rowspan="2"><input type="submit" value="Ok" width="30" height="40"></td-->
											</tr>
								<%cab = cab + 1
								nm_qtd_procedimentos = ""
								nm_qtd_empr = ""
								nr_total_geral = nr_total_geral + nm_total_unidade
								nm_total_unidade = 0
								total_procs = 0
								rs_proc.movenext
								wend
								
								'* Empréstimos *
								strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE (cd_unidade = "&strcd_unidade&") AND (cd_tipo_proc = 5) AND (dt_faturamento = '"&dt_mes&"/1/"&dt_ano&"')"
								SET rs_qtd_proc = dbconn.execute(strsql)
								if not rs_qtd_proc.EOF then
									cd_fatura = rs_qtd_proc("cd_codigo")
									nm_qtd_empr = rs_qtd_proc("nm_qtd_procedimentos")
								end if%>
									<%sql_proc_e = "SELECT COUNT (a002_numpro) AS conta ,a053_codung, nm_sigla,nm_sigla,a056_codrac FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_procedimentos&"/1/"&ano_procedimentos&"' AND '"&mes_procedimentos&"/"&dia_final_procedimentos&"/"&ano_procedimentos&"' AND a053_codung = "&strcd_unidade&" AND a002_tipage = 'E' AND cd_cancelado is null Group by a053_codung, nm_sigla,nm_sigla,a056_codrac"
											SET rs_procs_e = dbconn.execute(sql_proc_e)
												'while not rs_procs_e.EOF
												if  not rs_procs_e.EOF then
													qtd_proc_auto = rs_procs_e("conta")
													cd_rack = rs_procs_e("a056_codrac")
													'response.write(qtd_proc_auto)
													total_procs_e = total_procs_e + qtd_proc_auto
												'rs_procs_e.movenext
												'wend
												end if%>
								<tr>
									<td>Emprés.</td>
									<td align="center"><input type="hidden" name="cd_fatura_<%=strcd_unidade%>_5" value="<%=cd_fatura%>">
									<!--input type="text" name="qtd_empr_<%'=strcd_unidade%>" size="3" maxlength="5" value="<%'=total_procs_e%><%'=nm_qtd_empr%>"--><%=total_procs_e%></td>
									<td colspan="2">&nbsp;</td>
									<td colspan="1">&nbsp;<%if not isnull(nm_dia_vencimento2) then response.write(zero(nm_dia_vencimento2)&"/"&zero(dt_mes))%></td>
									<td align="center">OBS:<!--input type="text" size="30" name="complemento__<%=strcd_unidade%>" value="<%'=cd_fatura%>"-->
									<input type="text" size="30" name="complemento" value="<%=complemento%>"></td>
								</tr>
		</form>
		<form action="acoes\acoes3.asp" method="post" name="fatura" id="fatura">
			<input type="hidden" name="acao" value="fatura_desconto">
			<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
			<input type="hidden" name="mes_sel" value="<%=dt_mes%>">
			<input type="hidden" name="ano_sel" value="<%=dt_ano%>">		
								<tr>
									<%strsql = "SELECT * FROM TBL_financeiro_fatura_desc WHERE (cd_unidade = "&strcd_unidade&") AND (dt_data = '"&dt_mes&"/1/"&dt_ano&"')"
										SET rs_desc = dbconn.execute(strsql)
										if not rs_desc.EOF then
											cd_desc = rs_desc("cd_codigo")
											nr_desconto = rs_desc("nr_desconto")
											nm_text_desconto = rs_desc("nm_text_desconto")
												nr_desconto_porcent = nr_desconto/100
												valor_desconto = nr_total_geral*nr_desconto_porcent
										end if%>
									<td style="background-color:yellow;">Desc.</td>
									<td style="background-color:yellow;" align="center" align="center"><input type="text" name="nr_desconto" size="2" maxlength="3" value="<%=nr_desconto%>">%</td>
									<td style="background-color:yellow;" align="center" align="center"><i><%if nr_desconto <> "" Then%><%=formatnumber(valor_desconto,2)%><%else%>&nbsp;<%end if%></i></td>
									<td style="background-color:yellow;" align="right" style="font-size:11px;"><b><%if nr_desconto <> "" Then%> <%=formatnumber(nr_total_geral-valor_desconto,2)%><%else%>&nbsp;<%end if%></b></td>
									<td style="background-color:yellow;" align="center" colspan="2"><input type="text" name="nm_text_desconto" size="30" maxlength="500" value="<%=nm_text_desconto%>"></td>
									<td><input type="submit" name="desc" value="Desc"></td>
								</tr>
								</form>
								<form action="fatura3.asp" method="post" name="fatura" id="fatura">
									<input type="hidden" name="tipo" value="fatura">
									<input type="hidden" name="cd_user" value="<%=cd_user%>">
									<input type="hidden" name="data_atual" value="<%=data_atual%>">
									<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
									<input type="hidden" name="mes_sel" value="<%=dt_mes%>">
									<input type="hidden" name="ano_sel" value="<%=dt_ano%>">
									<tr>
										<td colspan="2">Novo venc.</td>
										<td colspan="4">
											<input type="text" name="novo_venc_dia" maxlength="2" size="2" value="<%=novo_venc_dia%>">/
											<input type="text" name="novo_venc_mes" maxlength="2" size="2" value="<%=novo_venc_mes%>">/
											<input type="text" name="novo_venc_ano" maxlength="4" size="4" value="<%=novo_venc_ano%>">
										</td>
										<td><input type="submit" name="venc" value="Venc"></td>
									</tr>
								</form>
							</table>							
						</td>
						
						
					</tr>
				<%cab = 0
				total_procs = 0
				total_procs_e = 0
				num_linha = num_linha + 1
				rs.movenext
				wend
				
				%>
				<!--tr><td colspan="4"><input type="submit" value="Gravar"></td></tr-->
		
		
	</table>

	
	
	
	<%if cd_unidade = "0" OR cd_unidade = "" Then
		'xsql = "SELECT DISTINCT cd_unidade AS unid_faturamento FROM TBL_financeiro_fatura3"
	else
		xsql = "SELECT DISTINCT cd_unidade AS unid_faturamento FROM TBL_financeiro_fatura3 where cd_unidade = "&strcd_unidade&""
	end if
		Set rs = dbconn.execute(xsql)
		while not rs.EOF 
			unid_faturamento = rs("unid_faturamento")
		
			'**** Mostra a Unidade ****
			strsql = "SELECT * FROM View_unidades where cd_codigo = "&unid_faturamento&""
			Set rs_unid = dbconn.execute(strsql)
			'while not rs_unid.EOF
				cod_unidade = rs_unid("cd_codigo")
				nm_unidade_nome = rs_unid("nm_unidade_nome")
				nm_unidade = rs_unid("nm_unidade")
				'nm_sigla = rs_unid("nm_sigla")
				nm_contato = rs_unid("nm_contato")
				
		'dia_vencimento = rs_unid("dia_vencimento")
		'dia_vencimento2 = rs_unid("dia_vencimento2")
				nm_endereco = rs_unid("nm_endereco")
				
					id_rua = instr(1,nm_endereco,"!",1)-1
					id_numero_i = instr(1,nm_endereco,"!",1)+1
					id_numero_f = instr(1,nm_endereco,"#",1)-1
						id_numero_f = id_numero_f - id_numero_i+1
					
					id_complemento_i = instr(1,nm_endereco,"#",1)+1
					id_complemento_f = instr(1,nm_endereco,"~",1)
				
					if id_rua > 0 then nm_rua = mid(nm_endereco,1,id_rua)&","
					if id_numero_f > 0 then nm_numero = mid(nm_endereco,id_numero_i,id_numero_f)
					if id_complemento_f > 0 then nm_complemento = mid(nm_endereco,id_complemento_i,id_complemento_f-id_complemento_i)
			
				nm_cnpj = rs_unid("nm_cnpj")
				cd_cliente_empresa = rs_unid("cd_cliente_empresa")
				nm_empresa = rs_unid("nm_empresa")
				cd_cnpj_emp = rs_unid("cd_cnpj_emp")
				nm_endereco_emp = rs_unid("nm_endereco_emp")
				nm_numero_emp = rs_unid("nm_numero_emp")
				nm_complemento_emp = rs_unid("nm_complemento_emp")
				nm_bairro_emp = rs_unid("nm_bairro_emp")
				nm_cidade_emp = rs_unid("nm_cidade_emp")
				nm_estado_emp = rs_unid("nm_estado_emp")
				nm_cep_emp = rs_unid("nm_cep_emp")
				telefone_emp = rs_unid("telefone_emp")
				nm_email_emp = rs_unid("nm_email_emp")
				nm_site_emp = rs_unid("nm_site_emp")
				
				cd_contacorrente = rs_unid("cd_conta_pagamento")
				'nm_banco_emp = rs_unid("nm_banco_emp")
				'nm_agencia_emp = rs_unid("nm_agencia_emp")
				'nm_contacorrente_emp = rs_unid("nm_contacorrente_emp")
				
				'if cd_empresa > 0 then
								strsql = "SELECT * FROM TBL_banco where cd_codigo = '"&cd_contacorrente&"'"
								Set rs_banco = dbconn.execute(strsql)
									while not rs_banco.EOF
										cd_banco = rs_banco("cd_codigo")
										nm_banco_emp = rs_banco("nm_banco")
										nm_agencia_emp = rs_banco("nm_agencia")
										nm_contacorrente_emp = rs_banco("nm_contacorrente")
									rs_banco.movenext
									wend
								
							'end if
			
		'	*** Cálculo para vencimento de pagamento ***
			xsql = "SELECT * FROM TBL_unidade_faturamento_preferencias where cd_unidade = "&cod_unidade&""
			Set rs_pref = dbconn.execute(xsql)
				if not rs_pref.EOF Then
					nm_dias_faturamento = rs_pref("nm_dias_faturamento")
					nm_dias_faturamento2 = rs_pref("nm_dias_faturamento2")
					
					nm_dia_vencimento1 = rs_pref("nm_dia_vencimento1")
					nm_dia_vencimento2 = rs_pref("nm_dia_vencimento2")
					
					cd_mostra_proc = rs_pref("cd_mostra_proc")
					cd_mostra_emp = rs_pref("cd_mostra_emp")
					cd_mostra_proc_duplo = rs_pref("cd_mostra_proc_duplo")
					
					cd_impress_pg1 = rs_pref("cd_impress_pg1")
					cd_impress_pg2 = rs_pref("cd_impress_pg2")
					cd_impress_pg3 = rs_pref("cd_impress_pg3")
					
					cd_valor_fixo = rs_pref("cd_valor_fixo")
					'response.write("aaa ")
						if nm_dias_faturamento > 30 then
								dia_procedimentos = nm_dias_faturamento - 30
								mes_procedimentos = dt_mes - 2
								ano_procedimentos = dt_ano
									if mes_procedimentos = 0 then
										mes_procedimentos = 12
										ano_procedimentos = dt_ano - 1
									elseif mes_procedimentos < 0 then
										mes_procedimentos = 11
										ano_procedimentos = dt_ano - 1
									end if
									
									dia_final_procedimentos = ultimodiames(mes_procedimentos,ano_procedimentos)
							else
								dia_procedimentos = nm_dias_faturamento
								mes_procedimentos = dt_mes - 1
								ano_procedimentos = dt_ano
									if mes_procedimentos = 0 then
										mes_procedimentos = 12
										ano_procedimentos = dt_ano - 1
									elseif mes_procedimentos < 0 then
										mes_procedimentos = 11
										ano_procedimentos = dt_ano - 1
									end if
									
									dia_final_procedimentos = ultimodiames(mes_procedimentos,ano_procedimentos)
							end if
						
						
						'**** MOSTRA A QUANTIDADE DE CIRURGIAS ******************************
								sql_proc = "SELECT COUNT (a002_numpro) AS conta ,a053_codung, nm_sigla,nm_sigla,a056_codrac FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_procedimentos&"/1/"&ano_procedimentos&"' AND '"&mes_procedimentos&"/"&dia_final_procedimentos&"/"&ano_procedimentos&"' AND a053_codung = "&cod_unidade&" "&cond_rack&" AND a002_tipage <> 'E' AND cd_cancelado is null Group by a053_codung, nm_sigla,nm_sigla,a056_codrac"
								SET rs_procs = dbconn.execute(sql_proc)
									while not rs_procs.EOF
										qtd_proc_auto = rs_procs("conta")
										cd_rack = rs_procs("a056_codrac")
										total_procs = total_procs + qtd_proc_auto
										'response.write("xxx")
									
									rs_procs.movenext
									wend
									
								if cd_valor_fixo = 1 then
									total_procs = 1
								end if
						'*** VALOR DO PROCEDIMENTO *********************************************
								strsql = "SELECT * FROM TBL_unidades_valores_procedimentos WHERE (cd_unidade = "&cod_unidade&") AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' AND cd_procedimento_tipo = "&cd_procedimento_tipo&"  order by dt_atualizacao desc"
								SET rs_valor = dbconn.execute(strsql)
									'while not rs_proc.EOF
									if not rs_valor.EOF then
										nm_valor = rs_valor("nm_valor")%>
									<%'=formatnumber(nm_valor)%>
									<%end if
				end if
				
				'if dia_novo <> "" Then
				'	nm_dia_vencimento1 = dia_novo
				'end if
				'if mes_novo <> "" Then
				'	dt_mes = mes_novo
				'end if
				'if dia_novo <> "" Then
				'	dt_ano = ano_novo
				'end if
				'nm_dia_vencimento2 = 
				'complemento = request("complemento")'", referente ao pedido 551560409"
				
				
				'teste = "<br>____"&cod_unidade&"-"&qtd_proc_auto&"-"&total_procs&"___<br>"
				total_procs = 0
				
				'*** Corrige a data de vencimento e inclui texto complementar ***
				if dia_novo <> "" AND mes_novo <> "" AND ano_novo <> "" Then
					nm_dia_vencimento1 = dia_novo
				end if
				if complemento <> "" then
					complemento = ", "&complemento
				end if
				
				
				
				if novo_venc_ano <> "" OR novo_venc_ano <> "" OR novo_venc_ano <> "" then
					novo_venc_dia = int(novo_venc_dia)
					novo_venc_mes = int(novo_venc_mes)
					novo_venc_ano = int(novo_venc_ano)
				end if
				
				%>
				
			<%if cd_unidade = "" and salta_linha = 1 then %><br style="page-break-before:always;"><%end if%>
	<!-- ********************************************
	*********  1 - PROTOCOLO DE RECEBIMENTO *********
	**********************************************-->
			<%if cd_impress_pg1 = 1 then%>
			<table align="center" width="500" border="0" style="border-collapse:collapse;">
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="47" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td>&nbsp;</td>
					<td align="right" valign="top"><img src="../imagens/logotipo<%=cd_cliente_empresa%>.gif" alt="" width="170" height="118" border="0"></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="84" border="0"></td></tr><tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td align="center" colspan="2" style="text-decoration:none; font-family:arial; font-size:21px;"><b>Protocolo de Recebimento de Documentos</b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="90" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="2" valign="top" style="font-family:arial; font-size:18px; text-align:justify; text-justify:inter-word;">
						<%if cd_valor_fixo = 1 then%>
							Recebi da <b><%=nm_empresa%></b>, o recibo de pagamento da locação dos equipamentos de videocirurgia relativo ao mesmo mês com vencimento em <b><%=zero(nm_dia_vencimento1)%> <%if nm_dia_vencimento2 <> "" Then response.write(" e "&zero(nm_dia_vencimento2))%> de <%=mes_selecionado(dt_mes)&" de "&dt_ano%></b><%=complemento%>.
						<%else%>
							Recebi da <b><%=nm_empresa%></b>, o relatório dos procedimentos realizados por videocirurgia no mês de <%=mes_selecionado(mes_procedimentos)&" de "&ano_procedimentos%>, no <b><%=nm_unidade_nome%></b>, junto com o recibo de pagamento da locação dos equipamentos de videocirurgia relativo ao mês com vencimento em <b>
								<%if novo_venc_ano <> "" OR  novo_venc_ano <> "" OR  novo_venc_ano <> "" Then%>
									<%=zero(novo_venc_dia)%> <%if nm_dia_vencimento2 <> "" Then response.write(" e "&nm_dia_vencimento2)%> de <%=mes_selecionado(novo_venc_mes)&" de "&novo_venc_ano%></b><%=complemento%>.
								<%else%>
									<%=zero(nm_dia_vencimento1)%> <%if nm_dia_vencimento2 <> "" Then response.write(" e "&nm_dia_vencimento2)%> de <%=mes_selecionado(dt_mes)&" de "&dt_ano%></b><%=complemento%>.
								<%end if%>
						<%end if%></td>
				</tr>
				<tr><td colspan="3"><%'=nm_dias_faturamento&"*"&nm_dias_faturamento2&" <br>"&nm_dia_vencimento1&"+"&nm_dia_vencimento2%><img src="../imagens/px.gif" alt="" width="1" height="130" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td align="right" colspan="2" style="font-family:arial; font-size:17px;"><b>São Paulo, ____/____/_______</b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="130" border="0"></td></tr>
				<tr>
					<td colspan="2">&nbsp;<!-- limitador de linha --></td>
					<td align="center" style="font-family:arial; font-size:15px;"><b>Nome / Assinatura</b></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;<!-- limitador de linha --></td>
					<td align="center" style="font-family:arial; font-size:14px;"><b><%=nm_unidade_nome%></b></td>
				</tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td><img src="../imagens/px.gif" alt="" width="415" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="253" height="1" border="0"></td>
				</tr>
			</table>
			<%End if%>
	<!-- ************************************************
	*********  2 - DEMONSTRATIVO DO FATURAMENTO *********
	**************************************************-->
			<%if cd_impress_pg2 = 1 then%>
			<br style="page-break-after:always;">
			
			<table width="650" align="center" border="0">
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="46" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="1" style="font-family:arial; font-size:15px;">
						São Paulo,
							<%if novo_venc_dia <> "" OR novo_venc_mes <> "" OR novo_venc_ano <> "" Then%>
								<%=zero(novo_venc_dia) &" de "&mes_selecionado(novo_venc_mes)&" de "&novo_venc_ano%>
							<%else%>
								<%=zero(nm_dia_vencimento1) &" de "&mes_selecionado(dt_mes)&" de "&dt_ano%>
							<%end if%>.</td>
					<td align="right" valign="top" rowspan="4"><img src="../imagens/logotipo<%=cd_cliente_empresa%>.gif" alt="" width="170" height="120" border="0"></td>
				</tr>
				<tr><td colspan="2"><img src="../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="1" style="font-family:arial; font-size:15px;"><b>Ao<br><%=nm_unidade_nome%><br>&nbsp;<br>A/C: <%=nm_contato%></b></td>
				</tr>
				<tr><td colspan="2"><img src="../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td align="center" colspan="2" style="text-decoration: underline; font-family:arial; font-size:17px;"><b> Faturamento de <%=mes_selecionado(dt_mes)&" "&dt_ano%></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="45" border="0"></td></tr>
				<tr>
					<td><img src="../imagens/px.gif" alt="" width="1" height="500" border="0"><!-- limitador de linha --></td>
					<td colspan="2" valign="top" style="font-family:arial; font-size:16px; text-align:justify; text-justify:inter-word;" >
						Em contrato firmado, informamos o número de procedimentos realizados no centro cirurgico do <%=nm_unidade_nome%> no período de 01 a <%=ultimodiames(mes_procedimentos,ano_procedimentos)&" de "&mes_selecionado(mes_procedimentos)&" de "&ano_procedimentos%>, conforme tabela anexa<%=complemento%>.<br>
						<img src="../imagens/px.gif" alt="" width="1" height="40" border="0"><br>
								
								
								<!--*******************************************************************************************************************-->
								<table border="1" width="500" style="font-family:arial; font-size:15px; border-collapse:collapse;padding:5px 4px; background:transparent;" align="center">
								<%'*** TIPO DE PROCEDIMENTO ***
								xsql = "SELECT * FROM View_unidades_procedimento_tipo where cd_unidade = "&strcd_unidade&""
								SET rs_proc = dbconn.execute(xsql)
									while not rs_proc.EOF
										cd_procedimento_tipo = rs_proc("cd_procedimento_tipo")
										nm_procedimento_tipo = rs_proc("nm_procedimento_tipo_abrev")
										cd_rack_unid = rs_proc("cd_rack")
										qtd_linha_procs = qtd_linha_procs + 1
											if cab = 0 then%>
											<tr>
												<td align="center" style="padding:15px 5px;"><%if cd_tipo_proc > 1 then%> <%larg_cel = "180"%><b>Descrição</b><%else%><%larg_cel = "110"%>&nbsp;<%end if%><br><img src="../imagens/px.gif" alt="" width="<%=larg_cel%>" height="1" border="0"></td>
												<td align="center" style="padding:15px 5px;"><b>N° de Procedimentos</b><br><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
												<td align="center" style="padding:15px 5px;"><b>Custo / Cirurgia</b><br><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
												<td align="center" style="padding:15px 5px;"><b>Valor Total</b><br><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
											</tr>
											<%end if%>
											<input type="hidden" name="cd_tipo_proc_<%=strcd_unidade%>" value="<%=cd_procedimento_tipo%>">
											<%'*** INFORMAÇÕES SOBRE A FATURA ***
											'* Procedimentos *
											strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE (cd_unidade = "&strcd_unidade&") AND (cd_tipo_proc = "&cd_procedimento_tipo&") AND (dt_faturamento = '"&mes_procedimentos&"/1/"&ano_procedimentos&"')"
											SET rs_qtd_proc = dbconn.execute(strsql)
												if not rs_qtd_proc.EOF then
													cd_fatura = rs_qtd_proc("cd_codigo")
													'nm_qtd_procedimentos = rs_qtd_proc("nm_qtd_procedimentos")
													dt_recebimento = rs_qtd_proc("dt_recebimento")
												end if%>
											<tr>
												<td align="center" style="padding:15px 5px;"><%=nm_procedimento_tipo%></td>
												<td align="center" style="padding:15px 5px;">
													<%if cd_valor_fixo = 1 then
														total_procs = 1%>
														<b>Fixo</b>
														<input type="hidden" name="qtd_proc_<%=strcd_unidade%>_<%=cd_procedimento_tipo%>" value="1"><%'=cd_fatura%>
													<%else%>
														<%if isnull(cd_rack_unid) then
															cond_rack = " "
														else
															cond_rack = " and a056_codrac="&cd_rack_unid&" "
														end if
														
														'**** MOSTRA A QUANTIDADE DE CIRURGIAS ******************************
														sql_proc = "SELECT COUNT (a002_numpro) AS conta ,a053_codung, nm_sigla,nm_sigla,a056_codrac FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_procedimentos&"/1/"&ano_procedimentos&"' AND '"&mes_procedimentos&"/"&dia_final_procedimentos&"/"&ano_procedimentos&"' AND a053_codung = "&strcd_unidade&" "&cond_rack&" AND a002_tipage <> 'E' AND cd_cancelado is null Group by a053_codung, nm_sigla,nm_sigla,a056_codrac"
														SET rs_procs = dbconn.execute(sql_proc)
															while not rs_procs.EOF
																qtd_proc_auto = rs_procs("conta")
																cd_rack = rs_procs("a056_codrac")
																
																total_procs = total_procs + qtd_proc_auto
																'total_procs = 159
															'response.write(qtd_proc_auto&"<br>")
															rs_procs.movenext
															wend%>
														<b><%=total_procs%></b>
													<%end if%>
													</td>
													<td align="right" style="padding:15px 5px;">
													<%'*** VALOR DO PROCEDIMENTO *********************************************
														strsql = "SELECT * FROM TBL_unidades_valores_procedimentos WHERE (cd_unidade = "&strcd_unidade&") AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' AND cd_procedimento_tipo = "&cd_procedimento_tipo&"  order by dt_atualizacao desc"
														SET rs_valor = dbconn.execute(strsql)
															'while not rs_proc.EOF
															if not rs_valor.EOF then
																nm_valor = rs_valor("nm_valor")%>
															<%=formatnumber(nm_valor)%>
															<%end if%>
													</td>
													<%'if nm_qtd_procedimentos <> "" then nm_total_unidade = nm_qtd_procedimentos * nm_valor
													if total_procs <> "" then nm_total_unidade =  total_procs * nm_valor
													total_geral = total_geral + nm_total_unidade%>
													<td align="right" style="padding:15px 5px;">&nbsp;<%=formatnumber(nm_total_unidade)%></td>
											</tr>
								<%cab = cab + 1
								nm_qtd_procedimentos = ""
								nm_qtd_empr = ""
								nm_total_unidade = 0
								total_procs = 0
								rs_proc.movenext
								wend
								
								'* Empréstimos *
								strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE (cd_unidade = "&strcd_unidade&") AND (cd_tipo_proc = 5) AND (dt_faturamento = '"&dt_mes&"/1/"&dt_ano&"')"
								SET rs_qtd_proc = dbconn.execute(strsql)
								if not rs_qtd_proc.EOF then
									cd_fatura = rs_qtd_proc("cd_codigo")
									nm_qtd_empr = rs_qtd_proc("nm_qtd_procedimentos")
								end if%>
									<%sql_proc_e = "SELECT COUNT (a002_numpro) AS conta ,a053_codung, nm_sigla,nm_sigla,a056_codrac FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_procedimentos&"/1/"&ano_procedimentos&"' AND '"&mes_procedimentos&"/"&dia_final_procedimentos&"/"&ano_procedimentos&"' AND a053_codung = "&strcd_unidade&" AND a002_tipage = 'E' AND cd_cancelado is null Group by a053_codung, nm_sigla,nm_sigla,a056_codrac"
											SET rs_procs_e = dbconn.execute(sql_proc_e)
												'while not rs_procs_e.EOF
												if  not rs_procs_e.EOF then
													qtd_proc_auto = rs_procs_e("conta")
													cd_rack = rs_procs_e("a056_codrac")
													'response.write(qtd_proc_auto)
													'total_procs_e = total_procs_e + qtd_proc_auto
												'rs_procs_e.movenext
												'wend
												end if
												
												%>
								<%if qtd_linha_procs > 1 then%>
								<tr>
									<td align="center" style="padding:15px 5px;"><b>TOTAL</b></td>
									<td align="center" style="padding:15px 5px;">-<%'=qtd_linha_procs%></td>
									<td align="center" style="padding:15px 5px;">-</td>
									<td align="right" style="padding:15px 5px;"><%=formatnumber(total_geral)%></td>
								</tr>
								<%end if%>
								
									<%strsql = "SELECT * FROM TBL_financeiro_fatura_desc WHERE (cd_unidade = "&strcd_unidade&") AND (dt_data = '"&dt_mes&"/1/"&dt_ano&"')"
										SET rs_desc = dbconn.execute(strsql)
										if not rs_desc.EOF then
											cd_desc = rs_desc("cd_codigo")
											nr_desconto = rs_desc("nr_desconto")
											nm_text_desconto = rs_desc("nm_text_desconto")
												nr_desconto_porcent = nr_desconto/100
												valor_desconto = nr_total_geral*nr_desconto_porcent%>
										<tr><td colspan="4">&nbsp;</td></tr>
										<tr>
											<td align="center" style="padding:15px 5px;">Desconto<br><span style="font-size:12px;"><%if nm_text_desconto <> "" Then response.write(nm_text_desconto)%></span></td>
											<td align="center" style="padding:15px 5px;"><%if nr_desconto <> "" Then%> <%=nr_desconto%>%<%else%>&nbsp;<%end if%></td>
											<td align="center" style="padding:15px 5px;"><%if nr_desconto <> "" Then%> <%=formatnumber(valor_desconto,2)%><%else%>&nbsp;<%end if%></td>
											<td align="right" style="padding:15px 5px; font-size:16px;"><b><%if nr_desconto <> "" Then%> <%=formatnumber(nr_total_geral-valor_desconto,2)%><%else%>&nbsp;<%end if%></b></td>
										</tr>
										<%end if%>
								<tr>
									<td align="center" style="padding:15px 5px;"><b>Vencimento</b></td>
									<td align="center" colspan="3" style="padding:15px 5px;">
												<%if novo_venc_dia <> "" OR novo_venc_mes <> "" OR novo_venc_ano <> "" then%>
													<%=zero(novo_venc_dia)&"/"&zero(novo_venc_mes)&"/"&novo_venc_ano%> 
												<%else%>													
															<%if nm_dias_faturamento > 30 then
																dia_recebimento = nm_dias_faturamento - 30
																mes_recebimento = dt_mes + 1
																ano_recebimento = dt_ano
																	if mes_recebimento > 12 then
																		mes_recebimento = 1
																		ano_recebimento = dt_ano' + 1
																	end if
															else
																dia_recebimento = nm_dias_faturamento
																mes_recebimento = dt_mes
																ano_recebimento = dt_ano
															end if%>
															<%if nm_dias_faturamento2 > 30 then
																dia_recebimento2 = nm_dias_faturamento2 - 30
																mes_recebimento2 = dt_mes + 1
																ano_recebimento2 = dt_ano
																	if mes_recebimento2 > 12 then
																		mes_recebimento2 = 1
																		ano_recebimento2 = dt_ano' + 1
																	end if
															else
																dia_recebimento2 = nm_dias_faturamento2
																mes_recebimento2 = dt_mes
																ano_recebimento2 = dt_ano
															end if%>
															
															<%=zero(nm_dia_vencimento1)&"/"&zero(dt_mes)&"/"&ano_recebimento%> 
															<%if nm_dia_vencimento2 <> "" Then
																response.write(" - "&formatnumber(total_geral/2)&"<br>"&zero(nm_dia_vencimento2)&"/"&zero(dt_mes)&"/"&ano_recebimento&" - "&formatnumber(total_geral/2))
															end if%>
												<%end if%>
											</td>
								</tr>
							</table>
							
							<img src="../imagens/px.gif" alt="" width="1" height="35" border="0"><br>
							<%'strsql = "SELECT * FROM VIEW_unidades_fatura_valores_procedimentos WHERE (cd_unidade = "&cod_unidade&") AND (cd_tipo_proc < 5)"
							strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE cd_unidade="&cod_unidade&" AND cd_tipo_proc >= 5 AND dt_faturamento='"&dt_mes&"/1/"&dt_ano&"'"
							SET rs_mens = dbconn.execute(strsql)
								while not rs_mens.EOF
									cd_tipo_proc = rs_mens("cd_tipo_proc")
									nm_qtd_procedimentos = rs_mens("nm_qtd_procedimentos")%>
								
									<%if cd_mostra_emp = 1 then
										if cd_tipo_proc = 5 AND nm_qtd_procedimentos > 0 Then%>										
											No mês de <%=mes_selecionado(mes_procedimentos)%> houve <%=nm_qtd_procedimentos%> empréstimo de material Video Cirúrgico para outras equipes.
										<%else%>
											No mês de <%=mes_selecionado(mes_procedimentos)%> não houve empréstimo de material Video Cirúrgico para outras equipes.
										<%end if
									end if%>
									<%if cd_mostra_proc_duplo = 1 then
										'*** Mensagem para procedimentos duplos ***
										if cd_tipo_proc = 6 AND nm_qtd_procedimentos > 0 Then%>										
											No mês de <%=mes_selecionado(mes_procedimentos)%> houve <%'=nm_qtd_procedimentos%> pacientes com procedimento duplo.
										<%else%>
											No mês de <%=mes_selecionado(mes_procedimentos)%> não houve pacientes com procedimento duplo.
										<%end if
									end if%>
									
								<%rs_mens.movenext
								wend%>
								</td>
							
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="65" border="0"></td></tr>
				<tr>
					<td colspan="2">&nbsp;<!-- limitador de linha --></td>
					<td align="center" colspan="1" style="font-family:arial; font-size:15px;"><b><%=nm_empresa%><br><span style="font-family:arial; font-size:13px;">CNPJ: <%=cd_cnpj_emp%></span></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="73" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td align="center" colspan="5" style="font-family:arial; font-size:13px;">
					<img src="../imagens/blackdot.gif" alt="" width="660" height="1" border="0"><br><%=nm_endereco_emp%>, <%=nm_numero_emp%> - Cep: <%=nm_cep_emp%> - <%=nm_cidade_emp%> - <%=nm_estado_emp%> - Fone:(11) <%=telefone_emp%></td>
				</tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td><img src="../imagens/px.gif" alt="" width="445" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="220" height="1" border="0"></td>
				</tr>
			</table>
			<%end if%>
	<!-- ***************************************
	*********  3 - RECIBO DE PAGAMENTO *********
	*****************************************-->
			<%if cd_impress_pg3 = 1 then%>
			<br style="page-break-after:always;">
			<table width="500" border="0" align="center" style="border-collapse:collapse;">
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td style="font-family:arial; font-size:14px; text-align:left;"><b><%=nm_empresa%></b><br>
					<%=nm_endereco_emp%>, <%=nm_numero_emp%> <%if nm_complemento_emp <> "" Then response.write(" - "&nm_complemento_emp)%><br>
					CEP: <%=nm_cep_emp%> - <%=nm_cidade_emp%> - <%=nm_estado_emp%><br>
					Tel/Fax: (11) 2631- 4959 <br>
					CNPJ: <%=cd_cnpj_emp%><br>
					<%=nm_email_emp%><br>
					<%=nm_site_emp%></td>
					<td align="right"><img src="../imagens/logotipo<%=cd_cliente_empresa%>.gif" alt="" width="170" height="120" border="0"></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="80" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="2" style="font-family:arial; font-size:15px; text-align:right;">
						São Paulo,
							<%if novo_venc_dia <> "" OR novo_venc_mes <> "" OR novo_venc_ano <> "" Then%>
								<%=zero(novo_venc_dia) &" de "&mes_selecionado(novo_venc_mes)&" de "&novo_venc_ano%>
							<%else%>
								<%=zero(nm_dia_vencimento1) &" de "&mes_selecionado(dt_mes)&" de "&dt_ano%>
							<%end if%>.
					</td>
				</tr>
				<tr>
					<td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="80" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="2" style="font-family:arial; font-size:17px; text-align:center;"><b>R E C I B O</b></td>
				</tr>
				<tr>
					<td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="10" border="0"></td>
				</tr>
				<%'strsql = "SELECT * FROM VIEW_unidades_fatura_valores_procedimentos WHERE (cd_unidade = "&cod_unidade&") AND (cd_tipo_proc < 5) AND dt_faturamento='"&dt_mes&"/1/"&dt_ano&"'"
				'						SET rs_proc = dbconn.execute(strsql)
				'							while not rs_proc.EOF
				'								cd_tipo_proc = rs_proc("cd_tipo_proc")
				'								nm_procedimento_tipo = rs_proc("nm_procedimento_tipo")
				'								nm_qtd_procedimentos = rs_proc("nm_qtd_procedimentos")
				'								
				'								strsql = "SELECT * FROM TBL_unidades_valores_procedimentos WHERE (cd_unidade = "&cod_unidade&") AND dt_atualizacao <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"' AND cd_procedimento_tipo = "&cd_tipo_proc&" order by dt_atualizacao desc"
				'								SET rs_valor = dbconn.execute(strsql)
				'									if not rs_valor.EOF then
				'										nm_valor_procedimento = rs_valor("nm_valor")
				'									rs_valor.movenext
				'									end if%>
												
													
													
													<%'valor_total_procedimentos = nm_qtd_procedimentos * nm_valor_procedimento%>
													
												
											
											<%'total_procedimentos = total_procedimentos + nm_qtd_procedimentos
											'valor_total_recibo = valor_total_recibo + valor_total_procedimentos
											
											'rs_proc.movenext
											'wend
											
											if isnumeric(cd_valor_fixo) then
												valor_total_recibo = nm_valor
											else 
												valor_total_recibo = total_geral'total_procs*nm_valor
												'valor_total_recibo = "71680,11"
											end if
											
											strsql = "SELECT * FROM TBL_financeiro_fatura_desc WHERE (cd_unidade = "&strcd_unidade&") AND (dt_data = '"&dt_mes&"/1/"&dt_ano&"')"
												SET rs_desc = dbconn.execute(strsql)
												if not rs_desc.EOF then
													cd_desc = rs_desc("cd_codigo")
													nr_desconto = rs_desc("nr_desconto")
														nr_desconto_porcent = nr_desconto/100
														valor_desconto = nr_total_geral*nr_desconto_porcent
														valor_total_recibo_s_desconto = valor_desconto
														valor_total_recibo = valor_total_recibo - valor_desconto
														
												end if%><tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="2" style="font-family:arial; font-size:17px; text-align:center;"><%if valor_total_recibo <> "" Then response.write(formatCurrency(valor_total_recibo))%></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="60" border="0"></td></tr>
				<tr>
					<td><img src="../imagens/px.gif" alt="" width="1" height="320" border="0"></td>
					<td colspan="2" valign="top" style="font-family:arial; font-size:15px; text-align:justify; text-justify:inter-word;">
						<table border="0" width="625">
							<tr>
								<td>&nbsp;</td>
								<td colspan="3" style="font-family:arial; font-size:15px; text-align:justify; text-justify:inter-word;">
								Recebemos da empresa infra citada a quantia de <%=formatCurrency(valor_total_recibo)%><%'=formatCurrency(valor_total_recibo)%> (<%response.write(extenso(valor_total_recibo,1))%>) referente a locação de equipamentos e instrumentais para videocirurgias, no período de 01 a <%=ultimodiames(mes_procedimentos,ano_procedimentos)&" de "&mes_selecionado(mes_procedimentos)&" de "&ano_procedimentos%><%=complemento%>.
								<br><img src="../imagens/px.gif" alt="" width="1" height="40" border="0"></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td style="font-family:arial; font-size:15px; text-align:left;">Cliente:</td>
								<td>&nbsp;</td>
								<td style="font-family:arial; font-size:15px; text-align:left;"><b><%=nm_unidade_nome%></b></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td style="font-family:arial; font-size:15px; text-align:left;">Endereço:</td>
								<td>&nbsp;</td>
								<td style="font-family:arial; font-size:15px; text-align:left;"><%=nm_rua&" "&nm_numero%><%if nm_complemento <> "" Then response.write(" - "&nm_complemento)%></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td style="font-family:arial; font-size:15px; text-align:left;">CNPJ:</td>
								<td>&nbsp;</td>
								<td style="font-family:arial; font-size:15px; text-align:left;"><%=nm_cnpj%></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td colspan="3" style="font-family:arial; font-size:15px; text-align:justify; text-justify:inter-word;">
								<img src="../imagens/px.gif" alt="" width="1" height="40" border="0"><br>
								Este recibo só terá validade plena mediante o efetivo depósito bancário ou quitação do título (boleto) a favor da <%=nm_empresa%> - CNPJ: <%=cd_cnpj_emp%> - Banco <%=nm_banco_emp%> - Ag: <%=nm_agencia_emp%> - C/C: <%=nm_contacorrente_emp%>
								</td>
							</tr>
							<tr>
								<td><img src="../imagens/px.gif" alt="" width="32" height="1" border="0"></td>
								<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
								<td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
								<td><img src="../imagens/px.gif" alt="" width="480" height="1" border="0"></td>
							</tr>
						</table>
						</td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="100" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td>&nbsp;</td>
					<td align="center" style="font-family:arial; font-size:15px;"><b><%=nm_empresa%>.</b></td>
				</tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td>&nbsp;</td>
					<td align="center" style="font-family:arial; font-size:13px;"><b>CNPJ: <%=cd_cnpj_emp%></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="80" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="2" style="font-family:arial; font-size:13px;"><img src="../imagens/blackdot.gif" alt="" width="668" height="1" border="0"><br>
					Dispensado da emissão de Nota fiscal de acordo com a Lei Complementar n° 116 de 31 de julho de 2003, DOU de 01 de Agosto de 2003.</td>
				</tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td><img src="../imagens/px.gif" alt="" width="440" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="230" height="1" border="0"></td>
				</tr>
			</table>
			<%end if%>
			
		<%id_rua = ""
		id_numero_i = ""
		id_numero_f = ""
		nm_rua = ""
		nm_numero = ""
		salta_linha = 1
		cab_descritivo = 0
		
		cd_fatura = ""
		nm_qtd_procedimentos = ""
		nm_qtd_proc2 = ""
		nm_qtd_proc3 = ""
		nm_qtd_empr = ""
		nm_obs = ""
		
		
		valor_total_procedimentos = 0
		valor_total_proc2 = 0
		valor_total_proc3 = 0
		total_geral_procedimentos = 0
		
		total_procedimentos = 0
		valor_total = 0
		valor_total_recibo = 0
		
		rs.movenext
		wend
end if%>
