<!--include file="../includes/numero_extenso.asp"-->
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../css/geral.htm"-->

<%dt_ano = request("ano_sel")
dt_mes = request("mes_sel")
	'if dt_mes = "" OR dt_ano = "" then
	if dt_ano = "" then
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

cd_unidade = request("cd_unidade")
cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)
dt_hoje_numerica = year(now)&zero(month(now))&zero(day(now))%>


<form action="../financeiro/fatura_consolidado3.asp" name="seleciona_periodo" method="post" id="seleciona_período">

		<input type="hidden" name="tipo" value="consolidado">		
	<table align="center" class="no_print" border="0" style="border:2px solid black; border-collapse:collapse;">
		<tr style="background-color:gray; color:white; font-size:14px;">
			<td align="center" colspan="3">Consolidado - <%=dt_ano%></td>
		</tr>
		
		<tr bgcolor="#c0c0c0">
			<td align="center">Ano:&nbsp;<input type="text" name="ano_sel" maxlength="4" size="4" value="<%=dt_ano%>"></td>
		<!--/tr>
		<tr bgcolor="#c0c0c0">
			<td colspan="2">&nbsp;<select name="cd_unidade">
					<option value="0">&nbsp;</option>
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
		<tr align="left" bgcolor="#c0c0c0"-->
			<td colspan="2"><input type="submit" name="OK" value="Ok" width="4" height="5" style="background-color:orange;">&nbsp;&nbsp;&nbsp;</td>
		</tr>
		</form>
	</table>
	<!--br class="no_print"-->

<table border="1" class="nok_print" style="border-collapse:collapse">
	<tr bgcolor="#e9e9e9">
		<%if int(dt_ano) < year(now) then
			abrangencia_mes2 = 12
		else
			abrangencia_mes2 = month(now)+1
				if abrangencia_mes2 > 12 then
					abrangencia_mes2 = 12
				end if
		end if
		total_recebidos_atrasados = 0
		total_atrasados = 0
		
		linha = 0
		for i = 1 to abrangencia_mes2
			dt_mes = i%>
		
			<td align="center" style="font-size:13px; color:red;"><b><%=mes_selecionado(i)%></b><br>
				<table align="center" border="1" style="border-collapse:collapse;" class="nok_print" width="100">
		
			<%num_linha = 1
			cab = 0
			
			'*** LISTA AS UNIDADES ***
			if cd_unidade = "0" OR cd_unidade = "" Then
				xsql = "SELECT * FROM TBL_unidades where cd_status = 1 and cd_hospital = 1 AND cd_codigo <> 43"
			else
				xsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" AND cd_status = 1 and cd_hospital = 1 AND cd_codigo <> 43"
			end if
			
			Set rs = dbconn.execute(xsql)%>
				<%while not rs.EOF
					strcd_unidade = rs("cd_codigo")
					strnm_unidade = rs("nm_unidade")
					strnm_sigla = rs("nm_sigla")
					
						'*** PREFERENCIAS DO UNIDADE/FATURAMENTO ***
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
						<td align="center" valign="middle" rowspan="1"><%=zero(num_linha)%><br><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
						<td align="center" valign="middle" rowspan="1"><a href="javascript:void(0);" onClick="window.open('ferramentas/unidade_cad.asp?cd_unidade=<%=strcd_unidade%>&status=1&acao=editar&jan=1','Altera_nome','width=520,height=440,location=0,status=0,scrollbars=1')" return false;"><b><%=strnm_sigla%></b></a><br>
						<!--fat: <b><%=nm_dias_faturamento%></b> 	<%if nm_dias_faturamento2 <> "" Then response.write(" e <b>"&nm_dias_faturamento2)%></b><br>
						ven: <b><%=nm_dia_vencimento1%></b> <%if nm_dia_vencimento2 <> "" Then response.write(" e <b>"&nm_dia_vencimento2)%></b><br-->
						
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
							<b><%'=left(mes_selecionado(mes_procedimentos),3)&"/"&ano_procedimentos%></b>
							
						<img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
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
								<td align="center" class="ok_print"><b>&nbsp;</b><br><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<td align="center"><b>Valor</b><br><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
								<td align="center" class="ok_prints"><b>Qtd.</b><br><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Total</b><br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
								<td align="center"><b>Venc.</b><br><img src="../imagens/px.gif" alt="" width="65" height="1" border="0"></td>
								<td align="center"><b>Conf.</b><br><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
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
										<%if isnull(cd_rack_unid) then
											cond_rack = " "
										else
											cond_rack = " and a056_codrac="&cd_rack_unid&" "
										end if
										
										'**** MOSTRA A QUANTIDADE DE CIRURGIAS ******************************
										sql_proc = "SELECT COUNT (a002_numpro) AS conta ,a053_codung, nm_sigla,nm_sigla,a056_codrac FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_procedimentos&"/1/"&ano_procedimentos&"' AND '"&mes_procedimentos&"/"&dia_final_procedimentos&"/"&ano_procedimentos&"' AND a053_codung = "&strcd_unidade&" "&cond_rack&" AND a002_tipage <> 'E' Group by a053_codung, nm_sigla,nm_sigla,a056_codrac"
										SET rs_procs = dbconn.execute(sql_proc)
											while not rs_procs.EOF
												qtd_proc_auto = rs_procs("conta")
												cd_rack = rs_procs("a056_codrac")
											'	response.write(cd_rack&"-"&qtd_proc_auto&"<br>")
												total_procs = total_procs + qtd_proc_auto
											rs_procs.movenext
											wend%>
											<%'=total_procs%>
											
								<%if nm_dia_vencimento2 <> "" Then
									qtd_venc = 2
									qtd_total_procs = total_procs / 2
								else
									qtd_venc = 1
									qtd_total_procs = total_procs
								end if
								
								for i_venc = 1 to qtd_venc%>
								<tr class="ok_prints">
									<td class="ok_print"><%=nm_procedimento_tipo%></td>
									<%'*** VALOR DO PROCEDIMENTO ***
									strsql = "SELECT * FROM TBL_unidades_valores_procedimentos WHERE (cd_unidade = "&strcd_unidade&") AND (dt_atualizacao <= '"&i&"/1/"&dt_ano&"') AND cd_procedimento_tipo = "&cd_procedimento_tipo&" order by dt_atualizacao desc"
									SET rs_valor = dbconn.execute(strsql)
										if not rs_valor.EOF then
											nm_valor = rs_valor("nm_valor")%>										
										<%end if%>
										<%'if nm_qtd_procedimentos <> "" then nm_total_unidade = nm_qtd_procedimentos * nm_valor
											'if nm_total_unidade = 0  AND dt_ano >= year(now) AND dt_mes >= month(now) then
											data_numerica_atual = year(now)&zero(month(now))
											data_numerica_proc = ano_procedimentos&zero(mes_procedimentos)
											
											if cd_valor_fixo = 1 then
												cor_alerta = "#ffffff"
												qtd_total_procs = 0
											elseif data_numerica_atual = data_numerica_proc then
												cor_alerta = "#ffffae"
												qtd_total_procs = 0
											else
												cor_alerta = "#ffffff"
											end if%>
										<td align="right"><%if nm_valor > 0 then response.write(formatnumber(nm_valor))%></td>
										<td align="center" class="ok_prints" bgcolor="<%=cor_alerta%>"><%=venc2%><%'=data_numerica_atual&"~"&data_numerica_proc%>
											<%if cd_valor_fixo = 1 then
												'if dt_hoje_numerica <= 20140731 then
													qtd_total_procs = 1
												'else
												'	qtd_total_procs = 0
												'end if%>
												<b>Fixo</b>
												<%'=cd_fatura%>
											<%else%>
												<%if isnull(cd_rack_unid) then
													cond_rack = " "
												else
													cond_rack = " and a056_codrac="&cd_rack_unid&" "
												end if%>
												<%=qtd_total_procs%>
											<%end if%>
												
										</td>
										<td align="right" bgcolor="<%=cor_alerta%>">&nbsp;<%if qtd_total_procs <> 0 then response.write(formatnumber(qtd_total_procs*nm_valor))%></td>
										<%total_geral = total_geral + qtd_total_procs*nm_valor%>
										
										<td align="center">&nbsp;<%if cab = 0 then
																		if i_venc = 1 then
																			'response.write(dt_recebimento)
																			response.write(zero(nm_dia_vencimento1)&"/"&zero(i))
																			dia_vencimento = nm_dia_vencimento1
																		elseif i_venc = 2 then
																			response.write(zero(nm_dia_vencimento2)&"/"&zero(i))
																			dia_vencimento = nm_dia_vencimento2
																		end if
																	end if%></td>
										<td align="center">&nbsp;<%'if cab = 0 then response.write(dt_recebimento)%><%if qtd_venc = 1 OR qtd_venc = 2 then%><a href="javascript:void(0);" return false;" onClick="window.open('fatura_recebimento_adiamento.asp?cd_unidade=<%=strcd_unidade%>&dt_dia=<%=dia_vencimento%>&dt_mes=<%=i%>&dt_ano=<%=dt_ano%>&nm_valor=<%=qtd_total_procs*nm_valor%>','adiar','location=0,status=0,width=300,height=200,scrollbars=1resizable')"><img src="../imagens/ic_valor_reprov.gif" alt="" width="12" height="12" border="0"></a><%end if%></td>
								</tr>
								<%next%>
								<%cab = cab + 1
								nm_qtd_procedimentos = ""
								nm_qtd_empr = ""
								dt_vencimento = ""
								total_procs = 0
								
								'venc2 = ""
								rs_proc.movenext
								wend
								
								%>
								<%strsql = "SELECT * FROM TBL_financeiro_fatura_desc WHERE (cd_unidade = "&strcd_unidade&") AND (dt_data = '"&dt_mes&"/1/"&dt_ano&"')"
										SET rs_desc = dbconn.execute(strsql)
										if not rs_desc.EOF then
											cd_desc = rs_desc("cd_codigo")
											nr_desconto = rs_desc("nr_desconto")
												nr_desconto_porcent = nr_desconto/100
												valor_desconto = total_geral*nr_desconto_porcent
												total_geral = total_geral - valor_desconto
												%>
								<tr>
									<td align="center"><b>Desc.</b></td>
									<td align="right"><%=nr_desconto%>%</td>
									<td align="right"><%=formatnumber(valor_desconto,2)%></td>
								</tr>
								<%end if%>
								<tr>
									<td class="ok_print">&nbsp;</td>
									<td align="center" class="ok_prints"><b>Total</b></td>
									<td align="center"><%if qtd_venc = 2 Then%><%=qtd_total_procs*2%><%else%><%'=qtd_total_procs%><%end if%></td>
									<td align="right" style="font-size:10px;" bgcolor="<%=cor_alerta%>"><b><%if total_geral > 0 then response.write(formatnumber(total_geral,2))%></b></td>
									<td align="center"><%'=dt_recebimento%><%'=data_numerica_atual&"<br>"&data_numerica_proc%></td>
									<td align="center">&nbsp;<%if cab = 0 then response.write(dt_recebimento)%></td>
								</tr>
							</table>							
						</td>
						
						
					</tr>
				<%cab = 0
				nm_qtd_procedimentos = 0
				'total_procs = 0
				qtd_total_procs = 0
				venc2 = ""
				
				nr_total_mes = nr_total_mes + total_geral
				total_geral = 0
				
				num_linha = num_linha + 1
				rs.movenext
				wend
				
				%>
				<tr>
					<td align="center" colspan="2"><b>A receber</b></td>
					<td align="right"><b><%if nr_total_mes > 0 then response.write(formatnumber(nr_total_mes,2))%></b><img src="../imagens/px.gif" alt="" width="113" height="1" border="0"></td>
				</tr>
				
				<%'*** VALORES NÃO RECEBIDOS ***
				strsql = "SELECT * FROM TBL_financeiro_fatura_adia_pgto WHERE dt_vencimento_ori BETWEEN '"&i&"/1/"&dt_ano&"' AND '"&i&"/"&ultimodiames(i,dt_ano)&"/"&dt_ano&"'"
				SET rs_atrasados = dbconn.execute(strsql)
					while not rs_atrasados.EOF
						nm_valor_atrasado = rs_atrasados("nm_valor")
						total_atrasados = int(total_atrasados+nm_valor_atrasado)
					rs_atrasados.movenext
					wend%>
				<tr>
					<td align="center" colspan="2"><b>Não recebido</b></td>
					<td align="right"><b><%if total_atrasados <> "" Then response.write(formatnumber(total_atrasados,2))%></b><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"><img src="../imagens/ic_editar.gif" alt="" width="13" height="14" border="0" onClick="window.open('fatura_lista_valores.asp?dt_mes=<%=i%>&dt_ano=<%=dt_ano%>&modo=atrasados','lista','location=0,status=0,width=300,height=200,scrollbars=1,resizable=0')"></td>
				</tr>
				
				<%'*** VALORES RECEBIDOS EM ATRASO ***
				strsql = "SELECT * FROM TBL_financeiro_fatura_adia_pgto WHERE dt_vencimento_novo BETWEEN '"&i&"/1/"&dt_ano&"' AND '"&i&"/"&ultimodiames(i,dt_ano)&"/"&dt_ano&"'"
				SET rs_rec_atrasados = dbconn.execute(strsql)
					while not rs_rec_atrasados.EOF
						nm_valor_recebido_atrasado = rs_rec_atrasados("nm_valor")
						total_recebidos_atrasados = int(total_recebidos_atrasados+nm_valor_recebido_atrasado)
					rs_rec_atrasados.movenext
					wend%>
				<tr>
					<td align="center" colspan="2"><b>Rec. atraso</b></td>
					<td align="right"><b><%if total_recebidos_atrasados <> "" then response.write(formatnumber(total_recebidos_atrasados,2))%></b><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"><img src="../imagens/ic_editar.gif" alt="" width="13" height="14" border="0" onClick="window.open('fatura_lista_valores.asp?dt_mes=<%=i%>&dt_ano=<%=dt_ano%>&modo=receb_atraso','lista','location=0,status=0,width=300,height=200,scrollbars=1,resizable=0')"></td>
				</tr>
				
				<%'*** VALORES RECEBIDOS TOTAL (total a receber - total atrasado + Total atrasado anterior) ***
					total_recebido_mes = nr_total_mes - total_atrasados + total_recebidos_atrasados%>
				<tr>
					<td align="center" colspan="2"><b>Recebido no mês</b></td>
					<td align="right" style="font-size:11px;"><b><%if total_recebido_mes <> "" Then response.write(formatcurrency(total_recebido_mes,2))%></b><img src="../imagens/px.gif" alt="" width="115" height="1" border="0"></td>
				</tr>
		</form>
	</table>
			</td>
			</tr><tr>
			<%if i < abrangencia_mes2 then%><td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td><%end if%>
		
		<%nr_total_mes = 0
		total_recebidos_atrasados = 0
		nm_valor_recebido_atrasado = 0
		total_atrasados = 0
		nm_valor_atrasado = 0
		next%>
		</tr>
</table>