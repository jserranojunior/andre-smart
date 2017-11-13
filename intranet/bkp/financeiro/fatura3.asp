<!--#include file="../includes/numero_extenso.asp"-->
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
		
		mes_posterior = dt_mes + 1
		ano_posterior = dt_ano
			if mes_posterior > 12 then
				mes_posterior = 1
				ano_posterior = dt_ano + 1
			end if
			dia_final_posterior = ultimodiames(mes_posterior,ano_posterior)

cd_unidade = request("cd_unidade")
	


cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)%>

	<table align="center" class="no_print" border="0" style="border:2px solid black; border-collapse:collapse;">
	<form action="../financeiro.asp" name="seleciona_periodo" id="seleciona_período">
		<input type="hidden" name="tipo" value="fatura">
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		<tr style="background-color:gray; color:white; font-size:14px;">
			<td align="center" colspan="2">Emissão de Faturas</td>
		</tr>
		<tr bgcolor="#c0c0c0">
			<td>&nbsp;<b>M&ecirc;s</b></td>
			<td><b>Ano</b></td>
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
		<tr bgcolor="#c0c0c0">
			<td colspan="2">&nbsp;<input type="submit" name="OK" value="Seleciona o mês de faturamento" width="4" height="5" style="background-color:orange;"></td>
		</tr>
		</form>
	</table>
	<br class="no_print">
	<br class="no_print">
	<br class="no_print">
	<table align="center" border="1" style="border-collapse:collapse;" class="no_print" width="100">
		<form action="financeiro/acoes/acoes3.asp" method="post" name="fatura" id="fatura">
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		<input type="hidden" name="acao" value="fatura">
		<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
		<input type="hidden" name="dt_faturamento" value="<%=dt_mes&"/1/"&dt_ano%>">
		<input type="hidden" name="name="dia_vencimento_<%=strcd_unidade%>"" value="<%=dia_vencimento%>">	
		
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
					
						'*** PREFERENCIAS DO UNIDADE/FATURAMENTO ***
						xsql = "SELECT * FROM TBL_unidade_faturamento_preferencias where cd_unidade = "&strcd_unidade&""
						Set rs_pref = dbconn.execute(xsql)
						
						if not rs_pref.EOF Then
							nm_dias_faturamento = rs_pref("nm_dias_faturamento")
							nm_dias_faturamento2 = rs_pref("nm_dias_faturamento2")
							cd_valor_fixo = rs_pref("cd_valor_fixo")
						end if%>
					<tr>
						<td align="center" valign="top" rowspan="1"><%=zero(num_linha)%><br><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
						<td valign="top" rowspan="1"><a href="javascript:void(0);" onClick="window.open('ferramentas/unidade_cad.asp?cd_unidade=<%=strcd_unidade%>&status=1&acao=editar&jan=1','Altera_nome','width=520,height=440,location=0,status=0,scrollbars=1')" return false;"><b><%=strnm_unidade%></b></a><br>
						<img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
						<td>
							<table align="left" border="1" style="border-collapse:collapse;border:1px solid gray;">
								<%'*** TIPO DE PROCEDIMENTO ***
								xsql = "SELECT * FROM View_unidades_procedimento_tipo where cd_unidade = "&strcd_unidade&""
								SET rs_proc = dbconn.execute(xsql)
									while not rs_proc.EOF
										cd_procedimento_tipo = rs_proc("cd_procedimento_tipo")
										nm_procedimento_tipo = rs_proc("nm_procedimento_tipo_abrev")
										
								if cab = 0 then%>
								<td align="center"><b>&nbsp;</b></td>
								<td align="center"><b>Num</b></td>
								<td align="center"><b>Valor</b></td>
								<td align="center"><b>Total</b></td>
								<td align="center"><b>Venc.</b></td>
								<%end if%>
								<input type="hidden" name="cd_tipo_proc_<%=strcd_unidade%>" value="<%=cd_procedimento_tipo%>">
								<%'*** INFORMAÇÕES SOBRE A FATURA ***
								'* Procedimentos *
								strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE (cd_unidade = "&strcd_unidade&") AND (cd_tipo_proc = "&cd_procedimento_tipo&") AND (dt_faturamento = '"&dt_mes&"/1/"&dt_ano&"')"
								SET rs_qtd_proc = dbconn.execute(strsql)
									if not rs_qtd_proc.EOF then
										nm_qtd_procedimentos = rs_qtd_proc("nm_qtd_procedimentos")
									end if
								%>
								<tr>
									<td><%=nm_procedimento_tipo%></td>
									<td align="center">
										<%if cd_valor_fixo = 1 then%>
											<b>Fixo</b>
											<input type="hidden" name="qtd_proc_<%=strcd_unidade%>_<%=cd_procedimento_tipo%>" value="1">
										<%else%>
											<input type="text" name="qtd_proc_<%=strcd_unidade%>_<%=cd_procedimento_tipo%>" size="3" maxlength="5" value="<%=nm_qtd_procedimentos%>">
										<%end if%></td>
									<td align="right"><%'*** VALOR DO PROCEDIMENTO ***
									strsql = "SELECT * FROM TBL_unidades_valores_procedimentos WHERE (cd_unidade = "&strcd_unidade&") AND (dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"') AND cd_procedimento_tipo = "&cd_procedimento_tipo&""
										SET rs_valor = dbconn.execute(strsql)
										'while not rs_proc.EOF
										if not rs_valor.EOF then
											nm_valor = rs_valor("nm_valor")%>
										<%=formatnumber(nm_valor)%>
										<%end if%>
										</td>
										<%if nm_qtd_procedimentos <> "" then nm_total_unidade = nm_qtd_procedimentos * nm_valor%>
										<td align="right">&nbsp;<%=formatnumber(nm_total_unidade)%></td>
										
										
										<td align="center"><%if cab = 0 then%>											
												<%if nm_dias_faturamento > 29 then
													dia_recebimento = nm_dias_faturamento - 30
													mes_recebimento = dt_mes + 1
														if mes_recebimento > 12 then
															mes_recebimento = 1
															ano_recebimento = dt_ano + 1
														end if
												else
													mes_recebimento = dt_mes
													ano_recebimento = dt_ano
												end if%>
												<%=zero(dia_recebimento)&"/"&zero(mes_recebimento)%>
												<input type="hidden" name="dt_recebimento_<%=strcd_unidade&"_"&cab%>" value="<%=mes_recebimento&"/"&dia_recebimento&"/"&dt_ano%>">												
											<%else%>
												&nbsp;
											<%end if%></td>
								</tr>
								<%cab = cab + 1
								nm_qtd_procedimentos = ""
								nm_qtd_empr = ""
								rs_proc.movenext
								wend
								
								'* Empréstimos *
								strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE (cd_unidade = "&strcd_unidade&") AND (cd_tipo_proc = 5) AND (dt_faturamento = '"&dt_mes&"/1/"&dt_ano&"')"
								SET rs_qtd_proc = dbconn.execute(strsql)
								if not rs_qtd_proc.EOF then
									nm_qtd_empr = rs_qtd_proc("nm_qtd_procedimentos")
								end if%>
								<tr>
									<td>Emprés.</td>
									<td align="center"><input type="text" name="qtd_empr_<%=strcd_unidade%>" size="3" maxlength="5" value="<%=nm_qtd_empr%>"></td>
									<td colspan="3">&nbsp;</td>
								</tr>
								<tr>
									<td><img src="../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
									<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
									<td><img src="../imagens/blackdot.gif" alt="" width="60" height="1" border="0"></td>
									<td><img src="../imagens/blackdot.gif" alt="" width="70" height="1" border="0"></td>
									<td><img src="../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
								</tr>
							</table>							
						</td>
						
						
					</tr>
				<%cab = 0
				num_linha = num_linha + 1
				rs.movenext
				wend
				
				%>
				<tr><td colspan="4"><input type="submit" value="Gravar"></td></tr>
		</form>
	</table>
	
	
	
	
	<%if cd_unidade = "0" OR cd_unidade = "" Then
		xsql = "SELECT DISTINCT cd_unidade AS unid_faturamento FROM TBL_financeiro_fatura3"
	else
		xsql = "SELECT DISTINCT cd_unidade AS unid_faturamento FROM TBL_financeiro_fatura3 where cd_unidade = 20"
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
				
				nm_banco_emp = rs_unid("nm_banco_emp")
				nm_agencia_emp = rs_unid("nm_agencia_emp")
				nm_contacorrente_emp = rs_unid("nm_contacorrente_emp")
			
		'	*** Cálculo para vencimento de pagamento ***
			xsql = "SELECT * FROM TBL_unidade_faturamento_preferencias where cd_unidade = "&cod_unidade&""
			Set rs_pref = dbconn.execute(xsql)
				if not rs_pref.EOF Then
					nm_dias_faturamento = rs_pref("nm_dias_faturamento")
					nm_dias_faturamento2 = rs_pref("nm_dias_faturamento2")
					
					cd_mostra_proc = rs_pref("cd_mostra_proc")
					cd_mostra_emp = rs_pref("cd_mostra_emp")
					cd_mostra_proc_duplo = rs_pref("cd_mostra_proc_duplo")
					
					cd_impress_pg1 = rs_pref("cd_impress_pg1")
					cd_impress_pg2 = rs_pref("cd_impress_pg2")
					cd_impress_pg3 = rs_pref("cd_impress_pg3")
					
						if nm_dias_faturamento > 30 then
							dia_vencimento = nm_dias_faturamento - 30
							mes_vencimento = dt_mes + 1
							ano_vencimento = dt_ano
								if mes_vencimento > 12 then
									mes_vencimento = 1
									ano_vencimento = dt_ano + 1
								end if
						else
							dia_vencimento = nm_dias_faturamento
							mes_vencimento = dt_mes
							ano_vencimento = dt_ano
						end if
				end if
		'		
		'	
		'	nm_qtd_procedimentos = rs("nm_qtd_procedimentos")
		'	nm_qtd_proc2 = rs("nm_qtd_proc2")
		'	nm_qtd_proc3 = rs("nm_qtd_proc3")
		'	nm_qtd_empr = rs("nm_qtd_empr")
		'	
			
		'	
			
		'		
				
		'		xsql = "SELECT * FROM TBL_unidades_valores_procedimentos where cd_unidade = "&cd_unidade&" and dt_atualizacao >= '"&dt_mes&"/1/"&dt_ano&"' order by dt_atualizacao, cd_procedimento_tipo"
		'		Set rs_valores = dbconn.execute(xsql)
						
		'		'		'São Luiz
		'				if cd_unidade = 2 or cd_unidade = 3 OR cd_unidade = 15 then
		'		'		'	while not rs_valores.EOF
		'						nm_valor1 = 320
		'		'		'	wend
		'				elseif cd_unidade = 11 then
		'					nm_valor1 = 327
		'				elseif cd_unidade = 14 then
		'					nm_valor1 = 330
		'				elseif cd_unidade = 19 then
		'					nm_valor1 = 340
		'				elseif cd_unidade = 22 then
		'					nm_valor1 = 21000
		'					nm_qtd_procedimentos = 1
		'					valor_total_procedimentos = nm_qtd_procedimentos * nm_valor1
		'				elseif cd_unidade =  20 then
		'		'			'if not rs_valores.EOF then
		'		'				'For i=1 to 3
		'		'			'		nm_valor1 =  rs_valores("nm_valor")
		''							nm_valor1 = 300
		'							nm_valor2 = 430
		'							nm_valor3 = 120
		'		'					'rs_valores.movenext
		'		'				'next
		'		'		'	end if
		'				end if
		'				'while not rs_valores.EOF 
		'					'rs_valor1 = rs_valores("nm_valor")
		'				'rs_valores.movenext
		'				'wend
		'	'nm_valor_procedimento = 300
		'	'nm_valor_proc2 = 430
		'	'nm_valor_proc3 = 120
		'	
		'	'nm_valor_empr = 0
		'	
		'	nm_obs = rs("nm_obs")%>
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
					<td colspan="2" valign="top" style="font-family:arial; font-size:18px; text-align:justify; text-justify:inter-word;">Recebi da <b><%=nm_empresa%></b>, o relatório dos procedimentos realizados por videocirurgia no mês de <%=mes_selecionado(mes_anterior)&" de "&ano_anterior%>, no <b><%=nm_unidade_nome%></b>, junto com o recibo de pagamento da locação dos equipamentos de videocirurgia relativo ao mês com vencimento em <b><%=dia_vencimento%> <%if dia_vencimento2 <> "" Then response.write(" e "&dia_vencimento2)%> de <%=mes_selecionado(mes_vencimento)&" de "&ano_vencimento%></b>.</td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="130" border="0"></td></tr>
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
					<td colspan="1" style="font-family:arial; font-size:15px;">São Paulo, <%=dia_vencimento &" de "&mes_selecionado(mes_vencimento)&" de "&ano_vencimento%>.</td>
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
					<td align="center" colspan="2" style="text-decoration: underline; font-family:arial; font-size:17px;"><b> Faturamento de <%=mes_selecionado(mes_anterior)&" "&ano_anterior%></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
				<tr>
					<td><img src="../imagens/px.gif" alt="" width="1" height="450" border="0"><!-- limitador de linha --></td>
					<td colspan="2" valign="top" style="font-family:arial; font-size:16px; text-align:justify; text-justify:inter-word;" >
						Em contrato firmado, informamos o número de procedimentos realizados no centro cirurgico do <%=nm_unidade_nome%> no período de 01 a <%=ultimodiames(mes_anterior,ano_anterior)&" de "&mes_selecionado(mes_anterior)&" de "&ano_anterior%>, conforme tabela anexa.<br>
						<img src="../imagens/px.gif" alt="" width="1" height="40" border="0"><br>
								<table border="1" width="500" style="font-family:arial; font-size:15px; border-collapse:collapse;padding:5px 4px; background:transparent;" align="center">
										<%strsql = "SELECT * FROM VIEW_unidades_fatura_valores_procedimentos WHERE (cd_unidade = "&cod_unidade&") AND (cd_tipo_proc < 5)"
										SET rs_proc = dbconn.execute(strsql)
											while not rs_proc.EOF
												cd_tipo_proc = rs_proc("cd_tipo_proc")
												nm_procedimento_tipo = rs_proc("nm_procedimento_tipo")
												nm_qtd_procedimentos = rs_proc("nm_qtd_procedimentos")
												
												strsql = "SELECT * FROM TBL_unidades_valores_procedimentos WHERE (cd_unidade = "&cod_unidade&")"
												SET rs_valor = dbconn.execute(strsql)
													if not rs_valor.EOF then
														nm_valor_procedimento = rs_valor("nm_valor")
													rs_valor.movenext
													end if
												
												if cab_descritivo = 0 then%>
												<tr >
													<td align="center" style="padding:15px 5px;"><%if cd_tipo_proc > 1 then%> <%larg_cel = "180"%><b>Descrição</b><%else%><%larg_cel = "110"%>&nbsp;<%end if%><br>
													<img src="../imagens/px.gif" alt="" width="<%=larg_cel%>" height="1" border="0"></td>
													<td align="center" style="padding:15px 5px;"><b>N° de Procedimentos</b><br>
													<img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>			
													<td align="center" style="padding:15px 5px;"><b>Custo / Cirurgia</b><br>
													<img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
													<td align="center" style="padding:15px 5px;"><b>Valor Total</b><br>
													<img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
												</tr>
												<%cab_descritivo = cab_descritivo + 1
												end if%>
												<tr>
													<td align="center" style="padding:15px 5px;"><%if cd_tipo_proc <>  1 then%><%=nm_procedimento_tipo%><%else%>&nbsp;<%end if%></td>
													<td align="center" style="padding:15px 5px;"><%=nm_qtd_procedimentos%></td>
													<td align="center" style="padding:15px 5px;"><%=formatCurrency(nm_valor_procedimento)%></td>
													<%valor_total_procedimentos = nm_qtd_procedimentos * nm_valor_procedimento%>
													<td align="center" style="padding:15px 5px;"><%=formatCurrency(valor_total_procedimentos)%></td>
												</tr>
											
											<%total_procedimentos = total_procedimentos + nm_qtd_procedimentos
											valor_total = valor_total + valor_total_procedimentos
											rs_proc.movenext
											wend%>
										
										<%%>
										<tr>
											<td align="center" style="padding:15px 5px;"><b>Total</b></td>
											<td align="center" style="padding:15px 5px;"><%=total_procedimentos%></td>
											<td>&nbsp;</td>
											<td align="center" style="padding:15px 5px;"><%=formatCurrency(valor_total)%></td>
										</tr>
										<tr>
											<td align="center" style="padding:15px 5px;">Vencimento</td>
											<td align="center" style="padding:15px 5px;" colspan="3"><%=dia_vencimento&"/"&mes_vencimento&"/"&ano_vencimento%></td>
										</tr>
								</table>
								
							
							<img src="../imagens/px.gif" alt="" width="1" height="40" border="0"><br>
							<%'strsql = "SELECT * FROM VIEW_unidades_fatura_valores_procedimentos WHERE (cd_unidade = "&cod_unidade&") AND (cd_tipo_proc < 5)"
							strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE cd_unidade="&cod_unidade&" AND cd_tipo_proc >= 5"
							SET rs_mens = dbconn.execute(strsql)
								while not rs_mens.EOF
									cd_tipo_proc = rs_mens("cd_tipo_proc")
									nm_qtd_procedimentos = rs_mens("nm_qtd_procedimentos")%>
								
									<%if cd_mostra_emp = 1 then
										if cd_tipo_proc = 5 AND nm_qtd_procedimentos > 0 Then%>										
											No mês de <%=mes_selecionado(mes_anterior)%> houve <%=nm_qtd_procedimentos%> empréstimo de material Video Cirúrgico para outras equipes.
										<%else%>
											No mês de <%=mes_selecionado(mes_anterior)%> não houve empréstimo de material Video Cirúrgico para outras equipes.
										<%end if
									end if%>
									<%if cd_mostra_proc_duplo = 1 then
										'*** Mensagem para procedimentos duplos ***
										if cd_tipo_proc = 6 AND nm_qtd_procedimentos > 0 Then%>										
											No mês de <%=mes_selecionado(mes_anterior)%> houve <%'=nm_qtd_procedimentos%> pacientes com procedimento duplo.
										<%else%>
											No mês de <%=mes_selecionado(mes_anterior)%> não houve pacientes com procedimento duplo.
										<%end if
										
										'*** Mensagem para empréstimos
										if cd_tipo_proc = 5 AND nm_qtd_procedimentos > 0 Then%>										
											No mês de <%=mes_selecionado(mes_anterior)%> houve <%=nm_qtd_procedimentos%> empréstimo de material Video Cirúrgico para outras equipes.
										<%else%>
											No mês de <%=mes_selecionado(mes_anterior)%> não houve empréstimo de material Video Cirúrgico para outras equipes.
										<%end if
									end if%>
								<%rs_mens.movenext
								wend%>
								</td>
							
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="85" border="0"></td></tr>
				<tr>
					<td colspan="2">&nbsp;<!-- limitador de linha --></td>
					<td align="center" colspan="1" style="font-family:arial; font-size:15px;"><b><%=nm_empresa%><br><span style="font-family:arial; font-size:13px;">CNPJ: <%=cd_cnpj_emp%></span></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="93" border="0"></td></tr>
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
					Tel/Fax: (11)<%=telefone_emp%><br>
					CNPJ: <%=cd_cnpj_emp%><br>
					<%=nm_email_emp%><br>
					<%=nm_site_emp%></td>
					<td align="right"><img src="../imagens/logotipo<%=cd_cliente_empresa%>.gif" alt="" width="170" height="120" border="0"></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="80" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="2" style="font-family:arial; font-size:15px; text-align:right;">São Paulo, <%=zero(dia_vencimento)&" de "&mes_selecionado(dt_mes)&" de "&dt_ano%>.</td>
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
				<%strsql = "SELECT * FROM VIEW_unidades_fatura_valores_procedimentos WHERE (cd_unidade = "&cod_unidade&") AND (cd_tipo_proc < 5)"
										SET rs_proc = dbconn.execute(strsql)
											while not rs_proc.EOF
												cd_tipo_proc = rs_proc("cd_tipo_proc")
												nm_procedimento_tipo = rs_proc("nm_procedimento_tipo")
												nm_qtd_procedimentos = rs_proc("nm_qtd_procedimentos")
												
												strsql = "SELECT * FROM TBL_unidades_valores_procedimentos WHERE (cd_unidade = "&cod_unidade&")"
												SET rs_valor = dbconn.execute(strsql)
													if not rs_valor.EOF then
														nm_valor_procedimento = rs_valor("nm_valor")
													rs_valor.movenext
													end if%>
												
													
													
													<%valor_total_procedimentos = nm_qtd_procedimentos * nm_valor_procedimento%>
													
												
											
											<%total_procedimentos = total_procedimentos + nm_qtd_procedimentos
											valor_total = valor_total + valor_total_procedimentos
											rs_proc.movenext
											wend%><tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="2" style="font-family:arial; font-size:17px; text-align:center;"><b><%if valor_total <> "" Then response.write(formatCurrency(valor_total))%></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="60" border="0"></td></tr>
				<tr>
					<td><img src="../imagens/px.gif" alt="" width="1" height="260" border="0"></td>
					<td colspan="2" valign="top" style="font-family:arial; font-size:15px; text-align:justify; text-justify:inter-word;">
						<table border="0" width="625">
							<tr>
								<td>&nbsp;</td>
								<td colspan="3" style="font-family:arial; font-size:15px; text-align:justify; text-justify:inter-word;">
								Recebemos da empresa infra citada a quantia de <%=formatCurrency(valor_total)%> (<%response.write(extenso(valor_total,1))%>) referente a locação de equipamentos e instrumentais para videocirurgias, no período de 01 a <%=ultimodiames(mes_anterior,ano_anterior)&" de "&mes_selecionado(mes_anterior)&" de "&ano_anterior%>
								<br><img src="../imagens/px.gif" alt="" width="1" height="40" border="0"></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td style="font-family:arial; font-size:15px; text-align:left;">Cliente:</td>
								<td>&nbsp;</td>
								<td style="font-family:arial; font-size:15px; text-align:left;"><b><%=nm_unidade%></b></td>
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
								Este recibo só terá validade plena mediante o efetivo depósito bancário a favor da <%=nm_empresa%> - CNPJ: <%=cd_cnpj_emp%> - Banco <%=nm_banco_emp%> - Ag: <%=nm_agencia_emp%> - C/C: <%=nm_contacorrente_emp%>
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
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="130" border="0"></td></tr>
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
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="100" border="0"></td></tr>
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
		
		rs.movenext
		wend%>