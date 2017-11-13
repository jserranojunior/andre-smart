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

	<table align="center" class="no_print">
	<form action="../financeiro.asp" name="seleciona_periodo" id="seleciona_período">
		<input type="hidden" name="tipo" value="fatura">
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		<tr>
			<td align="center" colspan="2">Emissão de Faturas</td>
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
		<tr>
			<td><select name="cd_unidade">
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
			
			</td></tr>
		</form>
	</table>
	<br class="no_print">
	<table align="center" border="1" style="border-collapse:collapse;" class="no_print">
		<form action="financeiro/acoes/acoes3.asp" method="post" name="fatura" id="fatura">
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		<input type="hidden" name="acao" value="fatura">
		<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
		<input type="hidden" name="dt_faturamento" value="<%=dt_mes&"/1/"&dt_ano%>">		
		
		<tr>
			<td align="center"> &nbsp; &nbsp; &nbsp; &nbsp; </td>
			<td align="center"><b>Unidades</b></td>
			<td align="center"><b>Venc.</b></td>
			<td align="center">&nbsp;&nbsp;</td>
			<td align="center"><b>Qtd.1</b></td>
			<td align="center"><b>qtd.2</b></td>
			<td align="center"><b>qtd.3</b></td>
			<td align="center"><b>Empr.</b></td>
			<td align="center"><b>Obs</b></td>
		</tr>
			<%num_linha = 1
			xsql = "SELECT * FROM TBL_unidades where cd_status = 1 and cd_hospital = 1"
			Set rs = dbconn.execute(xsql)%>
				<%while not rs.EOF
					strcd_unidade = rs("cd_codigo")
					strnm_unidade = rs("nm_unidade")
					strnm_sigla = rs("nm_sigla")
					dia_recebimento = rs("dia_recebimento")
						
						xsql = "SELECT * FROM TBL_financeiro_fatura3 where cd_unidade = "&strcd_unidade&" and dt_faturamento between '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
						Set rs_fatura = dbconn.execute(xsql)
							if not rs_fatura.EOF then
								cd_fatura = rs_fatura("cd_codigo")
								nm_qtd_procedimentos = rs_fatura("nm_qtd_procedimentos")
								nm_qtd_proc2 = rs_fatura("nm_qtd_proc2")
								nm_qtd_proc3 = rs_fatura("nm_qtd_proc3")
								nm_qtd_empr = rs_fatura("nm_qtd_empr")
								nm_obs = rs_fatura("nm_obs")
							end if
					%>
					<tr>
						<td align="center" valign="top"><%=num_linha%></td>
						<td valign="top"><a href="javascript:void(0);" onClick="window.open('ferramentas/unidade_cad.asp?cd_unidade=<%=strcd_unidade%>&status=1&acao=editar&jan=1','Altera_nome','width=520,height=400')" return false;"><%=strnm_unidade%></a><input type="hidden" name="cd_fatura_<%=strcd_unidade%>" value="<%=cd_fatura%>"></td>
						<td align="center" valign="top"><input type="text" name="dia_vencimento_<%=strcd_unidade%>" value="<%=dia_recebimento%>" size="2" maxlength="2"></td>
						<td>&nbsp;</td>
						<td align="center" valign="top"><input type="text" name="qtd_proc1_<%=strcd_unidade%>" size="3" maxlength="5" value="<%=nm_qtd_procedimentos%>"></td>
						<td align="center" valign="top"><%if strcd_unidade = 20 then%><input type="text" name="qtd_proc2_<%=strcd_unidade%>" size="3" maxlength="5" value="<%=nm_qtd_proc2%>"><%else%>&nbsp;<%end if%></td>
						<td align="center" valign="top"><%if strcd_unidade = 20 then%><input type="text" name="qtd_proc3_<%=strcd_unidade%>" size="3" maxlength="5" value="<%=nm_qtd_proc3%>"><%else%>&nbsp;<%end if%></td>
						<td align="center" valign="top"><input type="text" name="qtd_empr_<%=strcd_unidade%>" size="3" maxlength="5" value="<%=nm_qtd_empr%>"></td>
						<td align="left" valign="top"><textarea cols="30" rows="2" name="nm_obs_<%=strcd_unidade%>"><%=nm_obs%></textarea></td>
						
					</tr>
				<%num_linha = num_linha + 1
				
				rs.movenext
				wend%>
				<tr><td colspan="4"><input type="submit" value="Gravar"></td></tr>
		</form>
	</table>
	<%if cd_unidade = "0" OR cd_unidade = "" Then
		xsql = "SELECT * FROM View_financeiro_fatura3 where dt_faturamento between '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
	else
		xsql = "SELECT * FROM View_financeiro_fatura3 where cd_unidade = "&cd_unidade&" and dt_faturamento between '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
	end if	
		Set rs = dbconn.execute(xsql)
		
		
		while not rs.EOF
		cd_unidade = rs("cd_unidade")
		nm_unidade_nome = rs("nm_unidade_nome")
		nm_unidade = rs("nm_unidade")
		nm_sigla = rs("nm_sigla")
		dia_vencimento = rs("dia_vencimento")
		nm_contato = rs("nm_contato")
		nm_endereco = rs("nm_endereco")
			id_rua = instr(1,nm_endereco,"!",1)-1
			id_numero_i = instr(1,nm_endereco,"!",1)+1
			id_numero_f = instr(1,nm_endereco,"#",1)-1
			id_complemento_i = instr(1,nm_endereco,"#",1)+1
			id_complemento_f = instr(1,nm_endereco,"~",1)
			
				id_numero_f = id_numero_f - id_numero_i+1
			if id_rua > 0 then nm_rua = mid(nm_endereco,1,id_rua)&","
			if id_numero_f > 0 then nm_numero = mid(nm_endereco,id_numero_i,id_numero_f)
			if id_complemento_f > 0 then nm_complemento = mid(nm_endereco,id_complemento_i,id_complemento_f-id_complemento_i)
			
			nm_cnpj = rs("nm_cnpj")
		
		cd_prazo_faturamento = rs("cd_prazo_faturamento")
			
			if cd_prazo_faturamento <> "" Then
				mes_faturamento = dt_mes - 2'cd_prazo_faturamento
				ano_faturamento = dt_ano
					if mes_faturamento < 1 then
						mes_faturamento = 12
						ano_faturamento = dt_ano - 1
					end if
					dia_final_faturamento = ultimodiames(mes_faturamento,ano_faturamento)
			else
				mes_faturamento = mes_anterior
				ano_faturamento = ano_anterior
				dia_final_faturamento = dia_final_anterior
			end if
			
			nm_qtd_procedimentos = rs("nm_qtd_procedimentos")
			nm_qtd_proc2 = rs("nm_qtd_proc2")
			nm_qtd_proc3 = rs("nm_qtd_proc3")
			nm_qtd_empr = rs("nm_qtd_empr")
				
				
				'xsql = "SELECT * FROM TBL_unidades_valores_procedimentos where cd_unidade = "&cd_unidade&" and dt_atualizacao >= '"&dt_mes&"/1/"&dt_ano&"' order by dt_atualizacao, cd_tipo"
				'Set rs_valores = dbconn.execute(xsql)
						
						'São Luiz
						if cd_unidade = 2 or cd_unidade = 3 OR cd_unidade = 15 then
						'	while not rs_valores.EOF
								nm_valor1 = 320
						'	wend
						elseif cd_unidade = 11 then
							nm_valor1 = 327
						elseif cd_unidade = 14 then
							nm_valor1 = 330
						elseif cd_unidade = 19 then
							nm_valor1 = 340
						elseif cd_unidade = 22 then
							nm_valor1 = 21000
							nm_qtd_procedimentos = 1
							valor_total_procedimentos = nm_qtd_procedimentos * nm_valor1
						elseif cd_unidade =  20 then
							'if not rs_valores.EOF then
								'For i=1 to 3
							'		nm_valor1 =  rs_valores("nm_valor")
									nm_valor1 = 300
									nm_valor2 = 430
									nm_valor3 = 120
									'rs_valores.movenext
								'next
						'	end if
						
						end if
					
			'nm_valor_procedimento = 300
			'nm_valor_proc2 = 430
			'nm_valor_proc3 = 120
			
			'nm_valor_empr = 0
			
			nm_obs = rs("nm_obs")%>
			<%if cd_unidade = "" and salta_linha = 1 then %><br style="page-break-before:always;"><%end if%>
	<!-- ********************************************
	*********  1 - PROTOCOLO DE RECEBIMENTO *********
	**********************************************-->
			<table align="center" width="500" border="0" style="border-collapse:collapse;">
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="47" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td>&nbsp;</td>
					<td align="right" valign="top"><img src="../imagens/logotipo_vdlap.gif" alt="" width="170" height="118" border="0"></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="84" border="0"></td></tr><tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td align="center" colspan="2" style="text-decoration:none; font-family:arial; font-size:21px;"><b>Protocolo de Recebimento de Documentos</b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="90" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="2" valign="top" style="font-family:arial; font-size:18px; text-align:justify; text-justify:inter-word;">Recebi da <b>Vd Lap Cirúrgica Ltda</b>, o relatório dos procedimentos realizados por videocirurgia no mês de <%=mes_selecionado(mes_faturamento)&" de "&ano_faturamento%>, no <b><%=nm_unidade%></b>, junto com o recibo de pagamento da locação dos equipamentos de videocirurgia relativo ao mês com vencimento em <b><%=dia_vencimento%> de <%=mes_selecionado(dt_mes)&" de "&dt_ano%></b>.</td>
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
	<!-- ************************************************
	*********  2 - DEMONSTRATIVO DO FATURAMENTO *********
	**************************************************-->
			<%if cd_unidade <> 22 then%>
			<br style="page-break-after:always;">
			
			<table width="650" align="center" border="0" <%if cd_unidade <> 20 then%>style="background:url('../imagens/vdlap_logo_watermark.jpg') no-repeat;"><%end if%>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="46" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="1" style="font-family:arial; font-size:15px;">São Paulo, <%=dia_vencimento &" de "&mes_selecionado(dt_mes)&" de "&dt_ano%>.</td>
					<td align="right" valign="top" rowspan="4"><%if cd_unidade = 20 then%><img src="../imagens/logotipo_vdlap.gif" alt="" width="170" height="120" border="0"><%end if%>&nbsp;</td>
				</tr>
				<tr><td colspan="2"><img src="../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="1" style="font-family:arial; font-size:15px;"><b>Ao<br><%=nm_unidade_nome%><br>&nbsp;<br>A/C: <%=nm_contato%></b></td>
				</tr>
				<tr><td colspan="2"><img src="../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td align="center" colspan="2" style="text-decoration: underline; font-family:arial; font-size:17px;"><b> Faturamento de <%=mes_selecionado(mes_faturamento)&" "&ano_faturamento%></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
				<tr>
					<td><img src="../imagens/px.gif" alt="" width="1" height="450" border="0"><!-- limitador de linha --></td>
					<td colspan="2" valign="top" style="font-family:arial; font-size:16px; text-align:justify; text-justify:inter-word;" >
						Em contrato firmado, informamos o número de procedimentos realizados no centro cirurgico do <%=nm_unidade_nome%> no período de 01 a <%=ultimodiames(mes_faturamento,ano_faturamento)&" de "&mes_selecionado(mes_faturamento)&" de "&ano_faturamento%>, conforme tabela anexa.<br>
						<img src="../imagens/px.gif" alt="" width="1" height="40" border="0"><br>
							<%if cd_unidade = 20 then%>
								<table border="1" style="font-family:arial; font-size:15px; border-collapse:collapse;padding:50px 40px; background:transparent;" align="center" >
										<tr >
											<td align="center" style="padding:15px 5px;"><b>Descrição</b><br>
											<img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
											<td align="center" style="padding:15px 5px;"><b>N° de Procedimentos</b><br>
											<img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>			
											<td align="center" style="padding:15px 5px;"><b>Custo / Cirurgia</b><br>
											<img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
											<td align="center" style="padding:15px 5px;"><b>Valor Total</b><br>
											<img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
										</tr>
										<tr>
											<td align="center" style="padding:15px 5px;">Cirurgia com Rack Vd Lap</td>
											<td align="center" style="padding:15px 5px;"><%=nm_qtd_procedimentos%></td>
											<td align="center" style="padding:15px 5px;"><%=formatCurrency(nm_valor1)%></td>
											<%valor_total_procedimentos = nm_qtd_procedimentos * nm_valor1%>
											<td align="center" style="padding:15px 5px;"><%=formatCurrency(valor_total_procedimentos)%></td>
										</tr>
										<tr>
												<td align="center" style="padding:15px 5px;">Cirurgia com Rack e <br>instrumental Vd Lap</td>
												<td align="center" style="padding:15px 5px;"><%=nm_qtd_proc2%></td>
												<td align="center" style="padding:15px 5px;"><%=formatCurrency(nm_valor2)%></td>
												<%valor_total_proc2 = nm_qtd_proc2 * nm_valor2%>
												<td align="center" style="padding:15px 5px;"><%=formatCurrency(valor_total_proc2)%></td>
											</tr>
											<tr>
												<td align="center" style="padding:15px 5px;">Cirurgia realizadas com o<br>Rack do Hospital</td>
												<td align="center" style="padding:15px 5px;"><%=nm_qtd_proc3%></td>
												<td align="center" style="padding:15px 5px;"><%=formatCurrency(nm_valor3)%></td>
												<%valor_total_proc3 = nm_qtd_proc3 * nm_valor3%>
												<td align="center" style="padding:15px 5px;"><%=formatCurrency(valor_total_proc3)%></td>
											</tr>
										<%total_procs = nm_qtd_procedimentos + nm_qtd_proc2 + nm_qtd_proc3%>
										<tr>
											<td align="center" style="padding:15px 5px;"><b>Total</b></td>
											<td align="center" style="padding:15px 5px;"><%=total_procs%></td>
											<td>&nbsp;</td>
											<%valor_total_procedimentos = valor_total_procedimentos + valor_total_proc2 + valor_total_proc3%>
											<td align="center" style="padding:15px 5px;"><%=formatCurrency(valor_total_procedimentos)%></td>
										</tr>
								</table>
							<%else%>
								<table border="0" align="center" style="font-family:arial; font-size:15px; border-collapse:collapse;padding:50px 40px; background:transparent;">
										<tr>
											<td align="left" style="padding:15px 5px;"><b>Total Procedimento<%if nm_qtd_procedimentos > 1 then response.write("s")%></b></td>
											<td align="left" style="padding:15px 5px;"><b><%=nm_qtd_procedimentos%></b></td>
										</tr>
										<tr>
											<td align="left" style="padding:15px 5px;"><b>Valor Total</b></td>
											<%valor_total_procedimentos = nm_qtd_procedimentos * nm_valor1%>
											<td align="left" style="padding:15px 5px;"><b><%=formatCurrency(valor_total_procedimentos)%></b></td>
										</tr>										
										<tr>
											<td align="left" style="padding:15px 5px;"><b>Vencimento</b></td>
											<td align="left" style="padding:15px 5px;"><b><%=zero(dia_vencimento)&"/"&zero(dt_mes)&"/"&dt_ano%></b></td>
										</tr>
										<tr>
											<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
											<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
										</tr>
								</table>
							<%end if%>
							<img src="../imagens/px.gif" alt="" width="1" height="20" border="0"><br>
							<%=nm_obs%>
								<!--No mês de <%'=mes_selecionado(mes_anterior)%> não houve empréstimo de material Video Cirúrgico para outras equipes.--></td>
							
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="85" border="0"></td></tr>
				<tr>
					<td colspan="2">&nbsp;<!-- limitador de linha --></td>
					<td align="center" colspan="1" style="font-family:arial; font-size:15px;"><b>Vd Lap Cirúrgica Ltda<br><span style="font-family:arial; font-size:13px;">CNPJ: 00.255.640/0001-42</span></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="93" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td align="center" colspan="5" style="font-family:arial; font-size:13px;">
					<img src="../imagens/blackdot.gif" alt="" width="660" height="1" border="0"><br>Rua Dr. Afonso Vergueiro, 627 - Cep: 0211-001 - São Paulo-SP - Fone:(11)2631-4956</td>
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
			<br style="page-break-after:always;">
			<table width="500" border="0" align="center" style="border-collapse:collapse;">
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td style="font-family:arial; font-size:14px; text-align:left;"><b>Vd Lap Cirúrgica Ltda</b><br>
					Rua Dr. Afonso Vergueiro, 627<br>
					CEP: 02116-001 - São Paulo - SP<br>
					Tel/Fax: (11)2631-4959<br>
					CNPJ: 00.255.640/0001-42<br>
					vdlap@vdlap.com.br<br>
					www.vdlap.com.br</td>
					<td align="right"><img src="../imagens/logotipo_vdlap.gif" alt="" width="170" height="120" border="0"></td>
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
					<td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="10" border="0"></td></tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td colspan="2" style="font-family:arial; font-size:17px; text-align:center;"><b><%if valor_total_procedimentos <> "" Then response.write(formatCurrency(valor_total_procedimentos))%></b></td>
				</tr>
				<tr><td colspan="3"><img src="../imagens/px.gif" alt="" width="1" height="60" border="0"></td></tr>
				<tr>
					<td><img src="../imagens/px.gif" alt="" width="1" height="260" border="0"></td>
					<td colspan="2" valign="top" style="font-family:arial; font-size:15px; text-align:justify; text-justify:inter-word;">
						<table border="0" width="625">
							<tr>
								<td>&nbsp;</td>
								<td colspan="3" style="font-family:arial; font-size:15px; text-align:justify; text-justify:inter-word;">
								Recebemos da empresa infra citada a quantia de <%=formatCurrency(valor_total_procedimentos)%> (<%response.write(extenso(valor_total_procedimentos,1))%>) referente a locação de equipamentos e instrumentais para videocirurgias, no período de 01 a <%=ultimodiames(mes_faturamento,ano_faturamento)&" de "&mes_selecionado(mes_faturamento)&" de "&ano_faturamento%>
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
								<td style="font-family:arial; font-size:15px; text-align:left;"><%=nm_rua&" "&nm_numero&" - "&nm_complemento%></td>			
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
								Este recibo só terá validade plena mediante o efetivo depósito bancário a favor da Vd Lap Cirúrgica Ltda - CNPJ: 00.255.640/0001-42 - Banco Santander - Ag: 4342 - C/C: 1300389-9
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
					<td align="center" style="font-family:arial; font-size:15px;"><b>Vdlap Cirúrgica Ltda.</b></td>
				</tr>
				<tr>
					<td>&nbsp;<!-- limitador de linha --></td>
					<td>&nbsp;</td>
					<td align="center" style="font-family:arial; font-size:13px;"><b>CNPJ: 00.255.640/0001-42</b></td>
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
			
		<%id_rua = ""
		id_numero_i = ""
		id_numero_f = ""
		nm_rua = ""
		nm_numero = ""
		salta_linha = 1
		
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
		
		rs.movenext
		wend%>