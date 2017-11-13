<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%
cd_user = session("cd_codigo")

strcod = request("cod")
strcod = ""
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")

dt_dia = request("dt_dia")
dt_mes = request("dt_mes")
dt_ano = request("dt_ano")
	'if dt_dia = "" Then dt_dia = day(now) 'end if
	'if dt_mes = "" Then dt_mes = month(now) 'end if
	'if dt_ano = "" Then dt_ano = year(now) 'end if
	
	if dt_dia = "" AND  dt_mes = "" AND dt_ano = "" Then
		dt_dia = day(now)
		dt_mes = month(now)
		dt_ano = year(now)
			if dt_dia > 15 then
				dt_dia = day(now)
				dt_mes = month(now)+1
				dt_ano = year(now)
					if dt_mes = 13 then
						dt_mes = 1
						dt_ano = year(now)+1
					end if
						if dt_dia > ultimodiames(dt_mes,dt_ano) then
							dt_dia = ultimodiames(dt_mes,dt_ano)
						end if
			end if
	end if
	
dia_final = ultimodiames(dt_mes,dt_ano)
	
	mesanoatual=year(now)&zero(month(now))
	mesanosel=dt_ano&zero(dt_mes)
		
	mes_competencia = dt_mes - 1
		if mes_competencia < 1 then
			mes_competencia = 12
			ano_competencia = dt_ano - 1
		else
			ano_competencia = dt_ano
		end if
		final_competencia = ultimodiames(mes_competencia,ano_competencia)
	
'*** Quantidade de dias da competencia ***
	qtd_dias_comp = datediff("d","1/"&mes_competencia&"/"&ano_competencia,final_competencia&"/"&mes_competencia&"/"&ano_competencia)+1
	'qtd_dias_comp = datediff("d","1/9/2011","30/9/2011")

'*** Quantidade de domingos da competencia ***
	qtd_semanas_comp = datediff("ww","1/"&mes_competencia&"/"&ano_competencia,final_competencia&"/"&mes_competencia&"/"&ano_competencia,1)
	qtd_domingos_comp = datediff("ww","1/"&mes_competencia&"/"&ano_competencia,final_competencia&"/"&mes_competencia&"/"&ano_competencia,1)
	
	
	for i = 1 to dia_final
		var_dia_semana = var_dia_semana&","&i
	next
	
'response.write("<br>"&dt_mes&"1/"&dt_ano&"<br>"&dia_final&"/"&dt_mes&"/"&dt_ano)

			if ordem_total <> "" Then
				ordem_funcionarios = ordem_total
			else
				'ordem_funcionarios = "cd_contrato,nm_nome,nm_sigla"
				ordem_funcionarios = "cd_contrato,nm_nome"
			end if
			
data_atual = final_competencia&"/"&mes_competencia&"/"&ano_competencia%>
<!--#include file="../../includes/feriados.asp"-->
<%qtd_dias_dsr = qtd_domingos_comp + qtd_feriados - qtd_feriado_dsr%>
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print">
						
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="empresa.asp"  method="POST">
							<input type="hidden" name="tipo" value="folha">
							<input type="hidden" name="cod" value="<%=strcod%>">
							<tr>
								<td colspan="4" align="center" class="cabecalho" style="background-color:black; color:white;"><%'=data_atual%> <b>ADM - FOLHA DE PAGAMENTO <%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td align="center" style="background-color:silver;"><b>Mês para pagamento:</b></td>
								<td align="center" style="background-color:silver;" colspan="3">
									<select name="dt_mes">
										<%for i = 1 to 12%>
											<option value="<%=i%>" <%if i = int(dt_mes) then%>SELECTED<%end if%>><%=mesdoano(i)%></option>
										<%next%>
									</select> / <input type="text" name="dt_ano" value="<%=dt_ano%>" size="4" maxlength="4">
									<input type="submit" name="ok" value="Ver"></td>
							</tr>
							<tr>
								<td align="center" style="background-color:silver;">Mês de competência:</td>
								<td align="center" style="background-color:silver;"><%=mesdoano(mes_competencia)%>/<%=ano_competencia%></td>
							</tr>
							<tr>
								<td><img src="../../imagens/blackdot.gif" alt="" width="150" height="0" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="150" height="0" border="0"></td>
							</tr>
						</form>
					</table>	
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" style="border-collapse:collapse;">	
			<!-- ---------------------------------------------------------------- -->				
							<tr style="color:white;" class="no_print">
								<td align="left" colspan="32">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês &nbsp; &nbsp; &nbsp; &nbsp; (<%=(qtd_dias_comp-qtd_domingos_comp)-qtd_feriados%>)
									{<%=var_dia_semana%>}</td>
							</tr>							
							<%'**************************
							'*** VALOR SALÁRIO MÍNIMO ***
							xsql = "SELECT * FROM TBL_funcionario_valores_padroes WHERE (cd_tipo = 12) AND dt_atualizacao <= '"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"' ORDER BY dt_atualizacao DESC"
							Set rs = dbconn.execute(xsql)
							nr_salminimo = rs("nr_valor")
							'****************************
							'*** VALOR Aux. Creche	  ***
							xsql = "SELECT * FROM TBL_funcionario_valores_padroes WHERE (cd_tipo = 6) AND dt_atualizacao <= '"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"' ORDER BY dt_atualizacao DESC"
							Set rs = dbconn.execute(xsql)
							nr_auxcreche = rs("nr_valor")
							'******************************

					num_cor = 1
					num_linha = 1
					class_print = ""'ok_print"
					'xsql = "up_funcionario_rh_lista_individual @cd_funcionario="&strcod&", @dt_atualizacao='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
					'xsql = "up_funcionario_folhapagamento_geral @dt_pagamento_i='"&mes_pagamento&"/1/"&ano_pagamento&"', @dt_pagamento_f='"&mes_pagamento&"/"&final_pagamento&"/"&ano_pagamento&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
					'xsql = "up_funcionario_folhapagamento_geral @dt_pagamento_i='"&mes_pagamento&"/1/"&ano_pagamento&"', @dt_pagamento_f='"&mes_pagamento&"/"&final_pagamento&"/"&ano_pagamento&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
					xsql = "up_funcionario_folhapagamento_geral @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&mes_competencia&"/1/"&ano_competencia&"', @dt_f = '"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
						Set rs = dbconn.execute(xsql)
							while not rs.EOF
							
								strcod = rs("cd_codigo")
								cd_matricula = rs("cd_matricula")
								nm_nome = rs("nm_nome")
								cd_unidade = rs("cd_unidade")
								nm_unidade = rs("nm_unidade")
								nm_sigla = rs("nm_sigla")
								cd_hospital = rs("cd_hospital")
								cd_sexo = rs("cd_sexo")
								
								dt_admissao = rs("dt_admissao")
								dt_demissao = rs("dt_demissao")
								
								admissao = rs("admissao")
								demissao = rs("demissao")
									if demissao <> "null" then
										cor_registro = "red"
									else
										cor_registro = "black"
										demissao = ""
									end if
								
								if admissao = year(now)&zero(month(now)) then
								'	dias_trab_admissao = datediff(m,dt_admissao,now)
								end if
									
									'*** Caso a admissão for no mesmo mês de competencia da folha ***
									if admissao = dt_ano&zero(mes_competencia)  AND demissao = "" OR admissao = ano_competencia&zero(mes_competencia) AND demissao > ano_competencia&zero(mes_competencia) then
										tempo_trabalhado = DateDiff("d",day(dt_admissao)&"/"&month(dt_admissao)&"/"&year(dt_admissao),final_competencia&"/"&mes_competencia&"/"&ano_competencia)+1 '(1 = conta o dia da contratação)
									
									'*** Caso a admissão e a demissão forem no mesmo mês de competencia da folha ***
									elseif admissao = dt_ano&zero(dt_mes)  AND demissao = dt_ano&zero(dt_mes) then
										tempo_trabalhado = DateDiff("d",day(dt_admissao)&"/"&month(dt_admissao)&"/"&year(dt_admissao),final_competencia&"/"&mes_competencia&"/"&ano_competencia)+1 '(1 = conta o dia da contratação)
									
									else
										tempo_trabalhado = "30"
									
									'*** Caso a demissão for no mesmo mês de competencia da folha ***
									'elseif  admissao < dt_ano&zero(dt_mes)  AND demissao = dt_ano&zero(dt_mes) then 'OR admissao < dt_ano&zero(dt_mes) AND demissao > dt_ano&zero(dt_mes) then
									'elseif  admissao < dt_ano&zero(dt_mes)  AND month(dt_demissao)&"/"&day(dt_demissao)&"/"&year(dt_demissao) <= zero(dt_mes)&"/25/"&dt_ano then 'OR admissao < dt_ano&zero(dt_mes) AND demissao > dt_ano&zero(dt_mes) then
									'	tempo_trabalhado = DateDiff("d","1/"&mes_pagamento&"/"&ano_pagamento,day(dt_demissao)&"/"&month(dt_demissao)&"/"&year(dt_demissao))
									
									'else
									'	tempo_trabalhado = "0"
									end if
								
								cd_salario = rs("cd_salario")
								nr_salario = abs(rs("nr_salario"))
									if not IsNumeric(nr_salario) Then 
										nr_salario = 0
									else
										nr_salario = nr_salario
									end if
									if nr_salario <> 0 Then
										nr_salario_dia = nr_salario / 30
									else
										nr_salario = nr_salario
									end if
								
								'*** Valor Salário em caso de contratação no mesmo mês ***
								if tempo_trabalhado < "30" then
									nr_salario = nr_salario_dia * tempo_trabalhado
								end if
								
								expediente = rs("expediente")
									expediente = expediente / 60
										if expediente >= 9 AND expediente <= 11 then
											tipo_expediente = "8h"
											carga_mensal = 220
											carga_diaria = 8
										elseif expediente = 6 then
											tipo_expediente = "6h"
											carga_mensal = 180
											carga_diaria = 6
										elseif expediente = 12 then
											tipo_expediente = "12x36h"
											carga_mensal = 180
											carga_diaria = 12
										end if
									if nr_salario <> "0" AND carga_diaria <> "0" Then 
										'nr_salario_hora = nr_salario_dia / carga_diaria
									else
										nr_salario_hora = "0"
									end if
								
								cd_aj_custo = rs("cd_aj_custo")
								nr_aj_custo = rs("nr_aj_custo")
									if not IsNumeric(nr_aj_custo) Then nr_aj_custo = "0"
								
								cd_insalubridade_forca = rs("cd_insalubridade_forca")
								nr_insalubridade_forca = rs("nr_insalubridade_forca")
									if  not IsNumeric(nr_insalubridade_forca) Then nr_insalubridade_forca = "0"
								
								cd_ad_noturno = rs("cd_ad_noturno")
								qtd_dia_noturno = rs("qtd_dia_noturno")
									if  not IsNumeric(qtd_dia_noturno) Then qtd_dia_noturno = "0"
								
								cd_hextra = rs("cd_hextra")
								qtd_hextra = rs("qtd_hextra")
									if qtd_hextra > 36 then
										qtd_hextra = "36"
										he_marca = "1"
									end if
								
								cd_vrefeic_cancela = rs("cd_vrefeic_cancela")
								nr_vrefeic_cancela = rs("nr_vrefeic_cancela")
									if  not IsNumeric(nr_vrefeic_cancela) Then nr_vrefeic_cancela = "0"
								
								cd_vtransp_cancela = rs("cd_vtransp_cancela")
								nr_vtransp_cancela = rs("nr_vtransp_cancela")
									if  not IsNumeric(nr_vtransp_cancela) Then nr_vtransp_cancela = "0"
								
								cd_conv_medico = rs("cd_conv_medico")
								nr_conv_medico = rs("nr_conv_medico")
									if not IsNumeric(nr_conv_medico) then nr_conv_medico = "0"
								
								cd_contr_sind = rs("cd_contr_sind")
								nr_contr_sind = rs("nr_contr_sind")
									if  not IsNumeric(nr_contr_sind) Then nr_contr_sind = "0"
								
								cd_desc_diversos = rs("cd_desc_diversos")
								nr_desc_diversos = rs("nr_desc_diversos")
									if  not IsNumeric(nr_desc_diversos) Then nr_desc_diversos = "0"
								
								cd_faltas = rs("cd_faltas")
								nr_faltas = rs("nr_faltas")
									if  not IsNumeric(nr_faltas) Then nr_faltas = "0"
								
								cd_dsr_faltas = rs("cd_dsr_faltas")
								nr_dsr_faltas = rs("nr_dsr_faltas")
									if  not IsNumeric(nr_dsr_faltas) Then nr_dsr_faltas = "0"
								
								cd_atrasos = rs("cd_atrasos")
								nr_atrasos = rs("nr_atrasos")
									if nr_atrasos <> "" Then 
										nr_desc_atrasos = nr_atrasos * nr_salario_hora
										nr_salario_liq = nr_salario_liq - nr_desc_atrasos
										'response.write(nr_salario_liq)
									end if
									if  not IsNumeric(nr_atrasos) Then nr_atrasos = "0"
								
								'*** Salário desconto faltas ****
									if IsNumeric(nr_faltas) then
									'if abs(nr_faltas) >= "1" then
										total_faltas_mes = abs(nr_faltas) + abs(nr_dsr_faltas)
									'	nr_desc_salario = nr_salario_dia * nr_faltas
									'	nr_salario_desc = nr_salario - nr_desc_salario
									'	
									'	nr_desc_salario_dsr = nr_salario_dia * total_faltas_mes
									'	nr_salario_desc_dsr = nr_salario - nr_desc_salario_dsr
										nr_valor_falta = nr_salario_dia * total_faltas_mes
										nr_salario_liq = nr_salario - nr_valor_falta
									else
										'nr_salario_desc = nr_salario
										nr_salario_liq = nr_salario
									end if
								'********************************
									
									'if nr_faltas <> "0" AND nr_dsr_faltas <> "0" then 
									'	total_faltas_mes = abs(nr_faltas) + abs(nr_dsr_faltas)
									'	nr_desc_salario = nr_salario_dia * total_faltas_mes
									'	nr_salario_desc = nr_salario - nr_desc_salario
									'else
									'	nr_salario_desc = nr_salario
									'end if
									
								
								cd_credito_he = rs("cd_credito_he")
								nr_credito_he = rs("nr_credito_he")
									if  not IsNumeric(nr_credito_he) Then nr_credito_he = "0"
								
								cd_desconto_he = rs("cd_desconto_he")
								nr_desconto_he = rs("nr_desconto_he")
									if  not IsNumeric(nr_desconto_he) Then nr_desconto_he = "0"
								
								num_funcionario = num_funcionario + 1
									if num_cor = 1 then
										cor_linha = "#ccffff"
									else
										cor_linha = "#ffffff"
									end if
									
							'**** F É R I A S ***
							'xsql = "SELECT * FROM TBL_funcionario_ausencias WHERE cd_funcionario = "&strcod&" AND cd_motivo = 6 and dt_afastamento_i >= '"&dt_mes&"/1/"&dt_ano&"' and dt_afastamento_f <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
							'Set rs_ferias = dbconn.execute(xsql)
							'	while not rs_ferias.EOF
							'		
							'		periodo_ferias = datediff("d",rs_ferias("dt_afastamento_i"),rs_ferias("dt_afastamento_f"))
							'			if periodo_ferias <> "" Then 
							'				periodo_ferias = periodo_ferias + 1
							'				desconto_ferias = nr_salario_dia * periodo_ferias
							'			end if
							'	rs_ferias.movenext
							'	wend
							'		if periodo_ferias <> "" Then 
							'			nr_desconto_ferias = nr_salario - desconto_ferias
							'			nr_salario = nr_salario - nr_desconto_ferias
							'		end if
									
							'****** VALORES PADRÕES ******
							'insalubridade = nr_salminimo * 0.1
							'	if nr_faltas <> "0" Then 
							'		insalubridade_tempo_trab = (insalubridade / 30) * tempo_trabalhado
							'		insalubridade_falta = (insalubridade / 30) * nr_faltas
							'			insalubridade = insalubridade - insalubridade_tempo_trab
							'			insalubridade = insalubridade - insalubridade_falta
							'	end if
							
							if num_linha = 1 then%>
								<tr class="<%=class_print%>">
									<td colspan="36" align="center" class="cabecalho" style="background-color:black; color:white;"><b>ADM - FOLHA DE PAGAMENTO - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
									<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
								</tr>
								<tr style="background-color:gray; color:white;" class="<%=class_print%>">
									<td>&nbsp;</td>
									<td align="center"><b>Mat.</b></td>
									<td align="center"><b>Funcion&aacute;rio</b></td>
									<td align="center" class="no_print"><b>Unid.</b></td>
									<td align="center" class="no_print"><b>Carga Horária</b></td>
									<!--td align="center"><b>$/hora</b></td-->
									<td align="center"><b>Sal&aacute;rio<br>Bruto</b></td>
									<!--td align="center"><b>$/dia</b></td-->								
									<!--td align="center"><b>Sal&aacute;rio<br>Desc.</b></td-->
									<td align="center"><b>Aj. Custo</b></td>
									<td align="center"><b>Insalub.</b></td>
									<td align="center" class="no_print"><b>Qtd Not</b></td>
									<td align="center"><b>Ad. Not</b></td>
									<td align="center"><b>DSR <br>Ad. Not.</b></td>
									<!--td align="center" class="no_print"><b>Depend. <br>< 21anos</b></td-->
									<td align="center" class="no_print"><b>Qtd Dep</b></td>
									<td align="center"><b>Aux. Creche</b></td>
									<td align="center" class="no_print"><b>Qtd H.E.</b></td>
									<td align="center"><b>Valor H.E.</b></td>
									<td align="center"><b>DSR H.E.</b></td>
									<td align="center" class="no_print"><b>Total<br>Venc.</b></td>
									
									<td align="center" style="border-left:black solid 2px;"><b>Qtd.<br> faltas</b></td>								
									<td align="center"><b>Qtd.<br>DSR</b></td>
									<td align="center"><b>$ falta<br>+ DSR</b></td>
									
									<td align="center"><b>Qtd.<br>Atrasos</b></td>
									<td align="center"><b>Valor<br>Atrasos</b></td>
									<!--td align="center"><b>Desc. Férias</b></td-->
									<td align="center"><b>Cesta Básica</b></td>
									<td align="center"><b>Vale Refei.</b></td>
									<td align="center"><b>Vale Transp</b></td>
									<td align="center" class="no_print"><b>Base INSS</b></td>
									<td align="center"><b>INSS</b></td>
									<td align="center"><b>Conv. Méd.</b></td>
									<td align="center"><b>Contr. Sind.</b></td>
									<td align="center"><b>Descontos Diversos</b></td>
									<!--td align="center" class="no_print"><b>FGTS</b></td-->
									<td align="center" class="no_print"><b>Base IRRF</b></td>
									<td align="center" class="no_print"><b>Faixa IRRF</b></td>
									<td align="center" class="no_print"><b>Parc. a<br>deduzir</b></td>
									<td align="center"><b>IRRF</b></td>
									<td align="center" class="no_print"><b>Total<br>Desc.</b></td>
									<td align="center"><b>Total Liquido</b></td>
								</tr>
							<%end if%>
							
							
							<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');" bgcolor="<%=cor_linha%>" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_cad.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','','width=470,height=560,scrollbars=1')">								
								<td align="right"><b><%=zero(num_funcionario)%></b>&nbsp;</td>
								<td align="right" onClick="window.open('empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&cod=<%=strcod%>','','width=630,height=600,scrollbars=1')"><%=cd_matricula%>&nbsp;</td>
								<td style="color:<%'=cor_registro%>black;" align="left"><%=nm_nome%> <%'=tempo_trabalhado%><%'=" - "&dt_demissao%></td>
								<td class="no_print" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_demonstrativo.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','','width=600,height=600,scrollbars=1')"><%=nm_sigla%><%'=nm_unidade%></td>
								<td class="no_print"><%=tipo_expediente%></td>
								<!--td align="right"><%=formatnumber(nr_salario_hora,2)%></td-->
								<td align="right"><%=formatnumber(nr_salario,2)%></td>
								<!--td align="right"><%=formatnumber(nr_salario_dia,2)%></td>
								<td align="right"><%'=formatnumber(nr_salario_dia,2)%></td-->
								
								
								<!--td align="right">&nbsp;<%=formatnumber(nr_salario_desc,2)%></td-->
								<td align="right">&nbsp;<%=nr_aj_custo%><%'=dt_admissao%></td>
					<!-- *** Insalubridade *** -->
								<td align="right"><%'*** verifica se deve forçar o pagamento de insalubridade ***
													if nr_insalubridade_forca = 1 then cd_hospital = 1
							
														if cd_hospital = 1 AND nr_salario > 0 then
															insalubridade = nr_salminimo * 0.1
																if tempo_trabalhado < "30" Then
																	insalubridade_tempo_trab = (insalubridade / 30) * tempo_trabalhado
																		insalubridade = insalubridade_tempo_trab
																end if
																'if nr_faltas <> "0" Then
																'	insalubridade_falta = (insalubridade / 30) * nr_faltas
																'		insalubridade = insalubridade - insalubridade_falta
																'end if
															response.write(formatnumber(insalubridade,2))
														else
															response.write("0")
														end if%></td>
					<!-- *** Qtd. Adicional Noturno *** -->
								<td align="right" class="no_print">&nbsp;<%=zero(qtd_dia_noturno)%><!-- <%if qtd_dia_noturno = 1 then%> dia<%elseif qtd_dia_noturno > 1 then%>dias<%end if%>--></td>
					<!-- *** DSR Adicional Noturno *** -->
								<td align="right"><%if qtd_dia_noturno > 0 then
														ad_noturno = (((nr_salario*0.40)/carga_mensal)*7)*qtd_dia_noturno
														qtd_dias_mes = qtd_dias_comp-qtd_dias_dsr
														dsr_ad_not = ad_noturno/qtd_dias_mes
															dsr_ad_not = formatnumber(dsr_ad_not*qtd_dias_dsr,2)
														
														response.write(formatnumber(ad_noturno,2))
													else
														response.write("0")
													end if%></td>
								<td align="right">&nbsp;<%=dsr_ad_not%></td>
					<!-- *** Qtd. Dependentes menores de 21 anos *** -->
								<%xsql_dep = ("SELECT COUNT(nm_nome) AS qtd_dependentes FROM View_dependentes WHERE  (DATEDIFF(m, dt_nascimento, '"&dt_mes&"/"&dia_final&"/"&dt_ano&"') <= 251) AND (cd_funcionario = "&strcod&")")
									Set rs_dep = dbconn.execute(xsql_dep)
									if not rs_dep.EOF then
										qtd_dependentes_21 = rs_dep("qtd_dependentes")
									end if%>
								<!--td align="right" class="no_print">&nbsp;<%'if qtd_dependentes_21 = 1 then%> <%'=zero(qtd_dependentes_21)%> filho <%'elseif qtd_dependentes_21 > 1 then%><%'=zero(qtd_dependentes_21)%> filhos<%'end if%></td-->
					<!-- *** Qtd. Dependentes *** -->
								<%idade = ""
								xsql_dep = ("SELECT * FROM View_dependentes WHERE (DATEDIFF(m, dt_nascimento, '"&dt_mes&"/"&dia_final&"/"&dt_ano&"') <= 83) AND dt_nascimento <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"' AND (cd_funcionario = "&strcod&")")
								Set rs_dep = dbconn.execute(xsql_dep)
								while not rs_dep.EOF
									cd_depend = rs_dep("cd_codigo")
									nm_nome = rs_dep("nm_nome")
									dt_nascimento = rs_dep("dt_nascimento")
										idade = DateDiff("m",day(dt_nascimento)&"/"&month(dt_nascimento)&"/"&year(dt_nascimento),dia_final&"/"&dt_mes&"/"&dt_ano)
										'cd_idade = idade/12
									'	idade_a = int(cd_idade)
									'	idade_virg = instr(1,cd_idade,",",1)
									'		if idade_virg <> 0 then idade_m = round(12*(0&","&mid(cd_idade,idade_virg+1,len(idade))))&"m"
										qtd_dep_idade = qtd_dep_idade + 1
										'*** informa se o dependente está apto a receber o Aux. creche ***
										xsql_auxcreche = "SELECT top 1 * FROM View_funcionario_valores WHERE dt_atualizacao <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'  AND cd_funcionario="&strcod&"  AND cd_tipo=6 AND cd_indice= "&cd_depend&" ORDER BY dt_atualizacao desc"
											Set rs_auxcreche = dbconn.execute(xsql_auxcreche)
											if not rs_auxcreche.EOF then
												cd_auxcreche = rs_auxcreche("cd_codigo")
												ok_auxcreche = rs_auxcreche("nr_valor")
											end if
											if ok_auxcreche = 1 then
												qtd_dep_auxcreche = qtd_dep_auxcreche + 1
											end if%>
								<%rs_dep.movenext
								wend%>
								<td align="right" class="no_print"><%if qtd_dep_idade = 1 then%> <%=zero(qtd_dep_idade)%> filho <%elseif qtd_dep_idade > 1 then%><%=zero(qtd_dep_idade)%> filhos<%else%>&nbsp;<%end if%> <%'=formataidade(idade)%></td>
					<!-- *** Auxílio Creche *** -->
								<td align="right">&nbsp;<%if qtd_dep_auxcreche > 0  AND nr_salario > 0 Then%>
															<%auxcreche = qtd_dep_auxcreche*nr_auxcreche%>
															<%=formatnumber(auxcreche,2)%>														
														<%end if%>
														<%'=qtd_mes_dep%></td>
					<!-- *** Qtd Hora Extra *** -->
								<td align="right" class="no_print"><%if qtd_hextra >= 1 then%><%=zero(qtd_hextra)%>h<%end if%>&nbsp;</td>
					<!-- *** DSR hora Extra *** -->
								<td align="right">
									<%if qtd_hextra > 0 then
										hora_extra = ((nr_salario/carga_mensal)*1.9)*qtd_hextra
										qtd_dias_mes = qtd_dias_comp-qtd_dias_dsr
										dsr_he = hora_extra/qtd_dias_mes
											dsr_he = formatnumber(dsr_he*qtd_dias_dsr,2)
										
										response.write(formatnumber(hora_extra,2))
									else
										response.write("0")
									end if%></td>
								<td align="right">&nbsp;<%=dsr_he%></td>
					<!-- *** Vencimentos *** -->
							<%'total_vencimento = abs(nr_salario_liq) + abs(nr_aj_custo) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(auxcreche) + abs(hora_extra) + abs(dsr_he)
							total_vencimento = abs(nr_salario) + abs(nr_aj_custo) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(hora_extra) + abs(dsr_he) + abs(auxcreche)
							total_desconto = abs(nr_valor_falta) + abs(nr_desc_atrasos) + abs(cesta_b) + abs(valor_refeicao) + abs(v_trasnsp) + abs(inss) + abs(nr_conv_medico) + abs(nr_contr_sind)+ abs(nr_desc_diversos)%>							
								<td align="right" class="no_print">&nbsp;<b><%if total_vencimento <> "" Then response.write(formatnumber(total_vencimento,2))%></b></td>
								
	<!--*********************************************** D E S C O N T O S ********************************************************************************-->
					<!-- *** Faltas *** -->
								<td align="right" style="border-left:gray solid 2px;"><%=nr_faltas%></td>
								<td align="right"><%=nr_dsr_faltas%></td>
								<td align="right"><%if nr_faltas <> "" Then response.write(FormatNumber(nr_valor_falta,2))%></td>	
								
					<!-- *** Atrasos *** -->
								<td align="right"><%=nr_atrasos%></td>
								<td align="right">&nbsp;<%if nr_desc_atrasos <> "" Then response.write(FormatNumber(nr_desc_atrasos,2))%></td>
					
					<!-- *** Desc. Férias *** -->
								<!--td align="right">&nbsp;<%=nr_desconto_ferias%></td-->
					<!-- *** Cesta Básica *** -->
								<td align="right">&nbsp;<%xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=10 AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
											Set rs_refeic = dbconn.execute(xsql)
											if not rs_refeic.EOF AND nr_salario > 0 then
												cesta_b = rs_refeic("nr_valor")
												response.write(formatnumber(cesta_b,2))
											end if%></td>
					<!-- *** Vale Refeição *** -->
								<td align="right">&nbsp;<%if int(expediente) > 6 AND nr_salario > 0 then
														xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=2 AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
															Set rs_refeic = dbconn.execute(xsql)
															if not rs_refeic.EOF then
																valor_refeicao = rs_refeic("nr_valor")
																'response.write(formatnumber(valor_refeicao))
															'else
																'valor_refeicao = "0"
																'response.write(valor_refeicao)
															end if
													'else
														'valor_refeicao = "0"
														'response.write(valor_refeicao)
													end if%>
													
													<%if nr_salario <> "" AND nr_vrefeic_cancela = 0 Then
														response.write(valor_refeicao)
													else
														response.write("0")
														valor_refeicao = "0"
													end if%>
													</td>
					<!-- *** Vale Transporte *** -->
								<td align="right"><%if nr_salario > "0" AND nr_vtransp_cancela <> 1 Then
														v_trasnsp = nr_salario*0.06
														response.write(formatnumber(v_trasnsp))
													else
														response.write("0")
													end if%></td>
					<!-- *** Base INSS *** -->
								<td align="right" class="no_print">
								<%'base_inss = abs(nr_salario_liq) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(hora_extra) + abs(dsr_he)' - abs(nr_desc_atrasos)
									base_inss = abs(nr_salario) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(hora_extra) + abs(dsr_he)
										base_inss = base_inss - abs(nr_valor_falta)
										base_inss = base_inss - abs(nr_desc_atrasos)
									if not base_inss = 0 then response.write(formatnumber(base_inss,2))%></td>
					<!-- *** INSS *** -->
								<td align="right">&nbsp;<%xsql = "Select * From TBL_inss Where dt_atualizacao <= '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
												Set rs_inss = dbconn.execute(xsql)
												if not rs_inss.EOF then
													faixa_1 = rs_inss("pri_faixa")
													pri_aliq = rs_inss("pri_aliq")
													
													faixa_2 = rs_inss("seg_faixa")
													seg_aliq = rs_inss("seg_aliq")
													
													faixa_3 = rs_inss("ter_faixa")
													ter_aliq = rs_inss("ter_aliq")
												end if
												
												if base_inss <= abs(faixa_1) then
													inss = base_inss * pri_aliq
													response.write(formatnumber(inss,2))
													'response.write("1")
												elseif base_inss <= abs(faixa_2) then
													inss = base_inss * seg_aliq
													response.write(formatnumber(inss,2))
													'response.write("2")
												elseif base_inss <= abs(faixa_3) then
													inss = base_inss * ter_aliq
													response.write(formatnumber(inss,2))
													'response.write("3")
												end if%></td>
					<!-- *** Convênio Médico *** -->
								<td align="right">&nbsp;<%=nr_conv_medico%></td>
					<!-- *** Contribuição Sindical *** -->
								<td align="right">&nbsp;<%=nr_contr_sind%></td>
					<!-- *** Descontos Diversos *** -->
								<td align="right">&nbsp;<%=nr_desc_diversos%></td>
					<!-- *** FGTS *** -->
								<!--td align="right" class="no_print">&nbsp;<%fgts = nr_salario*0.08%><%if nr_salario <> "" Then response.write(formatnumber(fgts))%></td-->
								
					
						<!-- *** Base INSS *** -->		
							<%'base_irrf = total_vencimento - inss
							base_irrf = abs(nr_salario) + abs(nr_aj_custo) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(hora_extra) + abs(dsr_he)
								base_irrf = base_irrf - nr_valor_falta
								base_irrf = base_irrf - nr_desc_atrasos
								
							base_irrf = base_irrf - inss
							%>
								<td align="right" class="no_print">&nbsp;<%if base_irrf <> "" Then response.write(formatnumber(base_irrf,2))%></td>
					<!-- *** Faixa IRRF *** -->
					
								<%xsql = "Select * From TBL_IRRF Where dt_atualizacao <= '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
								Set rs_irrf = dbconn.execute(xsql)
								if not rs_irrf.EOF then
									deduc_dependente = rs_irrf("deduc_dependente")
										deduc_dependente = deduc_dependente * qtd_dependentes_21
									
									faixa_1 = rs_irrf("pri_faixa")
									pri_aliq = rs_irrf("pri_aliq")
									pri_deducao = rs_irrf("pri_deducao")
									
									faixa_2 = rs_irrf("seg_faixa")
									seg_aliq = rs_irrf("seg_aliq")
									seg_deducao = rs_irrf("seg_deducao")
									
									faixa_3 = rs_irrf("ter_faixa")
									ter_aliq = rs_irrf("ter_aliq")
									ter_deducao = rs_irrf("ter_deducao")
									
									faixa_4 = rs_irrf("qua_faixa")
									qua_aliq = rs_irrf("qua_aliq")
									qua_deducao = rs_irrf("qua_deducao")
									
									faixa_5 = rs_irrf("qui_faixa")
									qui_aliq = rs_irrf("qui_aliq")
									qua_deducao = rs_irrf("qua_deducao")
								end if
								
									if qtd_dependentes_21 <> "0" then
										base_irrf_calculo = base_irrf - deduc_dependente
									else
										base_irrf_calculo = base_irrf
									end if
								%>
								<td align="right" class="no_print">
					<!-- *** Parcela a deduzir - IRRF *** -->
								<%if base_irrf_calculo = 0 then
									irrf = "0"
									parc_deduzir = 0
									response.write("")
								elseif base_irrf_calculo <= abs(faixa_1) then
									irrf = base_irrf_calculo * (pri_aliq/100)
									parc_deduzir = pri_deducao
									faixa_irrf = "Isento"
									'response.write("Isento")
								elseif base_irrf_calculo <= abs(faixa_2) then
									irrf = base_irrf_calculo * (seg_aliq/100)
									parc_deduzir = seg_deducao
									faixa_irrf = seg_aliq&"%"
									'response.write(seg_aliq&"%")
								elseif base_irrf_calculo <= abs(faixa_3) then
									irrf = base_irrf_calculo * (ter_aliq/100)
									parc_deduzir = ter_deducao
									faixa_irrf = ter_aliq&"%"
									'response.write(ter_aliq&"%")
								elseif base_irrf_calculo <= abs(faixa_4) then
									irrf = base_irrf_calculo * (qua_aliq/100)
									parc_deduzir = qua_deducao
									faixa_irrf = qua_aliq&"%"
									'response.write(qua_aliq&"%")
								elseif base_irrf_calculo > abs(faixa_5) then
									irrf = base_irrf_calculo * (qui_aliq/100)
									parc_deduzir = qui_deducao
									faixa_irrf = qui_aliq&"%"
									'response.write(qui_aliq&"%")
								end if
									
									if irrf <> "" Then 
										irrf_final = irrf - parc_deduzir
										'response.write(base_irrf)
									end if
									
										if irrf_final < 10 then
											irrf_final = 0
											faixa_irrf = "Isento"
											parc_deduzir = "0"
										end if
									
								'response.write(faixa_irrf&","&irrf)
								response.write(faixa_irrf)%></td>
								<td align="right" class="no_print">&nbsp;<%if parc_deduzir <> "" Then response.write(formatnumber(parc_deduzir,2))%></td>
					<!-- *** IRRF *** -->
								<td align="right">&nbsp;<%if irrf_final <> "0" then response.write(formatnumber(irrf_final,2))%></td>
							<%'total_liquido = total_vencimento - total_desconto
							'total_liquido = total_liquido - irrf_final%>
					<!-- *** Descontos *** -->
							<%'total_desconto = abs(nr_desconto_ferias) + abs(cesta_b) + abs(valor_refeicao) + abs(v_trasnsp) + abs(inss) + abs(nr_conv_medico) + abs(contr_sind) + abs(nr_desc_diversos)
							total_desconto = abs(nr_desc_atrasos) + abs(nr_valor_falta) + abs(cesta_b) + abs(valor_refeicao) + abs(v_trasnsp) + abs(inss) + abs(nr_conv_medico) + abs(nr_contr_sind) + abs(nr_desc_diversos) + abs(irrf_final)%>
								<td align="right" class="no_print">&nbsp;<b><%if total_desconto <> "" Then response.write(formatnumber(total_desconto,2))%></b></td>
					<!-- *** Total Liquido *** -->	
							<%'total_desconto = total_desconto + irrf_final
							total_liquido = total_vencimento - total_desconto
								'if irrf_final <= "10" then
									'total_liquido = total_liquido + irrf
								'end if
							total_liquido = total_liquido' + auxcreche%>
								<td align="right">&nbsp;<%if total_liquido <> 0 then response.write(formatnumber(total_liquido,2))%></td>
							</tr>
								<%tt_salario = abs(tt_salario) + nr_salario
								tt_aj_custo = abs(tt_aj_custo) + nr_aj_custo
								tt_insalubridade = abs(tt_insalubridade) + insalubridade
								tt_qtd_dia_noturno = abs(tt_qtd_dia_noturno) + qtd_dia_noturno
								tt_ad_noturno = abs(tt_ad_noturno) + ad_noturno
								tt_dsr_ad_not = abs(tt_dsr_ad_not) + dsr_ad_not
								tt_qtd_dependentes_21 = abs(tt_qtd_dependentes_21) + qtd_dependentes_21
								tt_qtd_dependentes = abs(tt_qtd_dependentes) + qtd_dependentes
								tt_auxcreche = abs(tt_auxcreche) + auxcreche
								tt_qtd_hextra = tt_qtd_hextra + qtd_hextra
								tt_hora_extra = abs(tt_hora_extra) + hora_extra
								tt_dsr_he = abs(tt_dsr_he) + dsr_he
								tt_total_vencimento = abs(tt_total_vencimento) + total_vencimento
								tt_nr_faltas = abs(tt_nr_faltas) + nr_faltas
								tt_nr_dsr_faltas = abs(tt_nr_dsr_faltas) + nr_dsr_faltas
								
								strcod = "0"
								cd_matricula = "0"
								nm_nome = "0"
								tempo_trabalhado = ""
								cd_unidade = "0"
								nm_unidade = "0"
								nm_sigla = "0"
								cd_hospital = "0"
								cd_salario = "0"
								nr_salario = "0"
								nr_salario_hora = "0"
								nr_salario_dia = "0"
								nr_salario_desc = "0"
								nr_desc_salario = "0"
								insalubridade = "0"
								nr_insalubridade_forca = "0"
								total_faltas_mes = "0"
								expediente = "0"
								tipo_expediente = "0"
								carga_mensal = "0"
								cd_aj_custo = "0"
								nr_aj_custo = "0"
								cd_ad_noturno = "0"
								qtd_dia_noturno = "0"
								ad_noturno = "0"
								dsr_ad_not = "0"
								cd_hextra = "0"
								qtd_hextra = "0"
								hora_extra = "0"
								dsr_he = "0"
								cd_vtransp_cancela = "0"
								nr_vtransp_cancela = "0"
								cd_conv_medico = "0"
								nr_conv_medico = "0"
								cd_contr_sind = "0"
								nr_contr_sind = "0"
								qtd_dependentes = "0"
								qtd_dep_auxcreche = "0"
								qtd_mes_dep = "0"
								qtd_dep_idade = "0"
								idade = ""
								ok_auxcreche = "0"
								auxcreche = 0
								cesta_b = "0"
								valor_refeicao = "0"
								v_trasnsp = "0"
								base_inss = "0"
								nr_conv_medico = "0"
								nr_contr_sind = "0"
								fgts = "0"
								total_vencimento = "0"
								total_desconto = "0"
								base_irrf = "0"
								periodo_ferias = ""
								nr_desconto_ferias = "0"
								nr_desc_atrasos = "0"
								
								num_cor =  num_cor + 1
								if num_cor > 2 then
									num_cor = 1
								end if
								
								num_linha = num_linha + 1
									if num_linha = 51 then
										num_linha = 1
										class_print = "ok_print"%>
										<tr style="page-break-after:always;"><td></td></tr>
									<%end if
									
								rs.movenext
								wend%>
							
							<tr>
								<td><img src="../../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="225" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="55" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="48" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="55" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/px.gif" alt="" width="45" height="1" border="0"></td>								
							</tr>
							<tr>
								<td colspan="5">&nbsp;</td>
								<td align="right"><b><%if tt_salario > 0 then response.write(formatnumber(tt_salario,2))%></b></td>
								<td align="right"><b><%if tt_aj_custo > 0 then response.write(formatnumber(tt_aj_custo,2))%></b></td>
								<td align="right"><b><%if tt_insalubridade > 0 then response.write(formatnumber(tt_insalubridade,2))%></b></td><br>
								<td align="right"><b><%if tt_qtd_dia_noturno > 0 then response.write(tt_qtd_dia_noturno)%></b></td>
								<td align="right"><b><%if tt_ad_noturno > 0 then response.write(formatnumber(tt_ad_noturno,2))%></b></td>
								<td align="right"><b><%if tt_dsr_ad_not > 0 then response.write(formatnumber(tt_dsr_ad_not,2))%></b></td>
								<td align="right"><b><%if tt_qtd_dependentes_21 > 0 then response.write(tt_qtd_dependentes_21)%></b></td>
								<td align="right"><b><%if tt_auxcreche > 0 then response.write(tt_auxcreche)%></b></td>
								<td align="right"><b><%if tt_qtd_hextra > 0 then response.write(tt_qtd_hextra)%></b><%=tt_qtd_hextra%>X</td>
								<td align="right"><b><%if tt_hora_extra > 0 then response.write(FormatNumber(tt_hora_extra,2))%></b></td>
								<td align="right"><b><%if tt_dsr_he > 0 then response.write(FormatNumber(tt_dsr_he,2))%></b></td>
								<td align="right"><b><%if tt_total_vencimento > 0 then response.write(FormatNumber(tt_total_vencimento,2))%></b></td>
								<td align="right"><b><%if tt_nr_faltas > 0 then response.write(tt_nr_faltas)%></b></td>
								<td align="right"><b><%if tt_nr_dsr_faltas > 0 then response.write(tt_nr_dsr_faltas)%></b></td>
								
								
								
							</tr>
						</table>
			