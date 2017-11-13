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
			
data_atual = dia_final&"/"&dt_mes&"/"&dt_ano%>
<!--#include file="../../includes/feriados.asp"-->
<%qtd_dias_dsr = qtd_domingos_comp + qtd_feriados%>


					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print">
						
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="empresa.asp"  method="POST">
							<input type="hidden" name="tipo" value="variaveis">
							<input type="hidden" name="cod" value="<%=strcod%>">
							<tr>
								<td colspan="4" align="center" class="cabecalho" style="background-color:black; color:white;"><%'=data_atual%> <b>VARIÁVEIS - FOLHA DE PAGAMENTO <%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td align="center" style="background-color:silver;"><b>Mês para Pagamento:</b></td>
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
								<td align="center" style="background-color:silver; color:red;"><b><%=mesdoano(mes_competencia)%>/<%=ano_competencia%></b></td>
							</tr>
							<tr>
								<td><img src="../../imagens/blackdot.gif" alt="" width="150" height="0" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="150" height="0" border="0"></td>
							</tr>
						</form>
					</table>
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" style="border-collapse:collapse;">	
			<!-- ---------------------------------------------------------------- -->				
							<!--tr style="color:whites;" class="no_print">
								<td align="left" colspan="32">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês &nbsp; &nbsp; &nbsp; &nbsp; (<%=(qtd_dias_comp-qtd_domingos_comp)-qtd_feriados%>)
									{<%=var_dia_semana%>}</td>
							</tr-->
							<tr>
								<td colspan="8" align="center" class="cabecalho" style="background-color:black; color:white;"><b>Horas Extras / Ad. Noturno / Faltas e atrasos - <%=mesdoano(mes_pagamento)%>/<%=ano_pagamento%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr style="background-color:gray; color:white;">
								<td>&nbsp;</td>
								<td align="center"><b>Mat.</b></td>
								<td align="center"><b>Funcion&aacute;rio</b></td>
								<td align="center"><b>Ad. Not</b></td>
								<td align="center"><b>Qtd H.E.</b></td>
								<td align="center"><b>Faltas</b></td>
								<td align="center"><b>DSR.</b></td>
								<td align="center"><b>Atrasos</b></td>
							</tr>
							<%'**************************
							'*** VALOR SALÁRIO MÍNIMO ***
							xsql = "SELECT * FROM TBL_funcionario_valores_padroes WHERE (cd_tipo = 12) AND dt_atualizacao <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
							Set rs = dbconn.execute(xsql)
							nr_salminimo = rs("nr_valor")
							'****************************
							'*** VALOR Aux. Creche	  ***
							xsql = "SELECT * FROM TBL_funcionario_valores_padroes WHERE (cd_tipo = 6) AND dt_atualizacao <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
							Set rs = dbconn.execute(xsql)
							nr_auxcreche = rs("nr_valor")
							'******************************

					num_cor = 1
					'xsql = "up_funcionario_rh_lista_individual @cd_funcionario="&strcod&", @dt_atualizacao='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
					xsql = "up_funcionario_folhapagamento_geral @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
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
									if admissao = dt_ano&zero(dt_mes)  AND demissao = "" OR admissao = dt_ano&zero(dt_mes) AND demissao > dt_ano&zero(dt_mes) then
										tempo_trabalhado = DateDiff("d",day(dt_admissao)&"/"&month(dt_admissao)&"/"&year(dt_admissao),dia_final&"/"&dt_mes&"/"&dt_ano)+1 '(1 = conta o dia da contratação)
									
									'*** Caso a admissão e a demissão forem no mesmo mês de competencia da folha ***
									elseif admissao = dt_ano&zero(dt_mes)  AND demissao = dt_ano&zero(dt_mes) then
										tempo_trabalhado = DateDiff("d",day(dt_admissao)&"/"&month(dt_admissao)&"/"&year(dt_admissao),dia_final&"/"&dt_mes&"/"&dt_ano)+1 '(1 = conta o dia da contratação)
									
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
									
								expediente = rs("expediente")
									expediente = expediente / 60
										if expediente >= 9 AND expediente <= 11 then
											tipo_expediente = "8h"
											carga_mensal = 220
										elseif expediente = 6 then
											tipo_expediente = "6h"
											carga_mensal = 180
										elseif expediente = 12 then
											tipo_expediente = "12x36h"
											carga_mensal = 180
										end if
									if nr_salario <> "0" AND carga_mensal <> "0" Then 
										nr_salario_hora = nr_salario / carga_mensal
									else
										nr_salario_hora = "0"
									end if
								
								'cd_aj_custo = rs("cd_aj_custo")
								'nr_aj_custo = rs("nr_aj_custo")
								'	if not IsNumeric(nr_aj_custo) Then nr_aj_custo = "0"
								
								'cd_insalubridade_forca = rs("cd_insalubridade_forca")
								'nr_insalubridade_forca = rs("nr_insalubridade_forca")
								'	if  not IsNumeric(nr_insalubridade_forca) Then nr_insalubridade_forca = "0"
								
								cd_ad_noturno = rs("cd_ad_noturno")
								qtd_dia_noturno = rs("qtd_dia_noturno")
									if  not IsNumeric(qtd_dia_noturno) Then qtd_dia_noturno = ""
								
								cd_hextra = rs("cd_hextra")
								qtd_hextra = rs("qtd_hextra")
									if qtd_hextra > 36 then
										qtd_hextra = "36"
										he_marca = "1"
									end if
								
								'cd_vrefeic_cancela = rs("cd_vrefeic_cancela")
								'nr_vrefeic_cancela = rs("nr_vrefeic_cancela")
								'	if  not IsNumeric(nr_vrefeic_cancela) Then nr_vrefeic_cancela = "0"
								
								'cd_vtransp_cancela = rs("cd_vtransp_cancela")
								'nr_vtransp_cancela = rs("nr_vtransp_cancela")
								'	if  not IsNumeric(nr_vtransp_cancela) Then nr_vtransp_cancela = "0"
								
								'cd_conv_medico = rs("cd_conv_medico")
								'nr_conv_medico = rs("nr_conv_medico")
								'	if not IsNumeric(nr_conv_medico) then nr_conv_medico = "0"
								
								'cd_contr_sind = rs("cd_contr_sind")
								'nr_contr_sind = rs("nr_contr_sind")
								'	if  not IsNumeric(nr_contr_sind) Then nr_contr_sind = "0"
								
								'cd_desc_diversos = rs("cd_desc_diversos")
								'nr_desc_diversos = rs("nr_desc_diversos")
								'	if  not IsNumeric(nr_desc_diversos) Then nr_desc_diversos = "0"
								
								cd_faltas = rs("cd_faltas")
								nr_faltas = rs("nr_faltas")
									if  not IsNumeric(nr_faltas) Then nr_faltas = "0"
								
								cd_dsr_faltas = rs("cd_dsr_faltas")
								nr_dsr_faltas = rs("nr_dsr_faltas")
									if  not IsNumeric(nr_dsr_faltas) Then nr_dsr_faltas = "0"
								
								nr_atrasos = rs("nr_atrasos")
									if  not IsNumeric(nr_atrasos) Then nr_atrasos = "0"
								
								'*** Salário desconto faltas ****
								'	if IsNumeric(nr_faltas) then
									'if abs(nr_faltas) >= "1" then
								'		total_faltas_mes = abs(nr_faltas) + abs(nr_dsr_faltas)
									'	nr_desc_salario = nr_salario_dia * nr_faltas
									'	nr_salario_desc = nr_salario - nr_desc_salario
									'	
									'	nr_desc_salario_dsr = nr_salario_dia * total_faltas_mes
									'	nr_salario_desc_dsr = nr_salario - nr_desc_salario_dsr
								'		nr_valor_falta = nr_salario_dia * total_faltas_mes
								'		nr_salario_liq = nr_salario - nr_valor_falta
								'	else
								'		'nr_salario_desc = nr_salario
								'		nr_salario_liq = nr_salario
								'	end if
								
									
									'if nr_faltas <> "0" AND nr_dsr_faltas <> "0" then 
									'	total_faltas_mes = abs(nr_faltas) + abs(nr_dsr_faltas)
									'	nr_desc_salario = nr_salario_dia * total_faltas_mes
									'	nr_salario_desc = nr_salario - nr_desc_salario
									'else
									'	nr_salario_desc = nr_salario
									'end if
									
								
								
								num_funcionario = num_funcionario + 1
									if num_cor = 1 then
										cor_linha = "#ccffff"
									else
										cor_linha = "#ffffff"
									end if
									
							%>
							
							
							<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');" bgcolor="<%=cor_linha%>" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_cad.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>&var=1','','width=470,height=300,scrollbars=1')">								
								<td align="right"><b><%=zero(num_funcionario)%></b>&nbsp;<%'=" - "&periodo_ferias%></td>
								<td align="right" onClick="window.open('empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&cod=<%=strcod%>','','width=630,height=600,scrollbars=1')"><%=cd_matricula%>&nbsp;</td>
								<td style="color:<%=cor_registro%>;" align="left"><%=nm_nome%> <%'=tempo_trabalhado%></td>
								
								
					<!-- *** Qtd. Adicional Noturno *** -->
								<td align="right">&nbsp;<%if qtd_dia_noturno <> "" then%><%=zero(qtd_dia_noturno)%>d<%end if%><!-- <%'if qtd_dia_noturno > "" then%> dias<%'elseif qtd_dia_noturno <> "" then%>dia<%'end if%>--></td>
					<!-- *** Qtd Hora Extra *** -->
								<td align="right"><%if qtd_hextra >= 1 then%><%=zero(qtd_hextra)%>h<%end if%>&nbsp;</td>
					<!-- *** Faltas Não justificadas *** -->
								<td align="right"><%if nr_faltas >= 1 then%><%=zero(nr_faltas)%><%end if%>&nbsp;</td>
					<!-- *** DSR Faltas Não justificadas *** -->
								<td align="right"><%if nr_dsr_faltas >= 1 then%><%=zero(nr_dsr_faltas)%><%end if%>&nbsp;</td>
					<!-- *** Atrasos *** -->
								<td align="right"><%if nr_atrasos >= 1 then%><%=zero(nr_atrasos)&"h"%><%end if%>&nbsp;</td>
					
					<!-- *** Total Liquido *** -->	
							<%'total_desconto = total_desconto + irrf_final
							'total_liquido = total_vencimento - total_desconto
							'total_liquido = total_liquido + auxcreche%>
								<!--td align="right">&nbsp;<%'if total_liquido <> 0 then response.write(formatnumber(total_liquido,2))%></td-->
							</tr>
								<%tt_salario = abs(tt_salario) + nr_salario
								tt_aj_custo = abs(tt_aj_custo) + nr_aj_custo
								tt_insalubridade = abs(tt_insalubridade) + insalubridade
								if qtd_dia_noturno <> "" Then tt_qtd_dia_noturno = abs(tt_qtd_dia_noturno) + qtd_dia_noturno
								tt_ad_noturno = abs(tt_ad_noturno) + ad_noturno
								tt_dsr_ad_not = abs(tt_dsr_ad_not) + dsr_ad_not
								tt_qtd_dependentes_21 = abs(tt_qtd_dependentes_21) + qtd_dependentes_21
								tt_qtd_dependentes = abs(tt_qtd_dependentes) + qtd_dependentes
								
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
								qtd_dia_noturno = ""
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
								
								num_cor =  num_cor + 1
								if num_cor > 2 then
									num_cor = 1
								end if
								
								rs.movenext
								wend%>
							
							<!--tr>
								<td colspan="5">&nbsp;</td>
								<td align="right"><b><%if tt_salario > 0 then response.write(formatnumber(tt_salario,2))%></b></td>
								<td align="right"><b><%if tt_aj_custo > 0 then response.write(formatnumber(tt_aj_custo,2))%></b></td>
								<td align="right"><b><%if tt_insalubridade > 0 then response.write(formatnumber(tt_insalubridade,2))%></b></td><br>
								<td align="right"><b><%if tt_qtd_dia_noturno > 0 then response.write(tt_qtd_dia_noturno)%></b></td>
								<td align="right"><b><%if tt_ad_noturno > 0 then response.write(formatnumber(tt_ad_noturno,2))%></b></td>
								<td align="right"><b><%if tt_dsr_ad_not > 0 then response.write(formatnumber(tt_dsr_ad_not,2))%></b></td>
								<td align="right"><b><%if tt_qtd_dependentes_21 > 0 then response.write(tt_qtd_dependentes_21)%></b></td>
								<td align="right"><b><%if tt_qtd_dependentes > 0 then response.write(tt_qtd_dependentes)%></b></td>
							</tr-->
							
							<tr>
								<td><img src="../../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="25" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="225" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="45" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
								<!--td><img src="../../imagens/blackdot.gif" alt="" width="30" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="30" height="2" border="0"></td>
								<td class="no_print"><img src="../../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="45" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/blackdot.gif" alt="" width="30" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="35" height="2" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/blackdot.gif" alt="" width="55" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="45" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
								<td><img src="../../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/blackdot.gif" alt="" width="45" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/blackdot.gif" alt="" width="55" height="1" border="0"></td>
								<td class="no_print"><img src="../../imagens/blackdot.gif" alt="" width="45" height="1" border="0"></td-->								
							</tr>
						</table>
			