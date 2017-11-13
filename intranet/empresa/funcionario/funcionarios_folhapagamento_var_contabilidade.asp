<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--script language="JavaScript" type="text/javascript" src="js/richtext.js"></script-->
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!--include file="../../css/geral.htm"-->
<%
cd_user = session("cd_codigo")

strcod = request("cod")
strcod = ""
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strcat = request("cat")
strbusca = request("busca")
stracao = request("acao")

strcd_unidade = request("cd_unidade")
	if strcd_unidade = "" Then	strcd_unidade = 0
	
'cd_ordem = request("cd_ordem")
	'if cd_ordem = "1" then
		'nm_ordem = "nm_unidade,nm_nome"
	'Elseif cd_ordem = "2" then
		nm_ordem = "nm_nome,nm_unidade"
	'Elseif cd_ordem = "" AND strcat = "enfermagem"Then 
	'	nm_ordem = "nm_unidade,nm_nome"
	'	cd_ordem = "1"
	'elseif cd_ordem = "" AND strcat = "empresa" Then 
	'	nm_ordem = "nm_nome,nm_unidade"
	'	cd_ordem = "2"
	'end if

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
		
		mes_ant_comp = mes_competencia - 1
			if mes_ant_comp < 1 then
				mes_ant_comp = 12
				ano_ant_comp = ano_competencia - 1
			else
				ano_ant_comp = ano_competencia
			end if
			final_ant_comp = ultimodiames(mes_ant_comp,ano_ant_comp)
	
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
	<%
							
							
							
							
							mes_competencia = dt_mes - 1
		if mes_competencia < 1 then
			mes_competencia = 12
			ano_competencia = dt_ano - 1
		else
			ano_competencia = dt_ano
		end if
		final_competencia = ultimodiames(mes_competencia,ano_competencia)
		
		mes_ant = mes_competencia - 1
		ano_ant = ano_competencia
			if mes_ant < 1 then
				mes_ant = 12
				ano_ant = ano_ant - 1
			end if
								
				
		
		%>

					
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" style="border-collapse:collapse; font-family:arial;">
					
			<!-- ---------------------------------------------------------------- -->				
							<!--tr style="color:whites;" class="no_print">
								<td align="left" colspan="32">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês &nbsp; &nbsp; &nbsp; &nbsp; (<%=(qtd_dias_comp-qtd_domingos_comp)-qtd_feriados%>)
									{<%=var_dia_semana%>}</td>
							</tr-->
							
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
							'*** Funcionarios +1 dia  ***
							strsql = "SELECT * FROM TBL_funcionario_plantoes_1d where dt_data = '"&dt_mes&"/1/"&dt_ano&"'"
							Set rs_1d = dbconn.execute(strsql)
								if not rs_1d.EOF then nm_1d = rs_1d("nm_funcionarios")
								'response.write(nm_1d)

					num_cor = 1
					num_funcionario = 1
					cab_contab = 1
					print = "nok_print"
					
					pagina = 1
					'*** configura o total de linhas por pág.***
						total_linhas_pag = 50
					
					'xsql = "up_funcionario_rh_lista_individual @cd_funcionario="&strcod&", @dt_atualizacao='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
					'xsql = "up_funcionario_folhapagamento_geral @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"',@outras_cond=' ' ,@nm_ordem='"&nm_ordem&"'"
					'xsql = "up_funcionario_contrato_lista3 @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', '"
					xsql = "up_funcionario_contrato_lista3 @dt_data_i='"&mes_competencia&"/21/"&ano_competencia&"', @dt_data_f='"&mes_competencia&"/21/"&ano_competencia&"', @dt_atualizacao = '"&dt_mes&"/21/"&dt_ano&"', @ordem_funcionarios='"&nm_ordem&"', @outras_variaveis='"&outras_cond&"'"
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
								
								dt_admissao = rs("dt_admissao_geral")
								dt_demissao = rs("dt_demissao_geral")
								
								admissao = rs("dt_admissao_geral")
								'demissao = rs("demissao")
									if dt_demissao <> "null" then
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
								
								'*** carrega as informações dos plantões ***
								strsql = "SELECT * FROM TBL_funcionario_plantoes WHERE dt_atualizacao='"&dt_mes&"/1/"&dt_ano&"'"
								Set rs_plantoes = dbconn.execute(strsql)
									if not rs_plantoes.EOF then
										plantoes_6h = rs_plantoes("nr_plantoes_6h")
										plantoes_8h = rs_plantoes("nr_plantoes_8h")
										plantoes_12h = rs_plantoes("nr_plantoes_12h")
									end if
								
								expediente = rs("expediente")
									expediente = expediente / 60
										if expediente >= 9 AND expediente <= 11 then
											tipo_expediente = "8h"
											carga_mensal = 220
											qtd_plantoes = plantoes_8h
										elseif expediente = 6 then
											tipo_expediente = "6h"
											carga_mensal = 180
											qtd_plantoes = plantoes_6h
										elseif expediente = 12 then
											tipo_expediente = "12x36h"
											carga_mensal = 180
											qtd_plantoes = plantoes_12h
												
												'*** acerto +1 dia ***
												func_1d = instr(1,nm_1d,strcod,1)
												if func_1d > 0 then qtd_plantoes = qtd_plantoes + 1
										end if
								
								hr_entrada = rs("hr_entrada")
								hr_saida = rs("hr_saida")
									periodo = hour(hr_entrada) - hour(hr_saida)
										if periodo < 0 then
											nm_periodo = "diurno"
											cd_periodo = 1
											nr_ad_noturno = 0
										elseif periodo > 0 then
											nm_periodo = "noturno"
											cd_periodo = 2
											nr_ad_noturno = qtd_plantoes
										end if
											total_geral_nr_ad_noturno = total_geral_nr_ad_noturno + nr_ad_noturno
								
								
								if (cd_matricula = 137) and (dt_mes = 3)  and (dt_ano = 2017)  or (cd_matricula = 9999) and (dt_mes = 3) and (dt_ano = 2017) then	
								nm_nome = "IVAN"
								nr_ad_noturno = 0
								end if
								
								'ferias
								if (cd_matricula = 287) and (dt_mes = 4) and (dt_ano = 2017) or (cd_matricula = 253) and (dt_mes = 4) and (dt_ano = 2017)  then
											cd_periodo = 1
											nr_ad_noturno = 0
											nm_periodo = "diurno"
											end if
								
								
							

								
								'*** Carrega as variáveis do RCM ***
								'xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
								
								xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_ant&"/21/"&ano_ant&"', @dt_f='"&mes_competencia&"/20/"&ano_competencia&"'"
								
								
								'xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_competencia_anterior&"/21/"&ano_competencia_anterior&"', @dt_f='"&mes_competencia&"/20/"&ano_competencia&"'"
							
								Set rs_rcm = dbconn.execute(xsql)
									if not rs_rcm.EOF then
										total_ad_noturno_he = rs_rcm("total_ad_noturno")
										total_func_ad_noturno = nr_ad_noturno + total_ad_noturno
											total_geral_ad_noturno_he = total_geral_ad_noturno_he + total_ad_noturno_he
										total_he = rs_rcm("total_he")/60
											'********************************************
											'*** Arredondamento Hora Extra fracionada ***
													if instr(1,total_he,",",1) > 0 then 
														virgula_he = instr(1,total_he,",",1)
														fracao_he = mid(total_he,virgula_he+1,len(total_he))
															if fracao_he > 2 then
																total_he = int(total_he) + 1
																sinal_arredonda = "*"
															else
																total_he = total_he
																sinal_arredonda = ""
															end if
													end if
													if total_he > 0 then total_geral_he = total_geral_he + total_he
										total_falta_justif = rs_rcm("total_falta_justif")
											total_geral_falta_justif = total_geral_falta_justif + total_falta_justif
										total_falta_injust = rs_rcm("total_falta_injust")
											total_geral_falta_injust = total_geral_falta_injust + total_falta_injust
										total_faltas = rs_rcm("total_faltas")
											total_geral_faltas = total_geral_faltas + total_faltas
										total_atraso = rs_rcm("total_atraso")
											total_geral_atraso = total_geral_atraso + total_atraso
										total_vr_extra = rs_rcm("total_vr_extra")
											total_geral_vr_extra = total_geral_vr_extra + total_vr_extra
										total_vt_extra = rs_rcm("total_vt_extra")
											total_geral_vt_extra = total_geral_vt_extra + total_vt_extra
									end if
										if   nr_ad_noturno > 0 then
										total_func_ad_noturno = (nr_ad_noturno + total_ad_noturno_he)'- total_faltas'
										else
										total_func_ad_noturno = nr_ad_noturno + total_ad_noturno_he
										end if
										
											'total_geral_func_ad_noturno = total_geral_func_ad_noturno + total_func_ad_noturno
												'total_geral_ad_noturno = total_geral_ad_noturno + (total_ad_noturno_he + nr_ad_noturno)
												total_geral_ad_noturno = total_geral_ad_noturno +  total_func_ad_noturno
										'response.write total_geral_ad_noturno
										'*** Separa os valores de H.E. ***
										if total_he > 36 then
											total_he_plus = total_he - 36
											total_he_limit = 36
										else
											total_he_limit = total_he
											total_he_plus = 0
										end if
											if total_he <> "" Then 
												total_geral_he_plus = total_geral_he_plus + total_he_plus
												total_geral_he_limit = total_geral_he_limit + total_he_limit
											end if
										
										'*** VT Extra: Ida e Volta
										if total_vt_extra <> "" then
											total_vt_extra = total_vt_extra *2 
										end if
									
								if qtd_plantoes = 13 or qtd_plantoes = 22 then
									nr_vr = qtd_plantoes - qtd_total_faltas
									'nr_vr = qtd_total_faltas
								end if
								
								if qtd_plantoes <> "0" then
									nr_vt = (qtd_plantoes * 2)' - qtd_total_faltas
									qtd_total_faltas_vt = qtd_total_faltas * 2
									nr_vt = nr_vt - qtd_total_faltas_vt
								end if
								
									if num_cor = 1 then
										cor_linha = "#ccffff"
									else
										cor_linha = "#ffffff"
									end if%>
							
							
					<%if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
							if qtd_hextra <> "" Then tt_hextra = abs(tt_hextra) + abs(qtd_hextra)
							if qtd_hextra_plus <> "" Then tt_hextra_plus = abs(tt_hextra_plus) + abs(qtd_hextra_plus)
							if nr_faltas_justif <> "" Then tt_faltas_justif = abs(tt_faltas_justif) + abs(nr_faltas_justif)
							if nr_faltas_injust <> "" Then tt_faltas_injust = abs(tt_faltas_injust) + abs(nr_faltas_injust)
							if nr_atrasos <> "" Then tt_atrasos = abs(tt_atrasos) + abs(nr_atrasos)
							
							
							if nr_vr <> "" Then tt_vr = abs(tt_vr) + abs(nr_vr)
							if nr_vt <> "" Then tt_vt = abs(tt_vt) + abs(nr_vt)
							
							if cab_contab = 1 then%>
							<tr class="<%=print%>">
								<td align="center" colspan="2">&nbsp;</td>
								<td align="center" colspan="4">Mês de competência: <b><%=ucase(mes_selecionado(mes_competencia))%>/<%=ano_competencia%></b></td>
								<td align="center"><img src="../../imagens/ic_print.gif" alt="imprimir" width="24" height="26" border="0" onclick="window.print();" class="no_print">
								&nbsp;<!--img src="../../imagens/px.gif" alt="imprimir" width="24" height="26" border="0" onclick="visualizarImpressao();" class="ok_print"--></td></tr>
							<tr bgcolor="#000000" class="<%=print%>">
								<td colspan="7" align="center" class="cabecalho" style="background-color:black; color:white;"><b>Horas Extras / Ad. Noturno / Faltas e atrasos</b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr style="background-color:gray; color:white; font-size:12px;" class="<%=print%>">
								<td>&nbsp;<img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
								<td align="center"><b>Mat.</b><br><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
								<td align="center"><b>Funcion&aacute;rio</b><br><img src="../../imagens/px.gif" alt="" width="275" height="1" border="0"></td>
								<!--td align="center"><b>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<td align="center"><b>Qtd H.E.</b><br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<td align="center"><b>Faltas injust.</b><br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<td align="center"><b>Atrasos</b><br>(/horas)<br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<td>***</td-->
								<td align="center"><b>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<td align="center"><b>Qtd H.E.</b><br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<td align="center"><b>Faltas injust.</b><br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<!--td align="center"><b>Faltas just.</b><br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td-->
								<td align="center"><b>Atrasos</b><br>(horas)<br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>								
								
							</tr>
						<%cab_contab = cab_contab + 1
						print = "ok_print"
						end if%>
					
						<%'if int(total_func_ad_noturno) > 0 OR int(total_he) <> 0 OR int(total_falta_injust) <> 0 OR int(total_atraso) <> 0 then
						if int(total_func_ad_noturno) > 0 OR int(total_he) > 0 OR int(total_falta_injust) > 0 OR int(total_atraso) > 0 then%>
							<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');" bgcolor="<%=cor_linha%>" style=" font-size:12px;">								
								<td align="right" ><b><%=zero(num_funcionario)%></b>&nbsp;<%'=" - "&periodo_ferias%></td>
								<td align="right"><%=cd_matricula%>&nbsp;</td>
								<td style="color:<%=cor_registro%>;" align="left"><%=nm_nome%> <%'=tempo_trabalhado%></td>
								<!--td style="color:<%=cor_registro%>;" align="center"><%=nm_sigla%></td-->
								
													
						<!-- *** Qtd. Adicional Noturno *** -->
								<td align="center">&nbsp;<%if total_func_ad_noturno > 0 then%><b><%=total_func_ad_noturno%></b><%else%>0<%end if%></td>
						<!-- *** Qtd Hora Extra *** -->
								<td align="center"><%if total_he > 0 then%><b><%=(total_he_limit)%></b><%else%>0<%end if%></td>
						<!-- *** Faltas Não justificadas *** -->
								<td align="center"><%if total_falta_injust > 0 then%><b><%=(total_falta_injust)%></b><%else%>0<%end if%></td>
						<!-- *** Faltas Não justificadas *** -->
								<!--td align="center"><%if total_falta_justif > 0 then%><b><%=(total_falta_justif)%></b><%else%>0<%end if%></td-->
						<!-- *** Atrasos *** -->
								<td align="center"><%if total_atraso > 0 then%><b><%=(total_atraso)%></b><%else%>0<%end if%></td>
							
							</tr>
							<%num_funcionario = num_funcionario + 1
							linha_conta = linha_conta + 1
							num_cor =  num_cor + 1
								if num_cor > 2 then
									num_cor = 1
								end if
								
							if linha_conta = total_linhas_pag then%>
								<tr class="ok_print"><td colspan="7">&nbsp;</td></tr>
								<tr align="right" style="page-break-after:always;" class="ok_print"><td colspan="7">Página: <%=pagina%> &nbsp; </td></tr>
							<%pagina = pagina + 1
							end if
						end if
					end if		
								strcod = "0"
								cd_matricula = "0"
								nm_nome = "0"
								tempo_trabalhado = ""
								cd_unidade = "0"
								nm_unidade = "0"
								nm_sigla = "0"
								cd_hospital = "0"
								expediente = "0"
								tipo_expediente = "0"
								carga_mensal = "0"
								
								total_ad_noturno_he = "0"
								total_func_ad_noturno = ""
								total_he = "0"
								total_falta_injust = "0"
								total_atraso = "0"
								
								cd_ad_noturno = "0"
								qtd_dia_noturno = ""
								ad_noturno = "0"
								dsr_ad_not = "0"
								cd_hextra = "0"
								qtd_hextra = "0"
								qtd_hextra_plus = "0"
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
								nr_faltas_justif = "0"
								nr_faltas_injust = "0"
								nr_vr = "0"
								nr_vt = "0"
								qtd_total_faltas = "0"
								
								if linha_conta = total_linhas_pag then
									cab_contab = 1
									linha_conta = 1
								end if
								
								rs.movenext
								wend
								
								if total_geral_he_limit = 000000000000000000000000000000000000000000000000000000000000	then
								total_geral_he_limit = 0
								end if
								
								%>
							
							<tr><td colspan="7">&nbsp;</td></tr>
							
							<tr>
								<td align="center" colspan="3">&nbsp;<b>TOTAL</b></td>
								<td align="center">&nbsp;<b><%=total_geral_ad_noturno%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_he_limit%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_falta_injust%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_atraso%></b></td>
									
								
							</tr>
							<tr class="ok_print"><td colspan="7">&nbsp;</td></tr>
							<tr class="ok_print"><td align="right" colspan="7">impr.:<%=now%><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"> Página: <%=pagina%> &nbsp; </td></tr>
						</table>
			