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


					
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" style="border-collapse:collapse; font-family:arial;">
					
			<!-- ---------------------------------------------------------------- -->				
							<!--tr style="color:whites;" class="no_print">
								<td align="left" colspan="32">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês &nbsp; &nbsp; &nbsp; &nbsp; (<%=(qtd_dias_comp-qtd_domingos_comp)-qtd_feriados%>)
									{<%=var_dia_semana%>}</td>
							</tr-->
							<tr><td align="center" colspan="6">Mês de pagamento: <b><%=ucase(mes_selecionado(dt_mes))%>/<%=dt_ano%></b></td>
								<td align="center"><img src="../../imagens/ic_print.gif" alt="imprimir" width="24" height="26" border="0" onclick="visualizarImpressao();" id="no_print"></td></tr>
							<tr bgcolor="#000000">
								<td colspan="100%" align="center" class="cabecalho" style="background-color:black; color:white;"><b>Horas Extras / Ad. Noturno / Faltas e atrasos</b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr style="background-color:gray; color:white; font-size:12px;">
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
								<td align="center"><b>Atrasos</b><br>(/horas)<br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>								
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
					num_funcionario = 1
					'xsql = "up_funcionario_rh_lista_individual @cd_funcionario="&strcod&", @dt_atualizacao='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
					xsql = "up_funcionario_folhapagamento_geral_old @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"',@outras_cond=' ' ,@nm_ordem='"&nm_ordem&"'"
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
								
								cd_ad_noturno = rs("cd_ad_noturno")
								qtd_dia_noturno = rs("qtd_dia_noturno")
									if  not IsNumeric(qtd_dia_noturno) Then qtd_dia_noturno = 0'""
								
								cd_hextra = rs("cd_hextra")
								qtd_hextra = rs("qtd_hextra")
									if qtd_hextra = "" Then qtd_hextra = 0
									if qtd_hextra > 36 then
										qtd_hextra_plus = qtd_hextra - 36
										'he_marca = "1"
									end if
										
										
									if qtd_hextra > 36 then
										qtd_hextra = "36"
										he_marca = "1"
									end if
								cd_qtd_plantoes = rs("cd_qtd_plantoes")
								qtd_plantoes = rs("qtd_plantoes")
									if qtd_plantoes = "" Then qtd_plantoes = 0
								
								cd_faltas_justif = rs("cd_faltas_justif")
								nr_faltas_justif = rs("nr_faltas_justif")
									if  not IsNumeric(nr_faltas_justif) Then nr_faltas_justif = "0"
								
								cd_faltas_injust = rs("cd_faltas_injust")
								nr_faltas_injust = rs("nr_faltas_injust")
									if  not IsNumeric(nr_faltas_injust) Then nr_faltas_injust = "0"
									
									qtd_total_faltas = abs(nr_faltas_justif) + abs(nr_faltas_injust)
										if qtd_total_faltas =  "" then qtd_total_faltas = 0
								
								cd_dsr_faltas = rs("cd_dsr_faltas")
								nr_dsr_faltas = rs("nr_dsr_faltas")
									if  not IsNumeric(nr_dsr_faltas) Then nr_dsr_faltas = "0"
								
								nr_atrasos = rs("nr_atrasos")
									if  not IsNumeric(nr_atrasos) Then nr_atrasos = "0"
									
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
							
							
'***************************** ANTERIOR **********************************
							'xsql_ant = "up_funcionario_folhapagamento_individual @cd_funcionario='"&strcod&"', @dt_comp_i='"&mes_ant_comp&"/1/"&ano_ant_comp&"', @dt_comp_f='"&mes_ant_comp&"/"&final_ant_comp&"/"&ano_ant_comp&"', @dt_i = '"&mes_competencia&"/1/"&ano_competencia&"', @dt_f = '"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
							'Set rs_ant = dbconn.execute(xsql_ant)
							''while not rs_ant.EOF
							'if not rs_ant.EOF then
							'	cd_ad_noturno_ant = rs_ant("cd_ad_noturno")
							'	qtd_dia_noturno_ant = rs_ant("qtd_dia_noturno")
							'		if  not IsNumeric(qtd_dia_noturno_ant) Then qtd_dia_noturno_ant = 0'""
							'	
							'	'cd_hextra_ant = rs_ant("cd_hextra")
							'	qtd_hextra_ant = rs_ant("qtd_hextra")
							'		if qtd_hextra_ant = "" Then qtd_hextra_ant = 0
							'		if qtd_hextra_ant > 36 then
							'			qtd_hextra_plus_ant = qtd_hextra_ant - 36
							'		end if
							'			
							'		if qtd_hextra_ant > 36 then
							'			qtd_hextra_ant = "36"
							'			he_marca_ant = "1"
							'		end if
								
							'	'cd_faltas_injust = rs_ant("cd_faltas_injust")
							'	nr_faltas_injust_ant = rs_ant("nr_faltas_injust")
							'		if  not IsNumeric(nr_faltas_injust_ant) Then nr_faltas_injust_ant = "0"
							'		
							'		qtd_total_faltas_ant = abs(nr_faltas_justif_ant) + abs(nr_faltas_injust_ant)
							'			if qtd_total_faltas_ant =  "" then qtd_total_faltas_ant = 0
							'	
							'	nr_atrasos_ant = rs_ant("nr_atrasos")
							'		if  not IsNumeric(nr_atrasos_ant) Then nr_atrasos_ant = "0"
							'	'*************************************************************************
							'	rs_ant.movenext
							''wend
							'end if
							
					'if int(qtd_dia_noturno_ant) <> int(qtd_dia_noturno) OR int(qtd_hextra_ant) <> int(qtd_hextra) OR int(nr_faltas_injust_ant) <> int(nr_faltas_injust) OR int(nr_atrasos_ant) <> int(nr_atrasos)then
					if int(qtd_dia_noturno) <> 0 OR int(qtd_hextra) <> 0 OR int(nr_faltas_injust) <> 0 OR int(nr_atrasos) <> 0 then%>
							<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');" bgcolor="<%=cor_linha%>" style=" font-size:12px;">								
								<td align="right" ><b><%=zero(num_funcionario)%></b>&nbsp;<%'=" - "&periodo_ferias%></td>
								<td align="right"><%=cd_matricula%>&nbsp;</td>
								<td style="color:<%=cor_registro%>;" align="left"><%=nm_nome%> <%'=tempo_trabalhado%></td>
								<!--td style="color:<%=cor_registro%>;" align="center"><%=nm_sigla%></td-->
								
								<!--td align="center">&nbsp;<%if qtd_dia_noturno_ant > 0 then%><%=int(qtd_dia_noturno_ant)%><%else%>0<%end if%></td>
								<td align="center"><%if qtd_hextra_ant > 0 then%><%=int(qtd_hextra_ant)%><%else%>0<%end if%></td>
								<td align="center"><%if nr_faltas_injust_ant > 0 then%><%=int(nr_faltas_injust_ant)%><%else%>0<%end if%></td>
								<td align="center"><%if nr_atrasos_ant > 0 then%><%=zero(nr_atrasos_ant)%><%else%>0<%end if%></td>
						
	<td><%if int(qtd_dia_noturno_ant) <> int(qtd_dia_noturno) then response.write("a")%>
		<%if int(qtd_hextra_ant) <> int(qtd_hextra) then response.write("b")%>
		<%if int(nr_faltas_injust_ant) <> int(nr_faltas_injust) then response.write("c")%>
		<%if int(nr_atrasos_ant) <> int(nr_atrasos) then response.write("d")%>
	</td-->		
											
					<!-- *** Qtd. Adicional Noturno *** -->
								<td align="center">&nbsp;<%if qtd_dia_noturno > 0 then%><%=int(qtd_dia_noturno)%><%else%>0<%end if%><!-- <%'if qtd_dia_noturno > "" then%> dias<%'elseif qtd_dia_noturno <> "" then%>dia<%'end if%>--></td>
					<!-- *** Qtd Hora Extra *** -->
								<td align="center"><%if qtd_hextra > 0 then%><%=int(qtd_hextra)%><%else%>0<%end if%></td>
					<!-- *** Faltas Não justificadas *** -->
								<td align="center"><%if nr_faltas_injust > 0 then%><%=int(nr_faltas_injust)%><%else%>0<%end if%></td>
					<!-- *** Atrasos *** -->
								<td align="center"><%if nr_atrasos > 0 then%><%=zero(nr_atrasos)%><%else%>0<%end if%></td>
							</tr>
						<%num_funcionario = num_funcionario + 1
							num_cor =  num_cor + 1
								if num_cor > 2 then
									num_cor = 1
								end if
					%>
								<%tt_salario = abs(tt_salario) + nr_salario
								tt_aj_custo = abs(tt_aj_custo) + nr_aj_custo
								tt_insalubridade = abs(tt_insalubridade) + insalubridade
								if qtd_dia_noturno <> "" Then tt_qtd_dia_noturno = abs(tt_qtd_dia_noturno) + qtd_dia_noturno
								tt_ad_noturno = abs(tt_ad_noturno) + ad_noturno
								tt_dsr_ad_not = abs(tt_dsr_ad_not) + dsr_ad_not
								tt_qtd_dependentes_21 = abs(tt_qtd_dependentes_21) + qtd_dependentes_21
								tt_qtd_dependentes = abs(tt_qtd_dependentes) + qtd_dependentes
								
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
								
								rs.movenext
								wend%>
							
							<tr><td colspan="100%">&nbsp;</td></tr>
							
							<tr>
								<td align="center" colspan="3">&nbsp;<b>TOTAL</b></td>
								<td align="center">&nbsp;</td>
								<td align="center">&nbsp;<b><%=tt_hextra%></b></td>
								<td align="center">&nbsp;<b><%=tt_faltas_injust%></b></td>
								<td align="center">&nbsp;<b><%=tt_atrasos%></b></td>
							</tr>
							
						</table>
			