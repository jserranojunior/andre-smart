<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%
cd_user = session("cd_codigo")

pasta_loc = "empresa\funcionario\"
arquivo_loc = "funcionarios_folhapagamento_trans_extra.asp"

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
	
cd_ordem = request("cd_ordem")
	if cd_ordem = "1" then
		nm_ordem = "nm_unidade,nm_nome"
	Elseif cd_ordem = "2" then
		nm_ordem = "nm_nome,nm_unidade"
	Elseif cd_ordem = "" AND strcat = "enfermagem"Then 
		nm_ordem = "nm_unidade,nm_nome"
		cd_ordem = "1"
	elseif cd_ordem = "" AND strcat = "empresa" Then 
		nm_ordem = "nm_nome,nm_unidade"
		cd_ordem = "2"
	end if

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
<!--#include file="../../includes/arquivo_loc.asp"-->
<%qtd_dias_dsr = qtd_domingos_comp + qtd_feriados%>


					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print">
						
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="empresa.asp"  method="POST">
							<input type="hidden" name="tipo" value="<%=strtipo%>">
							<input type="hidden" name="cat" value="beneficios">
							<input type="hidden" name="cod" value="<%=strcod%>">
							<%if strcat = "enfermagem" then
								tit_doc = "ENFERMAGEM"
							Else
								tit_doc = "FOLHA DE PAGAMENTO"
							end if%>
							<tr>
								<td colspan="4" align="center" class="cabecalho" style="background-color:black; color:white;"><%'=data_atual%> <b>BENEF&Iacute;CIOS - Transporte (horas extras)</b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td align="center" style="background-color:silver;"><b>Mês para Pagamento:</b><br><img src="../../imagens/px.gif" alt="" width="150" height="0" border="0"></td>
								<td align="left" style="background-color:silver;" colspan="3">
									<select name="dt_mes">
										<%for i = 1 to 12%>
											<option value="<%=i%>" <%if i = int(dt_mes) then%>SELECTED<%end if%>><%=mesdoano(i)%></option>
										<%next%>
									</select> / <input type="text" name="dt_ano" value="<%=dt_ano%>" size="4" maxlength="4"></td>
							</tr>
							<tr>
								<td align="center" style="background-color:silver;">Mês de competência:</td>
								<td align="left" style="background-color:silver; color:red;"><b><%=ucase(mesdoano(mes_competencia))%>/<%=ano_competencia%></b></td>
							</tr>
							<tr>
								<td align="center" style="background-color:silver;">Unidade</td>
								<td><%strsql ="TBL_unidades where cd_hospital > 0 and cd_status = 1"
								  	Set rs_uni = dbconn.execute(strsql)%>
									<select name="cd_unidade" class="inputs"> 
										<option value="">Todas</option>
										<%Do While Not rs_uni.eof%>
										<option value="<%=rs_uni("cd_codigo")%>" <%if int(strcd_unidade) = rs_uni("cd_codigo") then response.write("SELECTED")%>><%=rs_uni("nm_Unidade")%></option>
										<%rs_uni.movenext
										loop
										rs_uni.close
										Set rs_uni = nothing%>
									</select></td>
							</tr>
							
							<tr>
								<td align="center" style="background-color:silver;">Ordem</td>
								<td style="background-color:silver;"><select name="cd_ordem" class="inputs">
										<option value="1" <%if int(cd_ordem) = 1 then response.write("SELECTED")%>>Unidade</option>
										<option value="2" <%if int(cd_ordem) = 2 then response.write("SELECTED")%>>Alfabética</option>
									</select></td>
							</tr>
							<tr>
								<td align="left" colspan="1" style="background-color:silver;">&nbsp;</td>
								<td align="left" colspan="1" style="background-color:silver;"><input type="submit" name="ok" value="Visualizar"> &nbsp; &nbsp; <a href="javascript:void(0);" return false;" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_var_contabilidade.asp?dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','nome','location=0,status=1,width=750,height=600,scrollbars=1')">Contabilidade</a></td></tr>
						</form>
					</table>
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" style="border-collapse:collapse;">
					<br class="no_print"><br class="no_print">
			<!-- ---------------------------------------------------------------- -->				
							<!--tr style="color:whites;" class="no_print">
								<td align="left" colspan="32">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês &nbsp; &nbsp; &nbsp; &nbsp; (<%=(qtd_dias_comp-qtd_domingos_comp)-qtd_feriados%>)
									{<%=var_dia_semana%>}</td>
							</tr-->
							<tr bgcolor="#000000">
								<td colspan="100%" align="center" class="cabecalho" style="background-color:black; color:white;"><b>Relatório Mensal - Transporte (horas extras)<br><%=mes_selecionado(dt_mes)%>/<%=dt_ano%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td colspan="6">&nbsp;</td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" colspan="2">SPTRANS</td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" colspan="2">CMT SP</td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" colspan="2">TRANSP Diadema</td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td colspan="1">&nbsp;</td>
							</tr>
							<tr style="background-color:gray; color:white;">
								<td>&nbsp;<img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
								<td align="center"><b>Mat.</b><br><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
								<td align="center"><b>Funcion&aacute;rio</b><br><img src="../../imagens/px.gif" alt="" width="225" height="1" border="0"></td>
								<td align="center"><b>Unidade</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>VT Extra</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
									<td align="center"><b>Valor</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td align="center"><b>Total</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Valor</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td align="center"><b>Total</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Valor</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td align="center"><b>Total</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Total Geral</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								
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
					outras_cond = " AND cd_contrato = 1"
					'xsql = "up_funcionario_rh_lista_individual @cd_funcionario="&strcod&", @dt_atualizacao='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
					xsql = "up_funcionario_folhapagamento_geral @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @outras_cond='"&outras_cond&"', @nm_ordem='"&nm_ordem&"'"
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
								
								'***** VT Cancelado **********************************************
								xsql = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod&" AND (cd_tipo = 4) AND (dt_atualizacao <= '04/26/2012') AND (nr_valor > 0)"
								Set rs_vt_canc = dbconn.execute(xsql)
									if not rs_vt_canc.EOF Then
										cd_vt_canc = rs_vt_canc("nr_valor")
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
								
								'cd_vrefeic_cancela = rs("cd_vrefeic_cancela")
								'nr_vrefeic_cancela = rs("nr_vrefeic_cancela")
								'	if  not IsNumeric(nr_vrefeic_cancela) Then nr_vrefeic_cancela = "0"
								
								'cd_vtransp_cancela = rs("cd_vtransp_cancela")
								'nr_vtransp_cancela = rs("nr_vtransp_cancela")
								'	if  not IsNumeric(nr_vtransp_cancela) Then nr_vtransp_cancela = "0"
								
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
								
								if qtd_plantoes <> "0" AND cd_vt_canc < 1 then
									nr_vt = (qtd_plantoes * 2)' - qtd_total_faltas
									qtd_total_faltas_vt = qtd_total_faltas * 2
									nr_vt = nr_vt - qtd_total_faltas_vt
								end if
								
								cd_vr_extra = rs("cd_vr_extra")
								'nr_vr_extra = rs("nr_vr_extra")
									if  not IsNumeric(nr_vr_extra) Then nr_vr_extra = "0"
									
								cd_vt_extra = rs("cd_vt_extra")
								nr_vt_extra = rs("nr_vt_extra")
									if  not IsNumeric(nr_vt_extra) Then 
										nr_vt_extra = "0"
									else
										dias_nr_vt_extra = nr_vt_extra * 2
										total_geral_vt_extra = total_geral_vt_extra + int(nr_vt_extra)
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
							
							
							%>
							<!--tr onmouseover="mOvr(this,'#CFC8FF');" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_var_cad.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>&var=1&cat=<%=strcat%>','variaveis','width=535,height=250,scrollbars=1')" onmouseout="mOut(this,'<%=cor_linha%>');" bgcolor="<%=cor_linha%>"-->
							<tr onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'<%=cor_linha%>');" bgcolor="<%=cor_linha%>">
								<td align="right" ><b><%=zero(num_funcionario)%></b>&nbsp;<%'=" - "&periodo_ferias%></td>
								<td align="right"><%=cd_matricula%>&nbsp;</td>
								<td style="color:<%=cor_registro%>;" align="left"><%=nm_nome%> <%'=cd_vt_canc%></td>
								<td style="color:<%=cor_registro%>;" align="center"><%=nm_sigla%></td>
											
					
					
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="right"><%if nr_vt_extra >= 1 then%><%=zero(nr_vt_extra)%><%else%>0<%end if%>
							<%strsql = "SELECT * FROM VIEW_funcionario_beneficios where cd_funcionario='"&strcod&"' AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' AND cd_beneficio_tipo=2  order by dt_atualizacao desc"
								SET rs_ben = dbconn.execute(strsql)
								While not rs_ben.EOF
									
									strcd_beneficio = rs_ben("cd_beneficio")
									strcd_fornecedor = rs_ben("cd_fornecedor")
									
									'if strcd_beneficio <> "" AND strcd_fornecedor <> "" Then
									if strcd_beneficio <> 0 then
										array_beneficio = array_beneficio&","&strcd_beneficio&":"&strcd_fornecedor
									end if
										
												
								rs_ben.movenext
								wend
									
									array_beneficio = mid(array_beneficio,2,len(array_beneficio))
										array_beneficio = split(array_beneficio,",")
										
										For Each beneficio_junto In array_beneficio
											conta_col = conta_col + 1
											separador = instr(1,beneficio_junto,":",1)
											str_beneficio = mid(beneficio_junto,1,separador)
												str_beneficio = replace(str_beneficio,":","")
											str_fornecedor = mid(beneficio_junto,separador,len(beneficio_junto))
												str_fornecedor = replace(str_fornecedor,":","")
												
													str_sql = "SELECT * FROM View_Beneficios_valores WHERE cd_codigo = "&str_beneficio&" AND cd_fornecedor="&str_fornecedor&" AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"'"
													SET rs_val = dbconn.execute(str_sql)
													
													if not rs_val.EOF then
														nr_valor = rs_val("nr_valor")
														valor_diario = nr_valor*2
														valor_total = valor_diario * nr_vt_extra
														valor_total_mes = valor_total_mes + valor_total
														valor_extra_total_mes = valor_extra_total_mes + valor_total_mes
														if conta_col = 1 then
															total_sptrans = total_sptrans + valor_total
														elseif conta_col = 2 then
															total_cmt = total_cmt + valor_total
														elseif conta_col = 3 then
															total_dia = total_dia + valor_total
														end if
													end if
													
													
											%>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
											<td align="right"><%=formatnumber(valor_diario,2)%></td>
											<td align="right"><%if valor_total <> 0 Then response.write(formatnumber(valor_total,2))%> </td>
										<%
										next
											while not conta_col = 3%>
											<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
												<td>&nbsp;<%'=conta_col%></td>
												<td>&nbsp;<%'=conta_col%></td>
											<%conta_col = conta_col + 1
											wend %>
					
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="right"><%if valor_total_mes <> 0 Then response.write("<b>"&formatnumber(valor_total_mes,2)&"</b>")%></td>
							</tr>
						<%num_funcionario = num_funcionario + 1
						end if%>
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
								
								total_vr = "0"
								total_vt = "0"
								
								qtd_total_faltas = "0"
								cd_vt_canc = "0"
								
								array_beneficio = ""
								conta_col = 0
								valor_diario = 0
								valor_total_mes = 0
								
								num_cor =  num_cor + 1
								if num_cor > 2 then
									num_cor = 1
								end if
								
								rs.movenext
								wend%>
							<tr><td colspan="100%" bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td></tr>
							<tr>
								<td align="center" colspan="4">&nbsp;<b>TOTAL</b></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="right">&nbsp;<b><%=total_geral_vt_extra%></b></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="right">&nbsp;</td>
									<td align="right">&nbsp;<b><%if total_sptrans > 0 then%><%=formatnumber(total_sptrans)%><%else%>0<%end if%></b></td>									
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
									<td align="right">&nbsp;</td>
									<td align="right">&nbsp;<b><%if total_cmt > 0 then%><%=formatnumber(total_cmt)%><%else%>0<%end if%></b></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
									<td align="right">&nbsp;</td>
									<td align="right">&nbsp;<b><%if total_dia > 0 then%><%=formatnumber(total_dia)%><%else%>0<%end if%></b></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="right">&nbsp;<b><%'=valor_extra_total_mes%><%=formatnumber(total_sptrans+total_cmt+total_dia)%></b></td>
									
							</tr>
							
						</table>
			