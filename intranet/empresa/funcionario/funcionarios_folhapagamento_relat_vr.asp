<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../css/geral.htm"-->
<%
cd_user = session("cd_codigo")
jan = request("jan")
if jan = "" Then jan = "0"

strcod = request("cod")
strcod = ""
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strcat = request("cat")
	if strcat = "" Then 
		nm_ordem = "nm_nome"',nm_unidade
		strcat = "empresa"
	end if
strbusca = request("busca")
stracao = request("acao")
blocos_10 = request("blocos_10")

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
	if dt_dia = "" Then dt_dia = day(now) end if
	'if dt_mes = "" Then dt_mes = month(now) 'end if
	'if dt_ano = "" Then dt_ano = year(now) 'end if
	
	if dt_mes = "" AND dt_ano = "" Then
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
	'qtd_dias_comp = datediff("d","1/"&mes_competencia&"/"&ano_competencia,final_competencia&"/"&mes_competencia&"/"&ano_competencia)'+1
	'qtd_dias_comp = datediff("d","1/9/2011","30/9/2011")

'*** Quantidade de domingos da competencia ***
	'qtd_semanas_comp = datediff("ww","1/"&mes_competencia&"/"&ano_competencia,final_competencia&"/"&mes_competencia&"/"&ano_competencia,1)
	'qtd_domingos_comp = datediff("ww","1/"&mes_competencia&"/"&ano_competencia,final_competencia&"/"&mes_competencia&"/"&ano_competencia,1)
	
	
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
			
action1 = "funcionarios_folhapagamento_relat_vr"
			
data_atual = dia_final&"/"&dt_mes&"/"&dt_ano%>
<!--include file="../../includes/feriados.asp"-->
<%qtd_dias_dsr = qtd_domingos_comp + qtd_feriados%>
<%if jan <> 1 then%><table>
	<tr class="no_print">
		<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
		<td><!--a href="javascript:void(0);" return false;" onClick="window.open('financeiro/diario_emissao_cheques3.asp?tipo=emissao_cheques&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','cheques_e','location=0,status=0,width=620,height=600,scrollbars=1')">Cheques emitidos</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('financeiro/fatura_consolidado3.asp','','location=0,status=0,width=500,height=620,scrollbars=1,resizable=1')">Fatura Consolidado</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('financeiro/fatura3.asp','','location=0,status=0,width=820,height=700,scrollbars=1,resizable=1')">Emissão de faturas</a><br /-->
			<a href="javascript:void(0);" return false;" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_relat_vt.asp?tipo=variaveis&cat=empresa&jan=1','','location=0,status=0,width=835,height=700,scrollbars=1,resizable=1')">Vale Transporte</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_var.asp?tipo=variaveis&cat=empresa&jan=1','','location=0,status=0,width=900,height=720,scrollbars=1,resizable=1')">Vari&aacute;veis mensais</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_var_contabilidade.asp?dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','contab<%=dt_ano&zero(dt_mes)%>','location=0,status=1,width=750,height=600,scrollbars=1')">Contabilidade</a>
		</td>
		<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
		<td>
<%end if%>

					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print">
						<%if session("cd_codigo") = 46 then%>
							<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">funcionarios_folhapagamento_relat_vr.asp</span></td></tr>
						<%end if%>
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="<%if jan = 1 then%><%=action1%><%else%>empresa<%end if%><%'=strcat%>.asp"  method="POST">
							<input type="hidden" name="tipo" value="<%=strtipo%>">
							<input type="hidden" name="cat" value="<%=strcat%>">
							<input type="hidden" name="cod" value="<%=strcod%>">
							<input type="hidden" name="jan" value="<%=jan%>">
							<%if strcat = "enfermagem" then
								tit_doc = "ENFERMAGEM"
							Else
								tit_doc = "FOLHA DE PAGAMENTO"
							end if%>
							<tr>
								<td colspan="4" align="center" class="cabecalho" style="background-color:black; color:white;"><%'=dt_mes&"/"&dia_final&"/"&dt_ano%> <b>VARIÁVEIS - <%=tit_doc%> <%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
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
								<td align="center" style="background-color:silver;">Blocos de 10 linhas</td>
								<td style="background-color:silver;"><input type="checkbox" name="blocos_10" value="1" <%if blocos_10 = 1 then response.write("Checked")%>/></td>
							</tr>
							<tr>
								<td align="left" colspan="1" style="background-color:silver;">&nbsp;</td>
								<td align="left" colspan="1" style="background-color:silver;"><input type="submit" name="ok" value="Visualizar"></td></tr>
						</form>
					</table>
<%if jan <> 1 then%>
			</td>
			<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>	
		</td>
	</tr>
</table>
<%end if%>					
					
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" style="border-collapse:collapse;">
					<br class="no_print"><br class="no_print">
			<!-- ---------------------------------------------------------------- -->				
							<!--tr style="color:whites;" class="no_print">
								<td align="left" colspan="32">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês &nbsp; &nbsp; &nbsp; &nbsp; (<%=(qtd_dias_comp-qtd_domingos_comp)-qtd_feriados%>)
									{<%=var_dia_semana%>}</td>
							</tr-->
							<tr bgcolor="#000000">
								<td colspan="100%" align="center" class="cabecalho" style="background-color:black; color:white;"><b>RELATÓRIO DO PAUL - REFEIÇÃO - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr style="background-color:gray; color:white;">
								<td>&nbsp;<img src="../../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
								<td align="center"><b>Mat.</b><br><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
								<td align="center"><b>Funcion&aacute;rio</b><br><img src="../../imagens/px.gif" alt="" width="265" height="1" border="0"></td>
								<td align="center"><b>Unidade</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Exped.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<!--td align="center"><b>Plantões</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Qtd H.E.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>H.E. +</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Faltas just.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>Faltas injust.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td-->
								<%'if strtipo <> "var_enf" then%><!--td align="center"><b>DSR.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td--><%'end if%>
								<!--td align="center"><b>Atrasos</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td-->
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<%'if strtipo <> "var_enf" then%><td align="center"><b>VR</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td><%'end if%>
									<td align="center"><b>VR Extra</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Total VR</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<%'if strtipo <> "var_enf" then%>
									<td align="center"><b>$VR</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>TOTAL</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
							</tr>
				<%	'*** Funcionarios +1 dia  ***
							strsql = "SELECT * FROM TBL_funcionario_plantoes_1d where dt_data = '"&dt_mes&"/1/"&dt_ano&"'"
							Set rs_1d = dbconn.execute(strsql)
								if not rs_1d.EOF then nm_1d = rs_1d("nm_funcionarios")
								'response.write(nm_1d)
					
					num_cor = 1
					num_funcionario = 1
					outras_cond = " AND cd_contrato = 1"
					conta_bloco = 1
					linhas_limite = 60
					linha = 1
					pagina = 1
					'xsql = "up_funcionario_rh_lista_individual @cd_funcionario="&strcod&", @dt_atualizacao='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
					xsql = "up_funcionario_folhapagamento_geral2 @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @outras_cond='"&outras_cond&"', @nm_ordem='"&nm_ordem&"'"
            		'xsql = "up_funcionario_contrato_lista4_inativos @dt_data_i='"&mes_competencia_i&"/1/"&ano_competencia_i&"', @dt_data_f='"&mes_competencia_f&"/"&final_competencia_f&"/"&ano_competencia_f&"', @dt_atualizacao = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @ordem_funcionarios='"&nm_ordem&"', @outras_variaveis='"&outras_cond&"'"
					'xsql = "up_funcionario_contrato_lista4_inativos @dt_data_i='"&mes_competencia_i&"/"&dia_competencia_i&"/"&ano_competencia_i&"', @dt_data_f='"&mes_competencia_f&"/"&final_competencia_f&"/"&ano_competencia_f&"', @dt_atualizacao = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @ordem_funcionarios='"&nm_ordem&"', @outras_variaveis='"&outras_cond&"'"
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
								'xsql = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod&" AND (cd_tipo = 4) AND (dt_atualizacao <= '04/26/2012') AND (nr_valor > 0)"
								xsql = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod&" AND (cd_tipo = 4) AND (dt_atualizacao <= '"&dt_mes&"/"&dt_dia&"/"&dt_ano&"') AND (nr_valor > 0)"
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
								
								'*** Carrega as variáveis do RCM ***
								xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
								Set rs_rcm = dbconn.execute(xsql)
									if not rs_rcm.EOF then
										total_ad_noturno_he = rs_rcm("total_ad_noturno")
											total_geral_ad_noturno_he = total_geral_ad_noturno_he + total_ad_noturno_he
										total_he = rs_rcm("total_he")/60
											total_geral_he = total_geral_he + total_he
										total_falta_justif = rs_rcm("total_falta_justif")
											total_geral_falta_justif = total_geral_falta_justif + total_falta_justif
										total_falta_injust = rs_rcm("total_falta_injust")
											total_geral_falta_injust = total_geral_falta_injust + total_falta_injust
										total_faltas = rs_rcm("total_faltas")
											total_geral_faltas = total_geral_faltas + total_faltas
										total_atraso = rs_rcm("total_atraso")
											total_geral_atraso = total_geral_atraso + total_atraso
										
										total_vr_extra = rs_rcm("total_vr_extra")
											'total_geral_vr_extra = total_geral_vr_extra + total_vr_extra
										total_vt_extra = rs_rcm("total_vt_extra")
											total_geral_vt_extra = total_geral_vt_extra + total_vt_extra
									end if
										total_func_ad_noturno = nr_ad_noturno + total_ad_noturno_he
											total_geral_func_ad_noturno = total_geral_func_ad_noturno + total_func_ad_noturno
												total_geral_ad_noturno = total_geral_ad_noturno + total_ad_noturno_he + nr_ad_noturno
								
								'*** Separa os valores de H.E. ***
										if total_he > 36 then
											total_he_plus = total_he - 36
											total_he_limit = 36
										else
											total_he_limit = total_he
											total_he_plus = 0
										end if
											total_geral_he_plus = total_geral_he_plus + total_he_plus
											total_geral_he_limit = total_geral_he_limit + total_he_limit
										
										
										'*** VT Extra: Ida e Volta
										if total_vt_extra <> "" then
											total_vt_extra = total_vt_extra *2 
										end if
								
								'**** Calcula a qtd de VR e VT ************
								'** expediente de 6hs, não tem VR **
								if expediente <> 6 then
									nr_vr = qtd_plantoes - total_faltas
										'total_geral_nr_vr = total_geral_nr_vr + nr_vr
										
									
								end if
								
								if qtd_plantoes <> "0" AND cd_vt_canc < 1 then
									nr_vt = (qtd_plantoes * 2)' - qtd_total_faltas
									qtd_total_faltas_vt = total_faltas * 2
									nr_vt = nr_vt - qtd_total_faltas_vt
										'total_geral_nr_vt = total_geral_nr_vt + nr_vt
								end if
								
								
									if num_cor = 1 then
										cor_linha = "#ccffff"
									else
										cor_linha = "#ffffff"
									end if%>
							
							
					<%if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
							
							if total_vr_extra <> "" Then
								total_vr = int(nr_vr) + int(total_vr_extra)
							else
								total_vr = int(nr_vr)
							end if
							
							if expediente <> 6 then 
								valor_vr = 12
								custo_vr = 12
							else
								valor_vr = 0
								custo_vr = 12
							end if
							
							
							'if total_vr > 0 then	nr_valor_total = total_vr * custo_vr
							if total_vr_extra <> "" Then
								total_vr = int(nr_vr) + int(total_vr_extra)
							else
								total_vr = int(nr_vr)
							end if
								total_geral_vr = total_geral_vr + total_vr
								total_geral_nr_vr = total_geral_nr_vr + nr_vr
								total_geral_vr_extra = total_geral_vr_extra + total_vr_extra
							
							if total_vt_extra <> "" Then
								total_vt = int(nr_vt) + int(total_vt_extra)
							else
								total_vt = int(nr_vt)
							end if
								total_geral_vt = total_geral_vt + total_vt
							
							total_custo_vr = total_vr*custo_vr
								total_geral_custo_vr = total_geral_custo_vr + total_custo_vr
							if expediente <> 6 or total_vr_extra <> 0 then
							
							conta_bloco = conta_bloco + 1%>
						<tr onmouseover="mOvr(this,'#CFC8FF');" onClick="window.open('empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>&var=1&cat=<%=strcat%>#beneficios','variaveis','width=790,height=250,scrollbars=1')" onmouseout="mOut(this,'<%=cor_linha%>');" bgcolor="<%=cor_linha%>">								
								<td align="right" ><b><%=zero(num_funcionario)%></b>&nbsp;<%'=" - "&periodo_ferias%></td>
								<td align="right"><%=cd_matricula%>&nbsp;</td>
								<td style="color:<%=cor_registro%>;" align="left"><%=nm_nome%> <%if expediente = 6 then%>**<%end if%><%'=qtd_plantoes%></td>
								<td style="color:<%=cor_registro%>;" align="center"><%=nm_sigla%></td>
								<td style="color:<%=cor_registro%>;" align="center"><%=tipo_expediente%></td>
								
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td-->
								<td align="center"><%if nr_vr >= 1 then%><%=(nr_vr)%><%else%>0<%end if%></td>
								<td align="center"><%if total_vr_extra >= 1 then%><%=int(total_vr_extra)%><%else%>0<%end if%></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><%if total_vr >= 1 then%><b><%=int(total_vr)%></b><%else%>0<%end if%></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="right"><%'=cd_beneficio_transp%><%'if nr_valor >= 1 then%><%'=formatnumber(nr_valor,2)%><%'else%><%if custo_vr > 0 then%><%=formatnumber(custo_vr,2)%><%else%>0<%end if%></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="right"><%'=nr_valor_total_sptrans%><%if total_vr >= 1 then%><%=formatnumber(total_custo_vr,2)%><%else%>0<%end if%><%'=tt_nr_valor_total_refeic%></td>
						</tr>
							<%if blocos_10 = 1 and conta_bloco > 10 then%>
								<tr><td colspan="14">&nbsp;</td></tr>
								<%conta_bloco = 1
							end if%>
						<%if linha = linhas_limite then%>
							<tr class="ok_print"><td colspan="17">&nbsp;</td></tr>
								<tr class="ok_print" style="page-break-after:always;"><td colspan="17" align="right">x&nbsp;impr.:<%=now%><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0">Página <%=pagina%></td></tr>
								<%cab = 1
								linha = 0
								pagina = pagina + 1
								if cab = 1 then%>
									<tr bgcolor="#000000"  class="ok_print">
										<td colspan="100%" align="center" class="cabecalho" style="background-color:black; color:white;"><b>RELATÓRIO DO PAUL - REFEIÇÃO - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
									</tr>
									<tr style="background-color:gray; color:white;" class="ok_print">
										<td>&nbsp;</td>
										<td align="center"><b>Mat.</b></td>
										<td align="center"><b>Funcion&aacute;rio</b></td>
										<td align="center"><b>Unidade</b></td>
										<td align="center"><b>Exped.</b></td>
											<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
										<%'if strtipo <> "var_enf" then%><td align="center"><b>VR</b></td><%'end if%>
											<td align="center"><b>VR Extra</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
											<td align="center"><b>Total VR</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
										<%'if strtipo <> "var_enf" then%>
											<td align="center"><b>$VR</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
											<td align="center"><b>TOTAL</b></td>
									</tr>
								<%cab = cab + 1
								end if%>
							<%end if%>
							
						<%linha = linha + 1
						num_funcionario = num_funcionario + 1
							num_cor =  num_cor + 1
								if num_cor > 2 then
									num_cor = 1
								end if
						end if%>
						
					<%end if%>
						
								<%
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
								
								nr_vr = "0"
								total_faltas = "0"
								total_vr_extra = "0"
								
								cd_vtransp_cancela = "0"
								nr_vtransp_cancela = "0"
								
								valor_refeicao = "0"
								
								periodo_ferias = ""
								nr_desconto_ferias = "0"
								
								
								rs.movenext
								wend%>
							<tr><td colspan="100%" bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td></tr>
							<tr>
								<td align="right" colspan="4">&nbsp;</td>
								<td align="center" colspan="1">&nbsp;<b>TOTAL</b></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<%'if strtipo <> "var_enf" then%><td align="center"><%if total_geral_nr_vr > 0 then%><b><%=total_geral_nr_vr%></b><%else%>0<%end if%></td><%'end if%>
									<td align="center">&nbsp;<%if total_geral_vr_extra > 0 then%><b><%=total_geral_vr_extra%></b><%else%>0<%end if%></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center">&nbsp;<%if total_geral_vr > 0 then%><b><%=total_geral_vr%></b><%else%>0<%end if%></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td colspan="1">&nbsp;</td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td colspan="2" style="font-size:12px;">&nbsp;<b><%if total_geral_custo_vr <> "" Then%><%=formatnumber(total_geral_custo_vr,2)%><%else%>0<%end if%></b></td>
								
							</tr>
							<tr><td colspan="14">&nbsp;</td></tr>
							<tr>
								<td colspan="14" align="center">&nbsp;<span style="font-size:9px;" class="no_print">** expediente de 6h não recebe vale refeição.</span>
								<span class="ok_print">** expediente de 6h não recebe vale refeição.<img src="../../imagens/px.gif" alt="" width="230" height="1" border="0">impr.:<%=now%><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0">Página <%=pagina%></span></td>
							</tr>
							
							
						</table>
			