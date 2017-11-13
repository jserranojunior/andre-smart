<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../css/geral.htm"-->
<%
cd_user = session("cd_codigo")
pasta_loc = "empresa\funcionario\"
arquivo_loc = "funcionarios_folhapagamento_relat_vt.asp"



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
		
		mes_ant = mes_competencia - 1
		ano_ant = ano_competencia
			if mes_ant < 1 then
				mes_ant = 12
				ano_ant = ano_ant - 1
			end if

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
			
action1 = "funcionarios_folhapagamento_relat_vt"

data_atual = dia_final&"/"&dt_mes&"/"&dt_ano%>
<!--include file="../../includes/feriados.asp"-->
<!--#include file="../../includes/arquivo_loc.asp"-->
<%qtd_dias_dsr = qtd_domingos_comp + qtd_feriados%>

<%if jan <> 1 then%>
<table>
	<tr class="no_print">
		<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
		<td><!--a href="javascript:void(0);" return false;" onClick="window.open('financeiro/diario_emissao_cheques3.asp?tipo=emissao_cheques&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','cheques_e','location=0,status=0,width=620,height=600,scrollbars=1')">Cheques emitidos</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('financeiro/fatura_consolidado3.asp','','location=0,status=0,width=500,height=620,scrollbars=1,resizable=1')">Fatura Consolidado</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('financeiro/fatura3.asp','','location=0,status=0,width=820,height=700,scrollbars=1,resizable=1')">Emissão de faturas</a><br /-->
			<a href="javascript:void(0);" return false;" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_relat_vr.asp?tipo=variaveis&cat=empresa&jan=1','','location=0,status=0,width=755,height=700,scrollbars=1,resizable=1')">Vale Refei&ccedil;&atilde;o</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_var.asp?tipo=variaveis&cat=empresa&jan=1','','location=0,status=0,width=820,height=700,scrollbars=1,resizable=1')">Vari&aacute;veis mensais</a><br />
			<a href="javascript:void(0);" return false;" onClick="window.open('empresa/funcionario/funcionarios_folhapagamento_var_contabilidade.asp?dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','contab<%=dt_ano&zero(dt_mes)%>','location=0,status=1,width=750,height=600,scrollbars=1')">Contabilidade</a>
		</td>
		<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
		<td>
<%end if%>
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print">
						
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
								<td colspan="4" align="center" class="cabecalho" style="background-color:black; color:white;"><%'=dt_mes&"/"&dia_final&"/"&dt_ano%> <b>RELAT&Oacute;RIO DE VALE TRANSPORTE  <%'=tit_doc%> <%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
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
								<td align="left" colspan="1" style="background-color:silver;"><input type="submit" name="ok" value="Visualizar"> &nbsp; &nbsp; </td></tr>
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
								<td colspan="100%" align="center" class="cabecalho" style="background-color:black; color:white;"><b>RELATÓRIO DO PAUL - TRANSPORTES - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr style="background-color:gray; color:white;">
								<td>&nbsp;<img src="../../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
								<td align="center"><b>Mat.</b><br><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
								<td align="center"><b>Funcion&aacute;rio</b><br><img src="../../imagens/px.gif" alt="" width="215" height="1" border="0"></td>
								<td align="center"><b>Unidade</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<!--td align="center"><b>Plantões</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Qtd H.E.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>H.E. +</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Faltas just.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>Faltas injust.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td-->
								<%'if strtipo <> "var_enf" then%><!--td align="center"><b>DSR.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td--><%'end if%>
								<!--td align="center"><b>Atrasos</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td><br />
								<%'if strtipo <> "var_enf" then%><td align="center"><b>VR</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td><%'end if%>
									<td align="center"><b>VR Extra</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td align="center"><b>Total VR</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td-->
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<%'if strtipo <> "var_enf" then%><td align="center"><b>VT</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td><%'end if%>
									<td align="center"><b>VT Extra</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Total VT</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="5" height="1" border="0"></td>
									<td align="center"><b>Sptrans</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td align="center"><b>CMT<br>BOM</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="5" height="1" border="0"></td>
									<td align="center"><b>Total Sptrans</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td align="center"><b>Total CMT</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>TOTAL</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
							</tr>
							<%'**************************
							'*** VALOR SALÁRIO MÍNIMO ***
							xsql = "SELECT * FROM TBL_funcionario_valores_padroes WHERE (cd_tipo = 12) AND dt_atualizacao <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
							'Set rs = dbconn.execute(xsql)
							'nr_salminimo = rs("nr_valor")
							'****************************
							'*** VALOR Aux. Creche	  ***
							xsql = "SELECT * FROM TBL_funcionario_valores_padroes WHERE (cd_tipo = 6) AND dt_atualizacao <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
							'Set rs = dbconn.execute(xsql)
							'nr_auxcreche = rs("nr_valor")
							'******************************
							'*** Funcionarios +1 dia  ***
							strsql = "SELECT * FROM TBL_funcionario_plantoes_1d where dt_data = '"&dt_mes&"/1/"&dt_ano&"'"
							Set rs_1d = dbconn.execute(strsql)
								if not rs_1d.EOF then nm_1d = rs_1d("nm_funcionarios")
								'response.write(nm_1d)

					num_cor = 1
					num_funcionario = 1
					conta_bloco = 1
					outras_cond = " AND cd_contrato = 1"
					linhas_limite = 60
					linha = 1
					pagina = 1
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
								'xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
								xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_ant&"/21/"&ano_ant&"', @dt_f='"&mes_competencia&"/20/"&ano_competencia&"'"
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
											total_geral_vr_extra = total_geral_vr_extra + total_vr_extra
										total_vt_extra = rs_rcm("total_vt_extra")
											total_geral_vt_extra = total_geral_vt_extra + total_vt_extra
									end if
										total_func_ad_noturno = nr_ad_noturno + total_ad_noturno_he
											total_geral_func_ad_noturno = total_geral_func_ad_noturno + total_func_ad_noturno
												total_geral_ad_noturno = total_geral_ad_noturno + total_ad_noturno_he + nr_ad_noturno
								
								
								
								
								
								
								if qtd_plantoes <> "0" AND cd_vt_canc < 1 then
									nr_vt = (qtd_plantoes)' * 2)' - qtd_total_faltas
									'nr_vt = (qtd_plantoes)' - qtd_total_faltas
									qtd_total_faltas_vt = qtd_total_faltas' * 2
									nr_vt = nr_vt - qtd_total_faltas_vt
								end if
								
									'*** Soma as qtds de VT e VT Extra ***
									total_vtn_vte = nr_vt + total_vt_extra
								
								'*** Valor transp - SPTRANS ******************************************
								strsql = "SELECT * FROM TBL_funcionario_beneficios where cd_funcionario = '"&strcod&"' AND cd_beneficio_tipo = 2  and cd_beneficio between 1 and 3 order by dt_atualizacao desc"
									SET rs_benef = dbconn.execute(strsql)
									while not rs_benef.EOF
										cd_beneficio_transp_sptrans = rs_benef("cd_beneficio")
									
										
										strsql = "SELECT top 3 * FROM View_Beneficios_valores WHERE cd_codigo="&cd_beneficio_transp_sptrans&" ORDER BY dt_atualizacao DESC"
										SET rs_transp = dbconn.execute(strsql)
										'while not rs_transp.EOF
										if not rs_transp.EOF then
											nm_beneficio = rs_transp("nm_beneficio")
											nm_categoria = rs_transp("nm_categoria")
											nr_valor_sptrans = rs_transp("nr_valor")*2
											dt_atualizacao = rs_transp("dt_atualizacao")
												if cor = 1 then
													color_transp = "#008000"
													cor = cor + 1
												else
													color_transp = "silver"
												end if
										end if%>
									<%rs_benef.movenext
									wend		
								'**********************************************************
								'*** Valor transp - CMT ******************************************
								strsql = "SELECT * FROM TBL_funcionario_beneficios where cd_funcionario = '"&strcod&"' AND cd_beneficio_tipo = 2  AND cd_beneficio between 4 and 10 order by dt_atualizacao desc"
									SET rs_benef = dbconn.execute(strsql)
									while not rs_benef.EOF
										cd_beneficio_transp_cmt = rs_benef("cd_beneficio")
									
										
										strsql = "SELECT top 3 * FROM View_Beneficios_valores WHERE cd_codigo="&cd_beneficio_transp_cmt&" ORDER BY dt_atualizacao DESC"
										SET rs_transp = dbconn.execute(strsql)
										'while not rs_transp.EOF
										if not rs_transp.EOF then
											nm_beneficio = rs_transp("nm_beneficio")
											nm_categoria = rs_transp("nm_categoria")
											nr_valor_cmt = rs_transp("nr_valor")*2
											dt_atualizacao = rs_transp("dt_atualizacao")
												if cor = 1 then
													color_transp = "#008000"
													cor = cor + 1
												else
													color_transp = "silver"
												end if
										end if%>
									<%rs_benef.movenext
									wend		
								'**********************************************************
								
									if num_cor = 1 then
										cor_linha = "#ccffff"
									else
										cor_linha = "#ffffff"
									end if%>
							
							
					<%if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
							
							nr_valor_total_sptrans = nr_valor_sptrans * total_vtn_vte
							nr_valor_total_cmt = nr_valor_cmt * total_vtn_vte
							
							nr_valor_total = nr_valor_total_sptrans + nr_valor_total_cmt
							
							tt_valor_total_sptrans = tt_valor_total_sptrans + nr_valor_total_sptrans
							tt_valor_total_cmt = tt_valor_total_cmt + nr_valor_total_cmt
							tt_nr_valor_total_transp = tt_nr_valor_total_transp + nr_valor_total
							
						'if total_vtn_vte <> 0 then
							
							conta_bloco = conta_bloco + 1%>
							<tr onmouseover="mOvr(this,'#CFC8FF');" onClick="window.open('empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>&var=1&cat=<%=strcat%>#beneficios','variaveis','width=790,height=250,scrollbars=1')" onmouseout="mOut(this,'<%=cor_linha%>');" bgcolor="<%=cor_linha%>">								
								<td align="right" ><b><%=zero(num_funcionario)%></b>&nbsp;<%'=" - "&periodo_ferias%></td>
								<td align="right"><%=cd_matricula%>&nbsp;</td>
								<td style="color:<%=cor_registro%>;" align="left"><%=nm_nome%> <%if cd_vt_canc > 0 then response.write("*")%></td>
								<td style="color:<%=cor_registro%>;" align="center"><%=nm_sigla%></td>
								
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td-->
									<td align="center"><%if nr_vt >= 1 then%><%=(nr_vt)%><%else%>0<%end if%></td>
									<td align="center"><%if total_vt_extra >= 1 then%><%=(total_vt_extra)%><%else%>0<%end if%>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><%if total_vtn_vte >= 1 then%><%=(total_vtn_vte)%><%else%>0<%end if%></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="right"><%'=cd_beneficio_transp%><%if nr_valor_sptrans >= 1 then%><%=formatnumber(nr_valor_sptrans,2)%><%else%>0<%end if%></td>
									<td align="right"><%'=cd_beneficio_transp%><%if nr_valor_cmt >= 1 then%><%=formatnumber(nr_valor_cmt,2)%><%else%>0<%end if%></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="right"><%'=nr_valor_total_sptrans%><%if nr_valor_total_sptrans >= 1 then%><%=formatnumber(nr_valor_total_sptrans,2)%><%else%>0<%end if%></td>
									<td align="right"><%'=nr_valor_total_sptrans%><%if nr_valor_total_cmt >= 1 then%><%=formatnumber(nr_valor_total_cmt,2)%><%else%>0<%end if%></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="right"><%'=nr_valor_total_sptrans%><%if nr_valor_total >= 1 then%><%=formatnumber(nr_valor_total,2)%><%else%>0<%end if%></td>
							<!-- *** Total Liquido *** -->	
								<%'total_desconto = total_desconto + irrf_final
								'total_liquido = total_vencimento - total_desconto
								'total_liquido = total_liquido + auxcreche%>
									<!--td align="right">&nbsp;<%'if total_liquido <> 0 then response.write(formatnumber(total_liquido,2))%></td-->
							</tr>
							<%if linha = linhas_limite then%>
								<tr class="oks_print"><td colspan="17">&nbsp;</td></tr>
								<tr class="ok_print" style="page-break-after:always;"><td colspan="17" align="right">&nbsp;impr.:<%=now%><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0">Página <%=pagina%></td></tr>
								<%cab = 1
								linha = 0
								pagina = pagina + 1
								if cab = 1 then%>
									<tr bgcolor="#000000" class="ok_print">
										<td colspan="100%" align="center" class="cabecalho" style="background-color:black; color:white;"><b>RELATÓRIO DO PAUL - TRANSPORTES - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
										<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
									</tr>
									<tr style="background-color:gray; color:white;" class="ok_print">
										<td>&nbsp;</td>
										<td align="center"><b>Mat.</b></td>
										<td align="center"><b>Funcion&aacute;rio</b></td>
										<td align="center"><b>Unidade</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
										<%'if strtipo <> "var_enf" then%><td align="center"><b>VT</b></td><%'end if%>
											<td align="center"><b>VT Extra</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
											<td align="center"><b>Total VT</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="5" height="1" border="0"></td>
											<td align="center"><b>Sptrans</b></td>
											<td align="center"><b>CMT<br>BOM</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="5" height="1" border="0"></td>
											<td align="center"><b>Total Sptrans</b></td>
											<td align="center"><b>Total CMT</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
											<td align="center"><b>TOTAL</b></td>
									</tr>
								<%cab = cab + 1
								end if%>
							<%end if%>
						<%if blocos_10 = 1 and conta_bloco > 10 then%>
								<tr class="ok_print"><td colspan="17">&nbsp;</td></tr>
								<%conta_bloco = 1
							end if
							
						'end if
						
							num_funcionario = num_funcionario + 1
							num_cor =  num_cor + 1
								if num_cor > 2 then
									num_cor = 1
								end if
						end if%>
								<%linha = abs(linha) + 1
								tt_salario = abs(tt_salario) + nr_salario
								tt_aj_custo = abs(tt_aj_custo) + nr_aj_custo
								tt_insalubridade = abs(tt_insalubridade) + insalubridade
								if qtd_dia_noturno <> "" Then tt_qtd_dia_noturno = abs(tt_qtd_dia_noturno) + qtd_dia_noturno
								tt_ad_noturno = abs(tt_ad_noturno) + ad_noturno
								tt_dsr_ad_not = abs(tt_dsr_ad_not) + dsr_ad_not
								tt_qtd_dependentes_21 = abs(tt_qtd_dependentes_21) + qtd_dependentes_21
								tt_qtd_dependentes = abs(tt_qtd_dependentes) + qtd_dependentes
								
								'tt_nr_vr_extra = abs(tt_nr_vr_extra) + nr_vr_extra
								'tt_final_vr = abs(tt_final_vr) + total_vr
								
								'tt_nr_vt_extra = abs(tt_nr_vt_extra) + nr_vt_extra
								'tt_final_vt = abs(tt_final_vt) + total_vt
								'tt_total_vr = tt_total_vr + total_vr
								
								
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
								
								total_vt_extra = "0"
								
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
								
								cd_beneficio_transp_sptrans = "0"
								cd_beneficio_transp_cmt = "0"
								nr_valor_sptrans = "0"
								nr_valor_cmt = "0"
								
								nr_vr = "0"
								nr_vt = "0"
								
								total_vr = "0"
								total_vt = "0"
								
								qtd_total_faltas = "0"
								cd_vt_canc = "0"
								
								
								
								rs.movenext
								wend%>
							<tr><td colspan="100%" bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td></tr>
							<tr>
								<td align="right" colspan="3">&nbsp;</td>
								<td align="center" colspan="1">&nbsp;<b>TOTAL</b></td>
									<td align="right" bgcolor="silver">&nbsp;</td>
									<td align="center" colspan="7">&nbsp;</td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="right">&nbsp;<b><%=formatnumber(tt_valor_total_sptrans,2)%></b></td>
								<td align="right">&nbsp;<b><%=formatnumber(tt_valor_total_cmt,2)%></b></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td colspan="2" style="font-size:12px;">&nbsp;<b><%if tt_nr_valor_total_transp <> "" Then%><%=formatnumber(tt_nr_valor_total_transp,2)%><%else%>0<%end if%></b></td>
								
							</tr>
							<tr><td colspan="17">&nbsp;</td></tr>
							<tr><td colspan="17" align="right">&nbsp;<span style="font-size:9px;" class="no_print">* optou por não receber vale transporte</span>
							<span class="ok_print">* optou por não receber vale transporte<img src="../../imagens/px.gif" alt="" width="230" height="1" border="0">impr.:<%=now%><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0">Página <%=pagina%></span></td></tr>
						</table>
			