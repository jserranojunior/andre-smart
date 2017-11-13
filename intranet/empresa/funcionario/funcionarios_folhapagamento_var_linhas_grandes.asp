<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../css/geral.htm"-->
<%
cd_user = session("cd_codigo")
pasta_loc = "empresa\funcionario\"
arquivo_loc = "funcionarios_folhapagamento_var.asp"

jan = request("jan")
if jan = "" Then jan = "0"
if jan = 1 then caminho = "../../"
mostrar_contrato = request("mostrar_contrato")

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
	
mostra_faltas = request("mostra_faltas")
	if mostra_faltas = "" Then mostra_faltas = 0
	
cd_ordem = request("cd_ordem")
	if cd_ordem = "1" then
		nm_ordem = "nm_unidade,nm_nome"
	Elseif cd_ordem = "2" then
		nm_ordem = "nm_nome,nm_unidade"
	Elseif cd_ordem = "3" then
		nm_ordem = "soma_faltas_justif"
	Elseif cd_ordem = "4" then
		nm_ordem = "soma_faltas_injust"
	Elseif cd_ordem = "5" then
		nm_ordem = "soma_faltas_geral"
	Elseif cd_ordem = "6" then
		nm_ordem = "expediente,nm_nome"
		
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

blocos_10 = request("blocos_10")

if strtipo = "variaveis" then
	form_bgcolor = silver
elseif strtipo = "" Then

elseif strtipo = "" Then

end if
			
data_atual = dia_final&"/"&dt_mes&"/"&dt_ano%>
<!--include file="../../includes/feriados.asp"-->
<%qtd_dias_dsr = qtd_domingos_comp + qtd_feriados%>
<!--#include file="../../includes/arquivo_loc.asp"-->

					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print">
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="<%if jan = 1 then%><%="funcionarios_folhapagamento_var"%><%else%><%=strcat%><%end if%>.asp"  method="POST">
							<!--input type="hidden" name="tipo" value="<%'=strtipo%>"-->
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
								<td><%strsql ="TBL_unidades where cd_hospital > 0 and cd_status = 1 and cd_hospital < 3 OR cd_hospital > 0 and cd_status = 1 and cd_hospital = 4"
								  	Set rs_uni = dbconn.execute(strsql)%>
									<select name="cd_unidade" class="inputs"> 
										<option value="0">Todas</option>
										<%Do While Not rs_uni.eof%>
										<option value="<%=rs_uni("cd_codigo")%>" <%if int(strcd_unidade) = rs_uni("cd_codigo") then response.write("SELECTED")%>><%=rs_uni("nm_Unidade")%></option>
										<%rs_uni.movenext
										loop
										rs_uni.close
										Set rs_uni = nothing%>
									</select></td>
							</tr>
							<tr>
								<td align="center" style="background-color:silver;">Faltas</td>
								<td style="background-color:silver;"><select name="mostra_faltas" class="inputs">
									<option value="0" <%if mostra_faltas = "" OR mostra_faltas = 0 then response.write("SELECTED")%> />Todos<br />
									<option value="1" <%if mostra_faltas = 1 then response.write("SELECTED")%> />Sem faltas<br />
									<option value="2" <%if mostra_faltas = 2 then response.write("SELECTED")%> />Com faltas
								</select></td>
							</tr>
							<tr>
								<td align="center" style="background-color:silver;">Ordem</td>
								<td style="background-color:silver;"><select name="cd_ordem" class="inputs">
										<option value="1" <%if int(cd_ordem) = 1 then response.write("SELECTED")%>>Unidade</option>
										<option value="2" <%if int(cd_ordem) = 2 then response.write("SELECTED")%>>Alfabética</option>
										<option value="3" <%if int(cd_ordem) = 3 then response.write("SELECTED")%>>Faltas Justif</option>
										<option value="4" <%if int(cd_ordem) = 4 then response.write("SELECTED")%>>Faltas injust</option>
										<option value="5" <%if int(cd_ordem) = 5 then response.write("SELECTED")%>>Total faltas</option>
										<option value="6" <%if int(cd_ordem) = 6 then response.write("SELECTED")%>>Carga Horária</option>
									</select></td>
							</tr>
							<tr>
								<td align="center" style="background-color:silver;">Tipo</td>
								<td style="background-color:silver;">
									<input type="radio" name="tipo" value="variaveis" <%if strtipo = "variaveis" OR strtipo = "var_enf" then response.write("Checked")%>/>Variáveis &nbsp; &nbsp;
									<input type="radio" name="tipo" value="relat_vr" <%if strtipo = "relat_vr" then response.write("Checked")%>/>VR &nbsp; &nbsp; &nbsp;
									<input type="radio" name="tipo" value="relat_vt" <%if strtipo = "relat_vt"then response.write("Checked")%>/>VT
								</td>
							</tr>
							<%if strtipo <> "variaveis" then%>
							<tr>
								<td align="center" style="background-color:silver;">Blocos de 10 linhas</td>
								<td style="background-color:silver;"><input type="checkbox" name="blocos_10" value="1" <%if blocos_10 = 1 then response.write("Checked")%>/></td>
							</tr>
							<%end if%>
							<tr>
								<td align="center" style="background-color:silver;">Contratos</td>
								<td style="background-color:silver;">
										<input type="radio" name="mostrar_contrato" value="0" <%if mostrar_contrato = 0 then response.write("checked")%> /> Ambos &nbsp; &nbsp;
										<input type="radio" name="mostrar_contrato" value="1" <%if mostrar_contrato = 1 then response.write("checked")%> /> Ativos &nbsp; &nbsp;
										<input type="radio" name="mostrar_contrato" value="2" <%if mostrar_contrato = 2 then response.write("checked")%> /> Inativos
								</td>
							</tr>
								
								<td align="left" colspan="1" style="background-color:silver;">&nbsp;</td>
								<td align="left" colspan="1" style="background-color:silver;"><input type="submit" name="ok" value="Visualizar"> &nbsp; &nbsp; <a href="javascript:void(0);" return false;" onClick="window.open('<%if jan = 1 then%><%else%>empresa/funcionario/<%end if%>funcionarios_folhapagamento_var_contabilidade.asp?dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','contab<%=dt_ano&zero(dt_mes)%>','location=0,status=1,width=750,height=600,scrollbars=1')">Contabilidade</a></td></tr>
						</form>
					</table>
					<br class = "no_print">
			<%if mostrar_contrato = 0 OR mostrar_contrato = 1 Then%>
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" style="border-collapse:collapse;">
					<!--br class="no_print"><br class="no_print"-->					
			<!-- ---------------------------------------------------------------- -->				
						<%if strtipo = "variaveis" OR  strtipo = "var_enf" then%>
							<tr bgcolor="#000000">
								<td colspan="30" align="center" class="cabecalho" style="background-color:black; color:white;"><b>Horas Extras / Ad. Noturno / Faltas e atrasos - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
							</tr>
						<%elseif strtipo = "relat_vr" then%>
							<tr bgcolor="#000000">
								<td colspan="30" align="center" class="cabecalho" style="background-color:black; color:white;"><b>RELATÓRIO DO PAUL - REFEIÇÃO - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
							</tr>
						<%elseif strtipo = "relat_vt" then%>
							<tr bgcolor="#000000">
								<td colspan="30" align="center" class="cabecalho" style="background-color:black; color:white;"><b>RELATÓRIO DO PAUL - TRANSPORTE - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
							</tr>
						<%end if%>
							
							<!--tr>
								<td colspan="23"style="background-color:gray;">&nbsp;</td>
								<td colspan="7" align="center" style="background-color:gray; color:white;">VR</td>
								<td colspan="11" align="center" style="background-color:gray; color:white;">VT</td>
								
							</tr-->
							<tr style="background-color:gray; color:white;">
								<td>&nbsp;<img src="../../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
								<td align="center"><b>Mat.</b><br><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
								<td align="center"><b>Funcion&aacute;rio</b><br><img src="../../imagens/px.gif" alt="" width="205" height="1" border="0"></td>
								<td align="center"><b>Hora</b><br><img src="../../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
								<td align="center"><b>Unid.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
						<%if strtipo = "variaveis" OR strtipo = "var_enf" Then%>
								<td align="center"><b>Dias a<br>Trabalhar</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
						<%end if%>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
						<%if strtipo = "variaveis" OR strtipo = "var_enf" Then%>
								<td align="center"><b>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Ad. Not<br>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Total<br>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Qtd H.E.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td align="center"><b>H.E.+</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td align="center"><b>Total<br>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Folga Enfer</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Faltas just.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td align="center"><b>Faltas injust.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>Total faltas</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Atrasos<br>(horas)</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
							<%end if%>
								<%'if strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vr" Then
								if strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vr" Then%>
									<td align="center"><b>VR</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>VR Extra</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>Total VR</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<%end if%>
								<%if strtipo = "relat_vr" Then%>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>$ VR</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>TOTAL</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<%end if%>
								<%if strtipo = "variaveis" OR strtipo = "var_enf" Then%>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<%end if%>
								<%if strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vt" Then%>
									<td align="center"><b>VT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>VT Extra</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>Total VT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<%end if%>
								<%if strtipo = "relat_vt" Then%>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
										<td align="center"><b>Sptrans</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
										<td align="center"><b>CMT<br>BOM</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
											<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
										<td align="center"><b>Total<br>Sptrans</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
										<td align="center"><b>Total<br>CMT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
											<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
										<td align="center"><b>TOTAL</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<%end if%>
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
							'****************************
							'*** Funcionarios +1 dia  ***
							strsql = "SELECT * FROM TBL_funcionario_plantoes_1d where dt_data = '"&dt_mes&"/1/"&dt_ano&"'"
							Set rs_1d = dbconn.execute(strsql)
								if not rs_1d.EOF then nm_1d = rs_1d("nm_funcionarios")
								'response.write(nm_1d)

					num_cor = 1
					num_funcionario = 1
					linhas_limite = 40
					linha = 0
					pagina = 1
					
					ano_competencia_i = ano_competencia
					mes_competencia_i = mes_competencia
					dia_competencia_i = 25
					
					ano_competencia_f = ano_competencia
					mes_competencia_f = mes_competencia
					final_competencia_f = final_competencia
					
					dt_ano = dt_ano
					dt_mes = dt_mes
					dia_final = dia_final
					
					cd_vt_canc = 0
					
					'outras_cond = " AND cd_contrato = 1"
					
					xsql = "up_funcionario_contrato_lista4 @dt_data_i='"&mes_competencia_i&"/"&dia_competencia_i&"/"&ano_competencia_i&"', @dt_data_f='"&mes_competencia_f&"/"&final_competencia_f&"/"&ano_competencia_f&"', @dt_atualizacao = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @ordem_funcionarios='"&nm_ordem&"', @outras_variaveis='"&outras_cond&"'"
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
								
								soma_faltas_justif = rs("soma_faltas_justif")
									if ISNULL(soma_faltas_justif)  Then soma_faltas_justif = "0"
								soma_faltas_injust = rs("soma_faltas_injust")
									if ISNULL(soma_faltas_injust) Then soma_faltas_injust = 0
								soma_faltas_geral = rs("soma_faltas_geral")
									if ISNULL(soma_faltas_geral) Then soma_faltas_geral = 0
								
								dt_admissao = rs("dt_admissao_geral")
								dt_demissao = rs("dt_demissao_geral")
								
									if dt_demissao <> "null" then
										cor_registro = "red"
									else
										cor_registro = "black"
										demissao = ""
									end if
								
								'***** VT Cancelado **********************************************
								'xsql = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod&" AND (cd_tipo = 4) AND (dt_atualizacao <= '04/26/2012') AND (nr_valor > 0)"
								'xsql = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod&" AND (cd_tipo = 4) AND (dt_atualizacao <= '"&dt_mes&"/"&dt_dia&"/"&dt_ano&"') AND (nr_valor > 0)"' ORDER BY dt_atualizacao desc" 
								xsql = "SELECT top 1 nr_valor FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod&" AND (cd_tipo = 4) AND (dt_atualizacao <= '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"') ORDER BY dt_atualizacao desc" 
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
								
								'*** Expediente ***
								expediente = rs("expediente")
									
										expediente = expediente / 60
											if expediente >= 9 AND expediente <= 11 then
												tipo_expediente = "8h"
												carga_mensal = 220
												qtd_plantoes = plantoes_8h
												cor_fincionario = "black"
											elseif expediente = 6 then
												tipo_expediente = "6h"
												carga_mensal = 180
												qtd_plantoes = plantoes_6h
												cor_fincionario = "black"
											elseif expediente = 12 then
												tipo_expediente = "12x36h"
												carga_mensal = 180
												qtd_plantoes = plantoes_12h
												cor_fincionario = "black"
													
													'*** acerto +1 dia ***
													func_1d = instr(1,nm_1d,strcod,1)
													if func_1d > 0 then qtd_plantoes = qtd_plantoes + 1
											else
												tipo_expediente = "x"
												carga_mensal = 0
												qtd_plantoes = 0
												cor_fincionario = "red"
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
								xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_ant&"/21/"&ano_ant&"', @dt_f='"&mes_competencia&"/20/"&ano_competencia&"'"
								Set rs_rcm = dbconn.execute(xsql)
									if not rs_rcm.EOF then
										total_ad_noturno_he = rs_rcm("total_ad_noturno")
										'total_func_ad_noturno = nr_ad_noturno + total_ad_noturno
											
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
											
										'** Folga enfermagem **
											total_folga_enf = rs_rcm("total_folga_enf")
												if total_folga_enf = "" Then
													total_folga_enf = int("0")
												end if
										'***********************

                                    '   total_falta_justif = rs_rcm("total_falta_justif")
											
									'	total_falta_injust = rs_rcm("total_falta_injust")
											
									'	total_faltas = rs_rcm("total_faltas")
											
										total_atraso = rs_rcm("total_atraso")
											
										total_vr_extra = rs_rcm("total_vr_extra")
											
										total_vt_extra = rs_rcm("total_vt_extra")
											
									end if
										'*** Separa os valores de H.E. ***
										if total_he > 36 then
											total_he_plus = total_he - 36
											total_he_limit = 36
										else
											total_he_limit = total_he
											total_he_plus = 0
										end if
										
										
										
								
								'**** Calcula a qtd de VR e VT ************
								if expediente <> 6 then
									nr_vr = qtd_plantoes - total_faltas
									
									'***** REGRA PARA ABRIL, MAIO E JUNHO DE 2014 - DESCONTO DA FOLGA DA ENFERMAGEM ****
									if dt_ano = 2014 and dt_mes = 4 AND total_folga_enf <> 0 OR dt_ano = 2014 and dt_mes = 5 AND total_folga_enf <> 0 OR dt_ano = 2014 and dt_mes = 6 AND total_folga_enf <> 0 then
										indicativo_folga_enf = "¹"
										aviso_folga_enf = " / ¹ Em "&mes_selecionado(mes_competencia)&"/"&ano_competencia&" as folgas de Enfermagem não foram descontadas dos VR e VT "
										'nr_vr = "999"
									elseif ano_competencia = 2014 and mes_competencia = 4 AND total_folga_enf <> "" OR ano_competencia = 2014 and mes_competencia = 5 AND total_folga_enf <> "" OR ano_competencia = 2014 and mes_competencia <= 6 AND total_folga_enf <> "" then
										'indicativo_folga_enf = "¹"
										aviso_folga_enf = " / ¹ Em "&mes_selecionado(mes_competencia)&"/"&ano_competencia&" as folgas de Enfermagem não foram descontadas dos VR e VT"
										'RESPONSE.WRITE(aviso_folga_enfs)
										'nr_vr = "888"
									else
										nr_vr = nr_vr - total_folga_enf
									end if
									'***************************************************************************************
								end if
								
									'*** Elimina VT caso o benefício esteja
									nr_vt = qtd_plantoes 
									if  cd_vt_canc = 0 Then
										nr_vt = qtd_plantoes
									'elseif cd_vt_canc > 0 then
									'	nr_vt = 1000
									else
										nr_vt = 0
									end if
									
									'***** REGRA PARA ABRIL, MAIO E JUNHO DE 2014 - DESCONTO DA FOLGA DA ENFERMAGEM ****
									if dt_ano = 2014 and dt_mes = 4 AND total_folga_enf <> 0 OR dt_ano = 2014 and dt_mes = 5 AND total_folga_enf <> 0 OR dt_ano = 2014 and dt_mes = 6 AND total_folga_enf <> 0 then
										indicativo_folga_enf = "¹"
										aviso_folga_enf = " / ¹ Em "&mes_selecionado(mes_competencia)&"/"&ano_competencia&" as folgas de Enfermagem não foram descontadas dos VR e VT "
										
									elseif ano_competencia = 2014 and mes_competencia = 4 AND total_folga_enf <> "" OR ano_competencia = 2014 and mes_competencia = 5 AND total_folga_enf <> "" OR ano_competencia = 2014 and mes_competencia <= 6 AND total_folga_enf <> "" then
										aviso_folga_enf = " / ¹ Em "&mes_selecionado(mes_competencia)&"/"&ano_competencia&" as folgas de Enfermagem não foram descontadas dos VR e VT "
										
									else
										nr_vt = nr_vt - total_folga_enf
										'response.write(num_funcionario&" - "&nr_vt&" - "&total_folga_enf&"<br>")
										
									end if
									'**************************************************************************************
								'end if
								
								'*** VT Extra: Ida e Volta
									if total_vt_extra <> "" then
										'total_vt_extra = total_vt_extra
											if strtipo = "variaveis" OR strtipo = "var_enf" then
												total_vt_extra = total_vt_extra * 2 '//Duplica_VT
											end if
									end if
									
									if qtd_plantoes <> "0" AND cd_vt_canc < 1 then
									
										if strtipo = "variaveis" OR strtipo = "var_enf" then
											'nr_vt = qtd_plantoes * 2 '// Duplica_VT
											nr_vt = nr_vt * 2
										end if
										'nr_vt_ = nr_vt - total_folga_enf
										
                                    qtd_total_faltas_vt = total_faltas
										if strtipo = "variaveis" OR strtipo = "var_enf" then
											qtd_total_faltas_vt = total_faltas * 2 '// Duplica_VT
										end if
										
									nr_vt = nr_vt - qtd_total_faltas_vt
								end if
								'****** AFASTADOS ******
								if cd_unidade = 43 then
									teste = "afasta"
									'tipo_expediente = " - "
									carga_mensal = 0'220
									qtd_plantoes = 0' plantoes_6h
									cor_fincionario = "red"
									nr_vr = 0
									nr_vt = 0
									total_ad_noturno_he = 0
								end if
								
									if num_cor = 1 then
										cor_linha = "#ccffff"
									else
										cor_linha = "#ffffff"
									end if
									
									
							'***************
							'*** Filtros ***
							'***************
									if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
										cond_1 = 1
									end if
									
									if mostra_faltas <> 0 then 
										if int(mostra_faltas) = "1" AND  int(total_faltas) = 0 then
											cond_2 = 1
										elseif int(mostra_faltas) = "2" AND int(total_faltas) > 0 then
											cond_2 = 1
										end if
									else
										cond_2 = 1
									end if
									
							'*********************
							'*** Custo VR e VT ***
							'*********************
							'if expediente <> 6 then 
								valor_vr = 12
								custo_vr = 12
							'else
							'	valor_vr = 0
							'	custo_vr = 12
							'end if
							
							'**********************************************************
							'*** Valor transp - SPTRANS *******************************
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
										nr_valor_sptrans = rs_transp("nr_valor")
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
								'*** Valor transp - CMT ***********************************
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
											nr_valor_cmt = rs_transp("nr_valor")
											dt_atualizacao = rs_transp("dt_atualizacao")
												if cor = 1 then
													color_transp = "#008000"
													cor = cor + 1
												else
													color_transp = "silver"
												end if
										end if%>
									<%rs_benef.movenext
									wend%>
							
							
					<%'if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
					if cond_1 = 1 AND cond_2 = 1 then 
					'if int(filtros)="00" AND  int(total_faltas)>=0 AND int(cd_unidade) = int(strcd_unidade) OR int(filtros)="00" AND int(total_faltas)<=0 AND int(cd_unidade) = int(strcd_unidade) OR int(filtros)="00" AND  int(total_faltas)>=1 AND int(cd_unidade) = int(strcd_unidade)then
							
					
					
							if total_vr_extra <> "" Then
								total_vr = int(nr_vr) + int(total_vr_extra)
							else
								total_vr = int(nr_vr)
							end if
							
							total_custo_vr = total_vr*custo_vr
								total_geral_custo_vr = total_geral_custo_vr + total_custo_vr
							
							
							if total_vt_extra <> "" Then
								total_vt = int(nr_vt) + int(total_vt_extra)
							else
								total_vt = int(nr_vt)
							end if
							
								nr_valor_sptrans_vt = (nr_valor_sptrans * total_vt) * 2
								nr_valor_cmt_vt = (nr_valor_cmt * total_vt) * 2
									
									nr_valor_total_valetransporte = nr_valor_sptrans_vt + nr_valor_cmt_vt
									
									if nr_valor_total_valetransporte > 0 AND total_vr > 0 Then
										total_valor_total_sptrans_vt = total_valor_total_sptrans_vt + nr_valor_sptrans_vt
										total_valor_total_cmt_vt = total_valor_total_cmt_vt + nr_valor_cmt_vt
										
										total_valor_total_valetransporte = total_valor_total_valetransporte + nr_valor_total_valetransporte
									end if
									
							'*** Tota geral variaveis ***
							if total_he_plus <> "" Then total_geral_he_plus = total_geral_he_plus + total_he_plus
							if total_he_limit <> "" Then total_geral_he_limit = total_geral_he_limit + total_he_limit
							if total_he <> "" Then total_geral_he = total_geral_he + total_he
							
							if total_folga_enf <> "" Then total_geral_folga_enf = total_geral_folga_enf + int(total_folga_enf)
							
                            if total_falta_justif <> "" Then total_geral_falta_justif = total_geral_falta_justif + total_falta_justif
							if total_falta_injust <> "" Then total_geral_falta_injust = total_geral_falta_injust + total_falta_injust
							if total_faltas <> "" Then total_geral_faltas = total_geral_faltas + total_faltas
							if total_atraso <> "" Then total_geral_atraso = total_geral_atraso + total_atraso
							
							total_geral_ad_noturno_he = total_geral_ad_noturno_he + total_ad_noturno_he
							'total_geral_ad_noturno = total_geral_ad_noturno_he + total_nr_ad_noturno
								if cd_periodo = 2 then 
									total_nr_ad_noturno = total_nr_ad_noturno + nr_ad_noturno
									total_ad_noturno = total_ad_noturno_he + nr_ad_noturno - total_faltas
										total_geral_ad_noturno =  total_geral_ad_noturno + nr_ad_noturno + total_ad_noturno_he
								else 
									total_ad_noturno = total_ad_noturno_he' + nr_ad_noturno
									'total_geral_ad_noturno =  total_nr_ad_noturno + nr_ad_noturno
										total_geral_ad_noturno =  total_geral_ad_noturno + total_ad_noturno_he
								end if
								
								
							'if total_vr <> "" Then total_geral_vr = total_geral_vr + total_vr
							'total_geral_vt = total_geral_vt + total_vt
							
							if nr_vr <> "" then total_geral_nr_vr = total_geral_nr_vr + nr_vr
							if total_vr_extra <> "" Then total_geral_vr_extra = total_geral_vr_extra + total_vr_extra
							total_geral_vr = total_geral_vr + nr_vr + total_vr_extra
							
							if nr_vt <> "" then total_geral_nr_vt = total_geral_nr_vt + nr_vt
							if total_vt_extra <> "" Then total_geral_vt_extra = total_geral_vt_extra + total_vt_extra
							
							total_geral_vt = total_geral_vt + nr_vt + total_vt_extra
							
							'total_valor_sptrans_vt = (nr_valor_sptrans * total_vt) 
								'nr_valor_cmt_vt = (nr_valor_cmt * total_vt) 
								
									'total_valor_total_valetransporte = total_valor_total_valetransporte + nr_valor_total_valetransporte
									
									'total_geral_valor_sptrans_vt = total_geral_valor_sptrans_vt + total_valor_sptrans_vt
							
					'IF total_vr > 0 THEN
					IF  strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vt" AND total_vt > 0  OR strtipo = "relat_vr" AND total_vr > 0 THEN
					'IF  strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vt" OR strtipo = "relat_vr" AND total_vr > 0 THEN%>
					
						<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'white<%'=cor_linha%>');" bgcolor="white<%'=cor_linha%>" style="color:<%=cor_funcionario%>;">
						<%variaveis1 = caminho&"empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro&cod="&strcod&"','cadastro','width=790,height=350,scrollbars=1,resizable=1"
						'variaveis2 = "empresa/funcionario/funcionarios_folhapagamento_var_cad.asp?tipo=cadastro&cod="&strcod&"&dt_dia="&dt_dia&"&dt_mes="&dt_mes&"&dt_ano="&dt_ano&"&var=1&cat="&strcat&"','variaveis','width=535,height=250,scrollbars=1"
						variaveis2 = caminho&"empresa/funcionario/funcionarios_folhapagamento_rcm_cad.asp?tipo=ver&cod="&strcod&"&dt_dia="&dt_dia&"&dt_mes="&mes_competencia&"&dt_ano="&ano_competencia&"&dia_sel=&mes_sel=&ano_sel=&var=1&cat=enfermagem','variaveisrcm','width=545,height=400,scrollbars=1,resizable=1"
						%>
													
								<td align="right"><b><%=zero(num_funcionario)%></b>&nbsp;</td>
								<td align="right" onClick="window.open('<%=variaveis1%>')"><%=cd_matricula%>&nbsp;</td>
								<td align="left" onClick="window.open('<%=variaveis1%>')"><%=nm_nome%> <%if cd_vt_canc > 0 then response.write("*")%> &nbsp;<%if expediente = 6 then%>**<%end if%><%'if user = 46 then%><%'=data_final%><%'=nr_falta_justif%><%'end if%></td>
								<td align="center" onClick="window.open('<%=variaveis1%>')"><%if tipo_expediente <> "" Then response.write(left(tipo_expediente,2))%></td>
								<td align="center" onClick="window.open('<%=variaveis1%>')"><%=nm_sigla%></td>
								
					<!-- *** Qtd Plantões *** -->
						<%if strtipo = "variaveis" OR strtipo = "var_enf" Then%>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if qtd_plantoes >= 1 then%><%=(qtd_plantoes-total_falta_injust)%><%else%>0<%end if%><%'=cd_unidade%></td>								
									
					<!-- *** Qtd. Adicional Noturno *** -->
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if cd_periodo = 2 then%><b><%=int(nr_ad_noturno)%></b><%else%>0<%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_ad_noturno_he > 0 then%><b><%=int(total_ad_noturno_he)%></b><%else%>0<%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_ad_noturno > 0 then%><b><%=int(total_ad_noturno)%></b><%else%>0<%end if%></td>
								
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Qtd Hora Extra *** -->
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_he_limit > 0 then%><%=(total_he_limit)%><%else%>0<%end if%></td>
					<!-- *** Qtd Hora Extra Excedentes*** -->
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_he_plus > 0 then%><span style="color:<%if int(total_he_plus) > 0 then%>black;<%else%>black;<%end if%>"><%=zero(total_he_plus)%><%else%>0</span><%end if%></td>
					<!-- *** Total Horas Extras *** -->
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_he > 0 then%><b><%=(total_he)%></b><%else%>0<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Folga Enfermagem *** -->
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_folga_enf > 0 then%><%=(total_folga_enf)%><%else%>0<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Faltas justificadas *** -->
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%'if total_falta_justif > 0 then%><%'=(total_falta_justif)%><%'else%><%'end if%>  <%=soma_faltas_justif%></td>
					<!-- *** Faltas Não justificadas *** -->
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%'if total_falta_injust > 0 then%><%'=(total_falta_injust)%><%'else%><%'end if%>  <%=soma_faltas_injust%></td>
					<!-- *** Total Faltas  *** -->
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%'if total_faltas > 0 then%><b><%'=(total_faltas)%></b><%'else%><%'end if%> <b> <%=soma_faltas_geral%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Atrasos *** -->
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_atraso > 0 then%><b><%=(total_atraso)%></b><%else%>0<%end if%></td>
									
						<%end if%>
						<%if strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vr" Then%>
					<!-- *** Benefícios *** -->
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if nr_vr >= 1 then%><%=(nr_vr)&indicativo_folga_enf%><%else%><%if expediente = 6 then%>**<%else%>0<%end if%><%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_vr_extra >= 1 then%><%=(total_vr_extra)%><%else%>0<%end if%></td>
								<td align="center"><%if total_vr >= 1 then%><b><%=(total_vr)%></b><%else%>0<%end if%></td>
						<%end if%>
						<%'if strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vr" Then
						if strtipo = "relat_vr" Then%>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_vr > 0 then%><%=valor_vr%><%else%>&nbsp;<%end if%></td>
								<td align="right" onClick="window.open('<%=variaveis2%>')"><%if total_vr > 0 then%><%=formatnumber(total_custo_vr,2)%><%else%>&nbsp;<%if expediente = 6 then%>**<%else%>0<%end if%><%end if%></td>
						<%end if%>
						<%if strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vt" Then%>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if nr_vt >= 1 then%><%=(nr_vt)%><%else%><%if cd_vt_canc > 0 then%>*<%else%>0<%end if%><%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_vt_extra >= 1 then%><%=(total_vt_extra)%><%else%>0<%end if%>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if total_vt >= 1 then%><b><%=(total_vt)%></b><%else%>0<%end if%></td>
						<%end if%>
						<%if strtipo = "relat_vt" Then%>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if nr_valor_sptrans >= 1 then%><%=formatnumber(nr_valor_sptrans,2)%><%else%>&nbsp;<%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if nr_valor_cmt >= 1 then%><%=formatnumber(nr_valor_cmt,2)%><%else%>&nbsp;<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if nr_valor_sptrans_vt >= 1 then%><%=formatnumber(nr_valor_sptrans_vt,2)%><%else%>&nbsp;<%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if nr_valor_cmt_vt >= 1 then%><%=formatnumber(nr_valor_cmt_vt,2)%><%else%>&nbsp;<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" onClick="window.open('<%=variaveis2%>')"><%if nr_valor_total_valetransporte >= 1 then%><%=formatnumber(nr_valor_total_valetransporte,2)%><%else%>&nbsp;<%end if%></td>
						<%end if%>
						</tr>
					
						<%num_funcionario = num_funcionario + 1
						
					'ELSE
						'num_funcionario = num_funcionario + 1
					END IF
						
						'num_funcionario = num_funcionario + 1
						'linha = linha + 1
							num_cor =  num_cor + 1
								if num_cor > 2 then
									num_cor = 1
								end if
						end if%>
						<%if linha = linhas_limite then%>
							<tr class="ok_print"><td colspan="28" align="right" style="page-break-after:always;">Página <%=pagina%></td></tr>
							<%cab = 1
							if cab = 1 then%>
								<tr bgcolor="#000000" class="ok_print">
									<td colspan="100%" align="center" class="cabecalho" style="background-color:black; color:white;"><b>Horas Extras / Ad. Noturno / Faltas e atrasos - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
									<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
								</tr>
								<tr style="background-color:gray; color:white;" class="ok_print">
									<td>&nbsp;</td>
									<td align="center"><b>Mat.</b></td>
									<td align="center"><b>Funcion&aacute;rio</b></td>
									<td align="center"><b>Hora</b></td>
									<td align="center"><b>Unid.</b></td>
									<td align="center"><b>Plantões</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Ad. Not</b></td>
									<td align="center"><b>Ad. Not<br>H.E.</b></td>
									<td align="center"><b>Total<br>Ad. Not</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Qtd H.E.</b></td>
									<td align="center"><b>H.E.+</b></td>
									<td align="center"><b>Total<br>H.E.</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Folga Enf.</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Faltas just.</b></td>
									<td align="center"><b>Faltas injust.</b></td>
									<td align="center"><b>Total faltas</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Atrasos<br>(horas)</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
										<td align="center"><b>VR</b></td>
										<td align="center"><b>VR Extra</b></td>
										<td align="center"><b>Total VR</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>VT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
										<td align="center"><b>VT Extra</b></td>
										<td align="center"><b>Total VT</b></td>
								</tr>
							<%cab = cab +1
							linha = 0
							end if%>
						<%pagina = pagina + 1
						end if%>
								<%'if qtd_dia_noturno <> "" Then tt_qtd_dia_noturno = abs(tt_qtd_dia_noturno) + qtd_dia_noturno
								
								'linha = linha + 1
								func_1d = "0"
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
								
								hr_entrada = "0"
								hr_saida = "0"
								cd_periodo = "0"
								nr_ad_noturno = "0"
								
								total_ad_noturno = "0"
								total_ad_noturno_he = "0"
								total_func_ad_noturno = "0"
								
								total_he_limit = "0"
								total_he_plus = "0"
								total_he = "0"
								total_folga_enf = "0"
								total_falta_justif = "0"
								total_falta_injust = "0"
								total_faltas = "0"
								total_atraso = "0"
								nr_vr = "0"
								total_vr_extra = "0"
								total_vt_extra = "0"
								
								
								
								
								cd_vtransp_cancela = "0"
								nr_vtransp_cancela = "0"
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
								
								nr_vt = "0"
								
								total_vr = "0"
								total_vt = "0"
								
								qtd_total_faltas = "0"
								cd_vt_canc = 0
								
								cond_1 = "0"
								cond_2 = "0"
								
								indicativo_folga_enf = ""
								
								nr_valor_sptrans = 0
								nr_valor_cmt = 0
								
								nr_valor_total_valetransporte = 0
								
								rs.movenext
								wend%>
							<tr><td colspan="30" bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td></tr>
							<tr>
								<td align="right" colspan="3">&nbsp;</td>
								<td align="center" colspan="3">&nbsp;<b>TOTAL ATIVOS</b></td>
							<%if strtipo = "variaveis" OR strtipo = "var_enf" Then%>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%if total_nr_ad_noturno > 0 then%><%=total_nr_ad_noturno%><%else%>0<%end if%></b></td>
								<td align="center">&nbsp;<b><%if total_geral_ad_noturno_he > 0 then%><%=total_geral_ad_noturno_he%><%else%>0<%end if%></b></td>
								<td align="center">&nbsp;<b><%if total_geral_ad_noturno > 0 then%><%=total_geral_ad_noturno%><%else%>0<%end if%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_he_limit%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_he_plus%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_he%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_folga_enf%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_falta_justif%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_falta_injust%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_faltas%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
								<td align="center"><b><%=total_geral_atraso%></b></td><%'end if%>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
							<%end if%>
							<%if strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vr" Then%>
									<td align="center">&nbsp;<b><%=int(total_geral_nr_vr)%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vr_extra%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vr%></b></td>
							<%end if%>
							<%'if strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vr" Then
							if strtipo = "relat_vr" Then%>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;</td>
								<td align="center">&nbsp;<b><%=formatnumber(total_geral_custo_vr,2)%></b></td>
							<%end if%>
							<%if strtipo = "variaveis" OR strtipo = "var_enf" Then%>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
							<%end if%>
							<%if strtipo = "variaveis" OR strtipo = "var_enf" OR strtipo = "relat_vt" Then%>	<td align="center">&nbsp;<b><%=total_geral_nr_vt%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vt_extra%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vt%></b></td>
							<%end if%>
							<%if strtipo = "relat_vt" Then%>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b> - </b></td>
								<td align="center"><b> - </b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=formatnumber(total_valor_total_sptrans_vt,2)%></b></td>
								<td align="center">&nbsp;<b><%=formatnumber(total_valor_total_cmt_vt,2)%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=formatnumber(total_valor_total_valetransporte,2)%></b></td>
							<%end if%>
							</tr>
							<tr><td colspan="30">&nbsp;</td></tr>
							<tr><td colspan="30" align="center">
								<span style="font-size:9px;" class="no_print">* optou por não receber vale transporte &nbsp; / &nbsp; ** expediente de 6h não recebe vale refeição <%=aviso_folga_enf%><br><img src="../../imagens/px.gif" alt="" width="430" height="1" border="0"></span>
								<span style="font-size:9px;" class="ok_print">* optou por não receber vale transporte &nbsp; / &nbsp; ** expediente de 6h não recebe vale refeição <%=aviso_folga_enf%><br><img src="../../imagens/px.gif" alt="" width="230" height="1" border="0">impr.:<%=now%><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"> Página <%=pagina%></span>
								</td>
							</tr>
							<tr><td></td></tr>
					</table>
	<%end if
	if mostrar_contrato = 0 AND strtipo = "variaveis"  OR mostrar_contrato = 2 AND strtipo = "variaveis"  Then
	'OR strtipo = "variaveis" OR strtipo = "var_enf" %>				
	<br class="no_print">
					
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" style="border-collapse:collapse;">
					<!--br class="no_print"><br class="no_print"-->
			<!-- ---------------------------------------------------------------- -->				
							<tr>
								<td colspan="30" align="center" class="cabecalho" style="background-color:Silver; color:white;"><b>Inativos : Horas Extras / Ad. Noturno / Faltas e atrasos - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
							</tr>
							<tr style="background-color:gray; color:white;">
								<td>&nbsp;<img src="../../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
								<td align="center"><b>Mat.</b><br><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
								<td align="center"><b>Funcion&aacute;rio</b><br><img src="../../imagens/px.gif" alt="" width="205" height="1" border="0"></td>
								<td align="center"><b>Hora</b><br><img src="../../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
								<td align="center"><b>Unid.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>Dias a<br>Trabalhar</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Ad. Not<br>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Total<br>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Qtd H.E.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td align="center"><b>H.E.+</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td align="center"><b>Total<br>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Folga Enfer</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Faltas just.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td align="center"><b>Faltas injust.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>Total faltas</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Atrasos<br>(horas)</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>VR</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>VR Extra</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>Total VR</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>VT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>VT Extra</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>Total VT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
							</tr>
							<%
							'****************************
							'*** Funcionarios +1 dia  ***
							strsql_desl = "SELECT * FROM TBL_funcionario_plantoes_1d where dt_data = '"&dt_mes&"/1/"&dt_ano&"'"
							Set rs_1d_desl = dbconn.execute(strsql_desl)
								if not rs_1d_desl.EOF then nm_1d_desl = rs_1d_desl("nm_funcionarios")
								'response.write(nm_1d)

					num_cor_desl = 1
					num_funcionario_desl = 1
					linhas_limite_desl = 40
					linha_desl = 0
					pagina_desl = 1
					
					ano_competencia_i = ano_competencia
					mes_competencia_i = mes_competencia
					dia_competencia_i = 1
					
					ano_competencia_f = ano_competencia
					mes_competencia_f = mes_competencia
					final_competencia_f = final_competencia
					
					dt_ano = dt_ano
					dt_mes = dt_mes
					dia_final = dia_final
					
					
					'outras_cond = " AND cd_contrato = 1"
				'xsql = "up_funcionario_contrato_lista4 @dt_data_i='"&mes_competencia_i&"/25/"&ano_competencia_i&"', @dt_data_f='"&mes_competencia_f&"/"&final_competencia_f&"/"&ano_competencia_f&"', @dt_atualizacao = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @ordem_funcionarios='"&nm_ordem&"', @outras_variaveis='"&outras_cond&"'"
					
					xsql = "up_funcionario_contrato_lista4_inativos @dt_data_i='"&mes_competencia_i&"/"&dia_competencia_i&"/"&ano_competencia_i&"', @dt_data_f='"&mes_competencia_f&"/"&final_competencia_f&"/"&ano_competencia_f&"', @dt_atualizacao = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @ordem_funcionarios='"&nm_ordem&"', @outras_variaveis='"&outras_cond&"'"
					Set rs = dbconn.execute(xsql)
							while not rs.EOF
							
								strcod_desl = rs("cd_codigo")
								cd_matricula_desl = rs("cd_matricula")
								nm_nome_desl = rs("nm_nome")
								cd_unidade_desl = rs("cd_unidade")
								nm_unidade_desl = rs("nm_unidade")
								nm_sigla_desl = rs("nm_sigla")
								cd_hospital_desl = rs("cd_hospital")
								cd_sexo_desl = rs("cd_sexo")
								
								dt_admissao_desl = rs("dt_admissao_geral")
								dt_demissao_desl = rs("dt_demissao_geral")
								
								'admissao = rs("admissao")
								'demissao = rs("demissao")
									if dt_demissao_desl <> "null" then
										cor_registro_desl = "red"
									else
										cor_registro_desl = "black"
										demissao_desl = ""
									end if
								
								'***** VT Cancelado **********************************************
								'xsql = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod&" AND (cd_tipo = 4) AND (dt_atualizacao <= '04/26/2012') AND (nr_valor > 0)"
								xsql_desl = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod_desl&" AND (cd_tipo = 4) AND (dt_atualizacao <= '"&dt_mes&"/"&dt_dia&"/"&dt_ano&"') AND (nr_valor > 0)"
								Set rs_vt_canc_desl = dbconn.execute(xsql_desl)
									if not rs_vt_canc_desl.EOF Then
										cd_vt_canc_desl = rs_vt_canc_desl("nr_valor")
									end if
								
								if admissao_desl = year(now)&zero(month(now)) then
								'	dias_trab_admissao = datediff(m,dt_admissao,now)
								end if
									
									'*** Caso a admissão for no mesmo mês de competencia da folha ***
									if admissao_desl = dt_ano&zero(dt_mes)  AND demissao_desl = "" OR admissao_desl = dt_ano&zero(dt_mes) AND demissao_desl > dt_ano&zero(dt_mes) then
										tempo_trabalhado_desl = DateDiff("d",day(dt_admissao_desl)&"/"&month(dt_admissao_desl)&"/"&year(dt_admissao_desl),dia_final&"/"&dt_mes&"/"&dt_ano)+1 '(1 = conta o dia da contratação)
									
									'*** Caso a admissão e a demissão forem no mesmo mês de competencia da folha ***
									elseif admissao_desl = dt_ano&zero(dt_mes)  AND demissao_desl = dt_ano&zero(dt_mes) then
										tempo_trabalhado_desl = DateDiff("d",day(dt_admissao_desl)&"/"&month(dt_admissao_desl)&"/"&year(dt_admissao_desl),dia_final&"/"&dt_mes&"/"&dt_ano)+1 '(1 = conta o dia da contratação)
									
									else
										tempo_trabalhado_desl = "30"
									
									'*** Caso a demissão for no mesmo mês de competencia da folha ***
									'elseif  admissao < dt_ano&zero(dt_mes)  AND demissao = dt_ano&zero(dt_mes) then 'OR admissao < dt_ano&zero(dt_mes) AND demissao > dt_ano&zero(dt_mes) then
									'elseif  admissao < dt_ano&zero(dt_mes)  AND month(dt_demissao)&"/"&day(dt_demissao)&"/"&year(dt_demissao) <= zero(dt_mes)&"/25/"&dt_ano then 'OR admissao < dt_ano&zero(dt_mes) AND demissao > dt_ano&zero(dt_mes) then
									'	tempo_trabalhado = DateDiff("d","1/"&mes_pagamento&"/"&ano_pagamento,day(dt_demissao)&"/"&month(dt_demissao)&"/"&year(dt_demissao))
									
									'else
									'	tempo_trabalhado = "0"
									end if
								
								
									
								'*** carrega as informações dos plantões ***
								'strsql = "SELECT * FROM TBL_funcionario_plantoes WHERE dt_atualizacao='"&dt_mes&"/1/"&dt_ano&"'"
								strsql_desl = "SELECT * FROM TBL_funcionario_plantoes WHERE dt_atualizacao='"&dt_mes&"/1/"&dt_ano&"'"
								Set rs_plantoes_desl = dbconn.execute(strsql_desl)
									if not rs_plantoes_desl.EOF then
										plantoes_6h_desl = rs_plantoes_desl("nr_plantoes_6h")
										plantoes_8h_desl = rs_plantoes_desl("nr_plantoes_8h")
										plantoes_12h_desl = rs_plantoes_desl("nr_plantoes_12h")
									end if
								
								expediente_desl = rs("expediente")
									expediente_desl = expediente_desl / 60
										if expediente_desl >= 9 AND expediente_desl <= 11 then
											tipo_expediente_desl = "8h"
											carga_mensal_desl = 220
											qtd_plantoes_desl = plantoes_8h_desl
											cor_fincionario_desl = "black"
										elseif expediente_desl = 6 then
											tipo_expediente_desl = "6h"
											carga_mensal_desl = 180
											qtd_plantoes_desl = plantoes_6h_desl
											cor_fincionario_desl = "black"
										elseif expediente_desl = 12 then
											tipo_expediente_desl = "12x36h"
											carga_mensal_desl = 180
											qtd_plantoes_desl = plantoes_12h_desl
											cor_fincionario_desl = "black"
												
												'*** acerto +1 dia ***
												func_1d_desl = instr(1,nm_1d,strcod_desl,1)
												if func_1d_desl > 0 then qtd_plantoes_desl = qtd_plantoes_desl + 1
										else
											tipo_expediente_desl = "x"
											carga_mensal_desl = 0
											qtd_plantoes_desl = 0
											cor_fincionario_desl = "red"
										end if
								
								hr_entrada_desl = rs("hr_entrada")
								hr_saida_desl = rs("hr_saida")
								
								
									periodo_desl = hour(hr_entrada_desl) - hour(hr_saida_desl)
										if periodo_desl < 0 then
											nm_periodo_desl = "diurno"
											cd_periodo_desl = 1
											nr_ad_noturno_desl = 0
										elseif periodo_desl > 0 then
											nm_periodo_desl = "noturno"
											cd_periodo_desl = 2
											nr_ad_noturno_desl = qtd_plantoes_desl
										end if
											total_geral_nr_ad_noturno_desl = total_geral_nr_ad_noturno_desl + nr_ad_noturno_desl
							
							
							
							
							
							'if cond_1 = 1 AND cond_2 = 1 then 
							
								'*** Carrega as variáveis do RCM ***
								'xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
								xsql_desl = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod_desl&"', @dt_i='"&mes_ant&"/21/"&ano_ant&"', @dt_f='"&mes_competencia&"/20/"&ano_competencia&"'"
								Set rs_rcm_desl = dbconn.execute(xsql_desl)
									if not rs_rcm_desl.EOF then
										total_ad_noturno_he_desl = rs_rcm_desl("total_ad_noturno")
										'total_func_ad_noturno = nr_ad_noturno + total_ad_noturno
											
										total_he_desl = rs_rcm_desl("total_he")/60
										'********************************************
										'*** Arredondamento Hora Extra fracionada ***
												if instr(1,total_he_desl,",",1) > 0 then 
													virgula_he_desl = instr(1,total_he_desl,",",1)
													fracao_he_desl = mid(total_he_desl,virgula_he_desl+1,len(total_he_desl))
														if fracao_he_desl > 2 then
															total_he_desl = int(total_he_desl) + 1
															sinal_arredonda_desl = "*"
														else
															total_he_desl = total_he_desl
															sinal_arredonda_desl = ""
														end if
												end if
											
										total_folga_enf_desl = rs_rcm_desl("total_folga_enf")
											
                                        total_falta_justif_desl = rs_rcm_desl("total_falta_justif")
											
										total_falta_injust_desl = rs_rcm_desl("total_falta_injust")
											
										total_faltas_desl = rs_rcm_desl("total_faltas")
											
										total_atraso_desl = rs_rcm_desl("total_atraso")
											
										total_vr_extra_desl = rs_rcm_desl("total_vr_extra")
											
										total_vt_extra_desl = rs_rcm_desl("total_vt_extra")
											
									end if
										'*** Separa os valores de H.E. ***
										if total_he_desl > 36 then
											total_he_plus_desl = total_he_desl - 36
											total_he_limit_desl = 36
										else
											total_he_limit_desl = total_he_desl
											total_he_plus_desl = 0
										end if
										
										
										'*** VT Extra: Ida e Volta
										if total_vt_extra_desl <> "" then
											total_vt_extra_desl = total_vt_extra_desl' *2  //Duplica_VT
										end if
								
								
								'cd_vrefeic_cancela = rs("cd_vrefeic_cancela")
								'nr_vrefeic_cancela = rs("nr_vrefeic_cancela")
								'	if  not IsNumeric(nr_vrefeic_cancela) Then nr_vrefeic_cancela = "0"
								
								'cd_vtransp_cancela = rs("cd_vtransp_cancela")
								'nr_vtransp_cancela = rs("nr_vtransp_cancela")
								'	if  not IsNumeric(nr_vtransp_cancela) Then nr_vtransp_cancela = "0"
								
								'**** Calcula a qtd de VR e VT ************
								if expediente_desl <> 6 then
									nr_vr_desl = qtd_plantoes_desl - total_faltas_desl
                                   ' nr_vr = nr_vr - total_folga_enf -20
								end if
								
								if qtd_plantoes_desl <> "0" AND cd_vt_canc_desl < 1 then
									nr_vt_desl = (qtd_plantoes_desl) '* 2)' - qtd_total_faltas // Duplica_VT
									'nr_vt_ = nr_vt - total_folga_enf
                                    qtd_total_faltas_vt_desl = total_faltas_desl' * 2 // Duplica_VT
									nr_vt_desl = nr_vt_desl - qtd_total_faltas_vt_desl
                                    
								end if
								
									if num_cor_desl = 1 then
										cor_linha_desl = "#ccffff"
									else
										cor_linha_desl = "#ffffff"
									end if
									
									
							'***************
							'*** Filtros ***
							'***************
									if int(strcd_unidade_desl) = "0" OR int(cd_unidade_desl) = int(strcd_unidade_desl) then
										cond_1_desl = 1
									end if
									
									if mostra_faltas_desl <> 0 then 
										if int(mostra_faltas_desl) = "1" AND  int(total_faltas_desl) = 0 then
											cond_2_desl = 1
										elseif int(mostra_faltas_desl) = "2" AND int(total_faltas_desl) > 0 then
											cond_2_desl = 1
										end if
									else
										cond_2_desl = 1
									end if%>
							
							
					<%'if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
					if cond_1_desl = 1 AND cond_2_desl = 1 then 
					'if int(filtros)="00" AND  int(total_faltas)>=0 AND int(cd_unidade) = int(strcd_unidade) OR int(filtros)="00" AND int(total_faltas)<=0 AND int(cd_unidade) = int(strcd_unidade) OR int(filtros)="00" AND  int(total_faltas)>=1 AND int(cd_unidade) = int(strcd_unidade)then
					
					
							
							if total_vr_extra_desl <> "" Then
								total_vr_desl = int(nr_vr_desl) + int(total_vr_extra_desl)
							else
								total_vr_desl = int(nr_vr_desl)
							end if
							
							if total_vt_extra_desl <> "" Then
								total_vt_desl = int(nr_vt_desl) + int(total_vt_extra_desl)
							else
								total_vt_desl = int(nr_vt_desl)
							end if
								
							
							'*** Tota geral variaveis ***
							
							
									
							if total_he_plus_desl <> "" Then total_geral_he_plus_desl = total_geral_he_plus_desl + total_he_plus_desl
							if total_he_limit_desl <> "" Then total_geral_he_limit_desl = total_geral_he_limit_desl + total_he_limit_desl
							if total_he_desl <> "" Then total_geral_he_desl = total_geral_he_desl + total_he_desl
							
							if total_folga_enf_desl <> "" Then total_geral_folga_enf_desl = total_geral_folga_enf_desl + total_folga_enf_desl
							
                            if total_falta_justif_desl <> "" Then total_geral_falta_justif_desl = total_geral_falta_justif_desl + total_falta_justif_desl
							if total_falta_injust_desl <> "" Then total_geral_falta_injust_desl = total_geral_falta_injust_desl + total_falta_injust_desl
							if total_faltas_desl <> "" Then total_geral_faltas_desl = total_geral_faltas_desl + total_faltas_desl
							if total_atraso_desl <> "" Then total_geral_atraso_desl = total_geral_atraso_desl + total_atraso_desl
							
							total_geral_ad_noturno_he_desl = total_geral_ad_noturno_he_desl + total_ad_noturno_he_desl
							'total_geral_ad_noturno = total_geral_ad_noturno_he + total_nr_ad_noturno
								if cd_periodo_desl = 2 then 
									total_nr_ad_noturno_desl = total_nr_ad_noturno_desl + nr_ad_noturno_desl
									total_ad_noturno_desl = total_ad_noturno_he_desl + nr_ad_noturno_desl - total_faltas_desl
										'total_geral_ad_noturno =  total_geral_ad_noturno + nr_ad_noturno + total_ad_noturno_he
										total_geral_ad_noturno_desl =  total_geral_ad_noturno_desl + nr_ad_noturno_desl + total_ad_noturno_he_desl
								else 
									total_ad_noturno_desl = total_ad_noturno_he_desl' + nr_ad_noturno
									'total_geral_ad_noturno =  total_nr_ad_noturno + nr_ad_noturno
										total_geral_ad_noturno_desl =  total_geral_ad_noturno_desl + total_ad_noturno_he_desl
								end if
									
									
							'if total_vr <> "" Then total_geral_vr = total_geral_vr + total_vr
							'total_geral_vt = total_geral_vt + total_vt
							
							if nr_vr_desl <> "" then total_geral_nr_vr_desl = total_geral_nr_vr_desl + nr_vr_desl
							if total_vr_extra_desl <> "" Then total_geral_vr_extra_desl = total_geral_vr_extra_desl + total_vr_extra_desl
							total_geral_vr_desl = total_geral_vr_desl + nr_vr_desl + total_vr_extra_desl
							
							if nr_vt_desl <> "" then total_geral_nr_vt_desl = total_geral_nr_vt_desl + nr_vt_desl
							if total_vt_extra_desl <> "" Then total_geral_vt_extra_desl = total_geral_vt_extra_desl + total_vt_extra_desl
							total_geral_vt_desl = total_geral_vt_desl + nr_vt_desl + total_vt_extra_desl
							%>
						<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'white<%'=cor_linha%>');" bgcolor="white<%'=cor_linha%>" style="color:<%=cor_funcionario%>;">	
						<%variaveis1_desl = caminho_desl&"empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro&cod="&strcod_desl&"','cadastro','width=790,height=350,scrollbars=1,resizable=1"
						'variaveis2 = "empresa/funcionario/funcionarios_folhapagamento_var_cad.asp?tipo=cadastro&cod="&strcod&"&dt_dia="&dt_dia&"&dt_mes="&dt_mes&"&dt_ano="&dt_ano&"&var=1&cat="&strcat&"','variaveis','width=535,height=250,scrollbars=1"
						variaveis2_desl = caminho_desl&"empresa/funcionario/funcionarios_folhapagamento_rcm_cad.asp?tipo=ver&cod="&strcod_desl&"&dt_dia="&dt_dia&"&dt_mes="&mes_competencia&"&dt_ano="&ano_competencia&"&dia_sel=&mes_sel=&ano_sel=&var=1&cat=enfermagem','variaveisrcm','width=545,height=400,scrollbars=1,resizable=1"
						
								data_fechamento = "25/"&mes_competencia&"/"&ano_competencia
								str_fechamento = "25/"&mes_competencia&"/"&right(ano_competencia,2)
								dias_trabalhados = datediff("d", dt_demissao_desl,data_fechamento)%>
													
								<td align="right"><b><%=zero(num_funcionario_desl)%></b>&nbsp;<%'=" - "&periodo_ferias%></td>
								<td align="right" onClick="window.open('<%=variaveis1_desl%>')"><%=cd_matricula_desl%>&nbsp; <%'=hr_entrada&"-"&hr_saida%></td>
								<td align="left" onClick="window.open('<%=variaveis1_desl%>')"><%'=strcod%><%=nm_nome_desl%> <%if cd_vt_canc_desl > 0 then response.write("*")%> &nbsp;<%if expediente_desl = 6 then%>**<%end if%><%=str_fechamento&" "&dias_trabalhados%><%'=hour(hr_entrada)&":"&minute(hr_entrada)&"/"&hour(hr_saida)&":"&minute(hr_saida)&" - "&tipo_expediente&"-"&cd_periodo%></td>
								<td align="center" onClick="window.open('<%=variaveis1_desl%>')"><%if tipo_expediente_desl <> "" Then response.write(left(tipo_expediente_desl,2))%></td>
								<td align="center" onClick="window.open('<%=variaveis1_desl%>')"><%=nm_sigla_desl%></td>
								
					<!-- *** Qtd Plantões *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if qtd_plantoes_desl >= 1 then%><%=(qtd_plantoes_desl-total_falta_injust_desl)%><%else%>0<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Qtd. Adicional Noturno *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')">&nbsp;<%if cd_periodo_desl = 2 then%><b><%=int(nr_ad_noturno_desl)%></b><%else%>0<%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')">&nbsp;<%if total_ad_noturno_he_desl > 0 then%><b><%=int(total_ad_noturno_he_desl)%></b><%else%>0<%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')">&nbsp;<%if total_ad_noturno_desl > 0 then%><b><%=int(total_ad_noturno_desl)%></b><%else%>0<%end if%></td>
								
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Qtd Hora Extra *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_he_limit_desl > 0 then%><%=(total_he_limit_desl)%><%else%>0<%end if%></td>
					<!-- *** Qtd Hora Extra Excedentes*** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_he_plus_desl > 0 then%><span style="color:<%if int(total_he_plus_desl) > 0 then%>black;<%else%>black;<%end if%>"><%=zero(total_he_plus)%><%else%>0</span><%end if%></td>
					<!-- *** Total Horas Extras *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_he_desl > 0 then%><b><%=(total_he_desl)%></b><%else%>0<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Folga Enfermagem *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_folga_enf_desl > 0 then%><%=(total_folga_enf_desl)%><%else%>0<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Faltas justificadas *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_falta_justif_desl > 0 then%><%=(total_falta_justif_desl)%><%else%>0<%end if%></td>
					<!-- *** Faltas Não justificadas *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_falta_injust_desl > 0 then%><%=(total_falta_injust_desl)%><%else%>0<%end if%></td>
					<!-- *** Total Faltas  *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_faltas_desl > 0 then%><b><%=(total_faltas_desl)%></b><%else%>0<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Atrasos *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_atraso_desl > 0 then%><b><%=(total_atraso_desl)%></b><%else%>0<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Benefícios *** -->
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if nr_vr_desl >= 1 then%><%=(nr_vr_desl)%><%else%><%if expediente_desl = 6 then response.write("**")%><%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_vr_extra_desl >= 1 then%><%=(total_vr_extra_desl)%><%else%>0<%end if%></td>
								<td align="center"><%if total_vr_desl >= 1 then%><b><%=(total_vr_desl)%></b><%else%>0<%end if%></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if nr_vt_desl >= 1 then%><%=(nr_vt_desl)%><%else%><%if cd_vt_canc_desl > 0 then response.write("*")%><%end if%></td>
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_vt_extra_desl >= 1 then%><%=(total_vt_extra_desl)%><%else%>0<%end if%>
								<td align="center" onClick="window.open('<%=variaveis2_desl%>')"><%if total_vt_desl >= 1 then%><b><%=(total_vt_desl)%></b><%else%>0<%end if%></td>
						</tr>
						<%num_funcionario_desl = num_funcionario_desl + 1
						linha_desl = linha_desl + 1
							num_cor_desl =  num_cor_desl + 1
								if num_cor_desl > 2 then
									num_cor_desl = 1
								end if
						end if%>
						<%if linha_desl = linhas_limite_desl then%>
							<tr class="ok_print"><td colspan="28" align="right" style="page-break-after:always;">Página <%=pagina_desl%></td></tr>
							<%cab_desl = 1
							if cab_desl = 1 then%>
								<tr bgcolor="#000000" class="ok_print">
									<td colspan="100%" align="center" class="cabecalho" style="background-color:black; color:white;"><b>Horas Extras / Ad. Noturno / Faltas e atrasos - <%=mesdoano(dt_mes)%>/<%=dt_ano%></b></td>
									<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
								</tr>
								<tr style="background-color:gray; color:white;" class="ok_print">
									<td>&nbsp;</td>
									<td align="center"><b>Mat.</b></td>
									<td align="center"><b>Funcion&aacute;rio</b></td>
									<td align="center"><b>Hora</b></td>
									<td align="center"><b>Unid.</b></td>
									<td align="center"><b>Plantões</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Ad. Not</b></td>
									<td align="center"><b>Ad. Not<br>H.E.</b></td>
									<td align="center"><b>Total<br>Ad. Not</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Qtd H.E.</b></td>
									<td align="center"><b>H.E.+</b></td>
									<td align="center"><b>Total<br>H.E.</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Folga Enf.</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Faltas just.</b></td>
									<td align="center"><b>Faltas injust.</b></td>
									<td align="center"><b>Total faltas</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>Atrasos<br>(horas)</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
										<td align="center"><b>VR</b></td>
										<td align="center"><b>VR Extra</b></td>
										<td align="center"><b>Total VR</b></td>
										<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>VT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
										<td align="center"><b>VT Extra</b></td>
										<td align="center"><b>Total VT</b></td>
								</tr>
							<%cab_desl = cab_desl +1
							linha_desl = 0
							end if%>
						<%pagina_desl = pagina_desl + 1
						end if%>
								<%'if qtd_dia_noturno <> "" Then tt_qtd_dia_noturno = abs(tt_qtd_dia_noturno) + qtd_dia_noturno
								
								'linha = linha + 1
								func_1d_desl = "0"
								strcod_desl = "0"
								cd_matricula_desl = "0"
								nm_nome_desl = "0"
								tempo_trabalhado_desl = ""
								cd_unidade_desl = "0"
								nm_unidade_desl = "0"
								nm_sigla_desl = "0"
								cd_hospital_desl = "0"
								cd_salario_desl = "0"
								nr_salario_desl = "0"
								nr_salario_hora_desl = "0"
								nr_salario_dia_desl = "0"
								nr_salario_desc_desl = "0"
								nr_desc_salario_desl = "0"
								insalubridade_desl = "0"
								nr_insalubridade_forca_desl = "0"
								total_faltas_mes_desl = "0"
								expediente_desl = "0"
								tipo_expediente_desl = "0"
								carga_mensal_desl = "0"
								cd_aj_custo_desl = "0"
								nr_aj_custo_desl = "0"
								cd_ad_noturno_desl = "0"
								qtd_dia_noturno_desl = ""
								
								hr_entrada_desl = "0"
								hr_saida_desl = "0"
								cd_periodo_desl = "0"
								nr_ad_noturno_desl = "0"
								
								total_ad_noturno_desl = "0"
								total_ad_noturno_he_desl = "0"
								total_func_ad_noturno_desl = "0"
								
								total_he_limit_desl = "0"
								total_he_plus_desl = "0"
								total_he_desl = "0"
								total_folga_enf_desl = "0"
								total_falta_justif_desl = "0"
								total_falta_injust_desl = "0"
								total_faltas_desl = "0"
								total_atraso_desl = "0"
								nr_vr_desl = "0"
								total_vr_extra_desl = "0"
								total_vt_extra_desl = "0"
								
								
								
								
								cd_vtransp_cancela_desl = "0"
								nr_vtransp_cancela_desl = "0"
								idade_desl = ""
								ok_auxcreche_desl = "0"
								auxcreche_desl = 0
								cesta_b_desl = "0"
								valor_refeicao_desl = "0"
								v_trasnsp_desl = "0"
								base_inss_desl = "0"
								nr_conv_medico_desl = "0"
								nr_contr_sind_desl = "0"
								fgts_desl = "0"
								total_vencimento_desl = "0"
								total_desconto_desl = "0"
								base_irrf_desl = "0"
								periodo_ferias_desl = ""
								nr_desconto_ferias_desl = "0"
								nr_faltas_justif_desl = "0"
								nr_faltas_injust_desl = "0"
								
								
								nr_vt_desl = "0"
								
								total_vr_desl = "0"
								total_vt_desl = "0"
								
								qtd_total_faltas_desl = "0"
								cd_vt_canc_desl = "0"
								
								cond_1_desl = "0"
								cond_2_desl = "0"
								
								rs.movenext
								wend%>
							<tr><td colspan="30" bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td></tr>
							<tr>
								<td align="right" colspan="3">&nbsp;</td>
								<td align="center" colspan="3">&nbsp;<b>TOTAL INATIVOS</b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%if total_nr_ad_noturno_desl > 0 then%><%=total_nr_ad_noturno_desl%><%else%>0<%end if%></b></td>
								<td align="center">&nbsp;<b><%if total_geral_ad_noturno_he_desl > 0 then%><%=total_geral_ad_noturno_he_desl%><%else%>0<%end if%></b></td>
								<td align="center">&nbsp;<b><%if total_geral_ad_noturno_desl > 0 then%><%=total_geral_ad_noturno_desl%><%else%>0<%end if%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_he_limit_desl%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_he_plus_desl%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_he_desl%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_folga_enf_desl%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_falta_justif_desl%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_falta_injust_desl%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_faltas_desl%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
								<td align="center"><b><%=total_geral_atraso_desl%></b></td><%'end if%>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
								<td align="center">&nbsp;<b><%=int(total_geral_nr_vr_desl)%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vr_extra_desl%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vr_desl%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
								<td align="center">&nbsp;<b><%=total_geral_nr_vt_desl%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vt_extra_desl%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vt_desl%></b></td>
							</tr>
							<tr><td colspan="30">&nbsp;</td></tr>
							

							<tr><td colspan="30" align="center">
								<span style="font-size:9px;" class="no_print">* optou por não receber vale transporte &nbsp; / &nbsp; ** expediente de 6h não recebe vale refeição<img src="../../imagens/px.gif" alt="" width="430" height="1" border="0"></span>
								<span style="font-size:9px;" class="ok_print">* optou por não receber vale transporte &nbsp; / &nbsp; ** expediente de 6h não recebe vale refeição<img src="../../imagens/px.gif" alt="" width="230" height="1" border="0">impr.:<%=now%><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"> Página <%=pagina_desl%></span>
								</td>
							</tr>
							<tr><td colspan="30">&nbsp;</td></tr>
							
						</table>
<%end if%>					
							
					<%'if then
					if mostrar_contrato = 0 AND strtipo = "variaveis"  OR mostrar_contrato = 2 AND strtipo = "variaveis" AND cd_user = 46 Then
	
					
					'*** Tota geral variaveis ***
							total_nr_ad_noturno_final = total_nr_ad_noturno + total_nr_ad_noturno_desl
							total_geral_ad_noturno_he_final = total_geral_ad_noturno_he + total_geral_ad_noturno_he_desl
							total_geral_ad_noturno_final = total_geral_ad_noturno + total_geral_ad_noturno_desl
							
							total_geral_he_limit_final = total_geral_he_limit + total_geral_he_limit_desl
							total_geral_he_plus_final = total_geral_he_plus + total_geral_he_plus_desl
							total_geral_he_final = total_geral_he + total_geral_he_desl
							total_geral_folga_enf_final = total_geral_folga_enf + total_geral_folga_enf_desl
							total_geral_falta_justif_final = total_geral_falta_justif + total_geral_falta_justif_desl
							total_geral_falta_injust_final = total_geral_falta_injust + total_geral_falta_injust_desl
							total_geral_faltas_final = total_geral_faltas + total_geral_faltas_desl
							total_geral_atraso_final = total_geral_atraso + total_geral_atraso_desl
							total_geral_nr_vr_final = total_geral_nr_vr + total_geral_nr_vr_desl
							total_geral_vr_extra_final = total_geral_vr_extra_final + total_geral_vr_extra_desl
							total_geral_vr_final = total_geral_vr + total_geral_vr_desl
							total_geral_nr_vt_final = total_geral_nr_vt + total_geral_nr_vt_desl
							total_geral_vt_extra_final = total_geral_vt_extra + total_geral_vt_extra_desl
							total_geral_vt_final = total_geral_vt + total_geral_vt_desl
							%>
						<table border="1" cellspacing="0" cellpadding="2" align="center" width="800" style="border-collapse:collapse;">
							<tr style="background-color:gray; color:white;">
								<td colspan="6"><img src="../../imagens/px.gif" alt="" width="252" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Ad. Not<br>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
								<td align="center"><b>Total<br>Ad. Not</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Qtd H.E.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td align="center"><b>H.E.+</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td align="center"><b>Total<br>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Folga Enfer</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Faltas just.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
								<td align="center"><b>Faltas injust.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>Total faltas</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Atrasos<br>(horas)</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
									<td align="center"><b>VR</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>VR Extra</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>Total VR</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>VT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>VT Extra</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>Total VT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
							</tr>
							<tr>
								<td align="right" colspan="3"><img src="../../imagens/px.gif" alt="" width="252" height="1" border="0"></td>
								<td align="center" colspan="3">&nbsp;<b>TOTAL GERAL</b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%if total_nr_ad_noturno_final > 0 then%><%=total_nr_ad_noturno_final%><%else%>0<%end if%></b></td>
								<td align="center">&nbsp;<b><%if total_geral_ad_noturno_he_final > 0 then%><%=total_geral_ad_noturno_he_final%><%else%>0<%end if%></b></td>
								<td align="center">&nbsp;<b><%if total_geral_ad_noturno_final > 0 then%><%=total_geral_ad_noturno_final%><%else%>0<%end if%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_he_limit_final%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_he_plus_final%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_he_final%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_folga_enf_final%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_falta_justif_final%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_falta_injust_final%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_faltas_final%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
								<td align="center"><b><%=total_geral_atraso_final%></b></td><%'end if%>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
								<td align="center">&nbsp;<b><%=int(total_geral_nr_vr_final)%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vr_extra_final%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vr_final%></b></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>	
								<td align="center">&nbsp;<b><%=total_geral_nr_vt_final%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vt_extra_final%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_vt_final%></b></td>
							</tr>
						</table>	
							<%end if%>
							
							
							
					</table>
