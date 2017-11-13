
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%
cd_user = session("cd_codigo")

strcod = request("cod")

strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")

dt_dia = request("dt_dia")
dt_mes = request("dt_mes")
dt_ano = request("dt_ano")
	if dt_dia = "" Then dt_dia = day(now) end if
	if dt_mes = "" Then dt_mes = month(now) end if
	if dt_ano = "" Then dt_ano = year(now) end if

dia_final = ultimodiames(dt_mes,dt_ano)

mes_pagamento = dt_mes + 1
	if mes_pagamento > 12 then
		mes_pagamento = 1
		ano_pagamento =	dt_ano + 1
	else
		ano_pagamento =	dt_ano
	end if
final_pagamento = ultimodiames(mes_pagamento,ano_pagamento)

	mesatual=year(now)&zero(month(now))
	messel=dt_ano&zero(dt_mes)
		if messel < mesatual then
			competencia = "1"
		elseif messel >= mesatual then
			competencia = "1"
		end if
	
'*** Quantidade de dias da competencia ***
	qtd_dias_comp = datediff("d","1/"&dt_mes&"/"&dt_ano,dia_final&"/"&dt_mes&"/"&dt_ano)+1
	'qtd_dias_comp = datediff("d","1/9/2011","30/9/2011")

'*** Quantidade de domingos da competencia ***
	qtd_domingos_comp = datediff("ww","1/"&dt_mes&"/"&dt_ano,dia_final&"/"&dt_mes&"/"&dt_ano,1)

'response.write("<br>"&dt_mes&"1/"&dt_ano&"<br>"&dia_final&"/"&dt_mes&"/"&dt_ano)

			if ordem_total <> "" Then
				ordem_funcionarios = ordem_total
			else
				'ordem_funcionarios = "cd_contrato,nm_nome,nm_sigla"
				ordem_funcionarios = "cd_contrato,nm_nome"
			end if
			
data_atual = dia_final&"/"&dt_mes&"/"&dt_ano%>
<!--#include file="../../includes/feriados.asp"-->
<%qtd_dias_dsr = qtd_domingos_comp + qtd_feriados - qtd_feriado_dsr%>

<style type="text/css" media="screen">
<!--
.ok_print{ visibility:hidden; display: none;}
.no_print{ visibility:visible; display:block;}
#frame{border:0px inset;
	position:absolute;
	height:93%;
	width:79%;
	top:20px;
	left:210px;
	margin: 0px;
	padding: 6px;
	text-align:left;
	overflow:scroll;
	text-decoration:none;}
table{background-color: #ffffff; 
    border: 1px solid #cccccc;
	font-family: verdana;
	font-size: 9px;
	text-decoration:none;}

a:link {color:#000000;text-decoration:none;}
a:visited {color:#000000;text-decoration:none;}
a:hover {color:#FF0000;}
a:active {color:#000000;}

.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.cabecalho {color: #000000;font-family: verdana;font-size:14px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; } 

-->
</style>

<style type="text/css" media="print">
<!--
.ok_print{ visibility:hidden; display: none;}
.no_print{ visibility:visible; display:block;}
#frame{border:0px inset;
	position:absolute;
	height:93%;
	width:79%;
	top:20px;
	left:210px;
	margin: 0px;
	padding: 6px;
	text-align:left;
	overflow:scroll;
	text-decoration:none;}
table{background-color: #ffffff; 
    border: 1px solid #cccccc;
	font-family: verdana;
	font-size: 9px;
	text-decoration:none;}

a:link {color:#000000;text-decoration:none;}
a:visited {color:#000000;text-decoration:none;}
a:hover {color:#FF0000;}
a:active {color:#000000;}

.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.cabecalho {color: #000000;font-family: verdana;font-size:14px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; } 

-->
</style>
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print" style="border-collapse:collapse;">
							<tr>
								<td colspan="6" align="center" class="cabecalho" style="background-color:black; color:white;"><b>DEMONSTRATIVO DE PAGAMENTO <%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td align="center" style="background-color:silver;" colspan="3">&nbsp;</td>
								<td align="center" style="background-color:silver;" colspan="3"><b><%=mesdoano(dt_mes)&"/"&dt_ano%></b></td>									
							</tr>
						
						
			<!-- ---------------------------------------------------------------- -->				
												
						
							<%'End if%>
							<tr style="color:white;">
								<td align="center" colspan="6">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> DSR.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês &nbsp; &nbsp; &nbsp; &nbsp;(<%=qtd_feriado_dsr%> feriado dom) </td>
							</tr>
							<%'**************************
							'*** VALOR SALÁRIO MÍNIMO ***
							xsql = "SELECT * FROM TBL_funcionario_valores_padroes WHERE (cd_tipo = 12) AND dt_atualizacao <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
							Set rs = dbconn.execute(xsql)
							nr_salminimo = rs("nr_valor")
							'****************************
							'*** VALOR Aux. Creche	  ***
							xsql = "SELECT * FROM TBL_funcionario_valores_padroes WHERE (cd_tipo = 6) AND dt_atualizacao <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
							Set rs = dbconn.execute(xsql)
							nr_auxcreche = rs("nr_valor")
							'******************************

							'xsql = "up_funcionario_rh_lista_individual @cd_funcionario="&strcod&", @dt_atualizacao='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
							xsql = "up_funcionario_folhapagamento_individual @cd_funcionario="&strcod&", @dt_pagamento_i='"&mes_pagamento&"/1/"&ano_pagamento&"', @dt_pagamento_f='"&mes_pagamento&"/"&final_pagamento&"/"&ano_pagamento&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
								Set rs = dbconn.execute(xsql)
								nm_nome = rs("nm_nome")
								cd_matricula = rs("cd_matricula")
								cd_unidade = rs("cd_unidade")
								nm_sigla = rs("nm_sigla")
								nm_unidade = rs("nm_unidade")
								cd_hospital = rs("cd_hospital")
								cd_sexo = rs("cd_sexo")
								
								nm_banco = rs("nm_banco")
								cd_banco_ag = rs("cd_banco_ag")
								cd_banco_conta = rs("cd_banco_conta")
								
								cd_cargo = rs("cd_cargo")
								nm_cargo = rs("nm_cargo")
								
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
									
								'*** Valor Salário em caso de contratação no mesmo mês ***
								if tempo_trabalhado < "30" then
									nr_salario = nr_salario_dia * tempo_trabalhado
								end if
								
								expediente = rs("expediente")
									expediente = expediente / 60
										if expediente = 9 then
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
										
										if nr_salario <> "" and carga_diaria <> "" Then
											nr_salario_hora = nr_salario_dia/carga_diaria
										end if
								
								cd_aj_custo = rs("cd_aj_custo")
								nr_aj_custo = rs("nr_aj_custo")
									if  not IsNumeric(nr_aj_custo) Then nr_aj_custo = "0"
								
								cd_insalubridade_forca = rs("cd_insalubridade_forca")
								nr_insalubridade_forca = rs("nr_insalubridade_forca")
									if  not IsNumeric(nr_insalubridade_forca) Then nr_insalubridade_forca = "0"
								
								cd_ad_noturno = rs("cd_ad_noturno")
								qtd_dia_noturno = rs("qtd_dia_noturno")
								
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
									if  not IsNumeric(nr_desconto_he) Then nr_desconto_he = "0"%>
							
							<tr style="background-color:gray; color: white;">
								<td align="center"><b>Cód</b></td>
								<td align="left" colspan="3"><b>Funcion&aacute;rio</b></td>
								<td align="center"><b>Unidade</b></td>
								<td align="center"><b>Carga Horária</b></td>
							</tr>
							<tr>
								<td align="center" valign="top"><b><%=cd_matricula%></b></td>
								<td align="left" colspan="3" valign="top"><b><%=nm_nome%></b></td>
								<td align="center" valign="top"><b><%=nm_sigla%><%'=nm_unidade%></b></td>
								<td align="center" valign="top"><b><%=tipo_expediente%></b></td>								
							</tr>
							<tr style="background-color:gray; color: white;">
								<td align="center" colspan="3"><b>Descrição</b></td>
								<td align="center"><b>Qtd.</b></td>
								<td align="center"><b>Vencimentos</b></td>
								<td align="center"><b>Descontos</b></td>
							</tr>
			<!-- ------------------------------------------- V E N C I M E N T O S ------------------------------------------- -->
							<%cor_linha = "#ccffff"%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Sal&aacute;rio Bruto</b></td>
								<td align="right">&nbsp;<%if tempo_trabalhado > "0" and tempo_trabalhado < "30" then response.write(tempo_trabalhado&" dias")%></td>
								<td align="right">&nbsp;<%if nr_salario <> "0" then response.write(FormatNumber(nr_salario,2))%></td>
								<td align="right">&nbsp;</td>								
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Ajuda de Custo</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;<%=nr_aj_custo%></td>
								<td align="right">&nbsp;</td>								
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Insalubridade</b></td>
								<td align="right">&nbsp;</td>
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
								<td align="right">&nbsp;</td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Adicional Noturno</b></td>
								<td align="right"><%=zero(qtd_dia_noturno)%> <%if qtd_dia_noturno = 1 then%> dia<%elseif qtd_dia_noturno > 1 then%>dias<%end if%>&nbsp;</td>
								<td align="right"><%if qtd_dia_noturno > 0 then
														ad_noturno = (((nr_salario*0.40)/carga_mensal)*7)*qtd_dia_noturno
														qtd_dias_mes = qtd_dias_comp-qtd_dias_dsr
														dsr_ad_not = ad_noturno/qtd_dias_mes
															dsr_ad_not = formatnumber(dsr_ad_not*qtd_dias_dsr,2)
														
														response.write(formatnumber(ad_noturno,2))
													else
														response.write("0")
													end if%></td>
								<td align="right">&nbsp;</td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>DSR Ad. Not.</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;<%=dsr_ad_not%></td>
								<td>&nbsp;</td>
							</tr>
							<%xsql_dep = ("SELECT COUNT(nm_nome) AS qtd_dependentes FROM View_dependentes WHERE  (DATEDIFF(m, dt_nascimento, '"&dt_mes&"/"&dia_final&"/"&dt_ano&"') <= 251) AND (cd_funcionario = "&strcod&")")
							Set rs_dep = dbconn.execute(xsql_dep)
							if not rs_dep.EOF then
								qtd_dependentes_21 = rs_dep("qtd_dependentes")
							end if%>
							<!--tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>Dependentes < 21anos</b></td>
								<td align="right">&nbsp;<%if qtd_dependentes = 1 then%> <%=zero(qtd_dependentes)%> filho <%elseif qtd_dependentes > 1 then%><%=zero(qtd_dependentes)%> filhos<%end if%></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;</td>
							</tr-->
							<%'xsql_dep = ("SELECT COUNT(nm_nome) AS qtd_dependentes FROM View_dependentes WHERE  (DATEDIFF(m, dt_nascimento, '"&dt_mes&"/"&dia_final&"/"&dt_ano&"') <= 83) AND (cd_funcionario = "&strcod&")")
							'Set rs_dep = dbconn.execute(xsql_dep)
							'if not rs_dep.EOF then
							'	qtd_dependentes = rs_dep("qtd_dependentes")
							'end if%>
							<%'*** Informa os dados o dependente ***
							xsql_dep = ("SELECT * FROM View_dependentes WHERE (DATEDIFF(m, dt_nascimento, '"&dt_mes&"/"&dia_final&"/"&dt_ano&"') <= 83) AND dt_nascimento <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'  AND (cd_funcionario = "&strcod&")")
								Set rs_dep = dbconn.execute(xsql_dep)
								while not rs_dep.EOF
									cd_depend = rs_dep("cd_codigo")
									nm_nome = rs_dep("nm_nome")
									dt_nascimento = rs_dep("dt_nascimento")
										idade = DateDiff("m",day(dt_nascimento)&"/"&month(dt_nascimento)&"/"&year(dt_nascimento),dia_final&"/"&dt_mes&"/"&dt_ano)
										if idade = 79 then
											auxcreche_color = "yellow"
										else
											auxcreche_color = "green"
										end if
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
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Aux&iacute;lio Creche</b></td>
								<td align="right">&nbsp;<%if qtd_dep_idade = 1 then%> <%=zero(qtd_dep_idade)%> filho <%elseif qtd_dep_idade > 1 then%><%=zero(qtd_dep_idade)%> filhos<%end if%></td>
								<td align="right">&nbsp;<%if qtd_dep_auxcreche > 0 AND nr_salario > 0 Then%>
															<%auxcreche = qtd_dep_auxcreche*nr_auxcreche%>
															<%=formatnumber(auxcreche,2)%>
														<%else%>
															0
														<%end if%>
									<%=qtd_mes_dep%></td>
								<td align="right">&nbsp;</td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Hora Extra</b> <%if he_marca = "1" Then response.write("*")%></td>
								<td align="right"><%if qtd_hextra >= 1 then%><%=zero(qtd_hextra)%>h<%end if%>&nbsp;</td>
								<td align="right"><%if qtd_hextra > 0 then
														hora_extra = ((nr_salario/carga_mensal)*1.9)*qtd_hextra
														qtd_dias_mes = qtd_dias_comp-qtd_dias_dsr
														dsr_he = hora_extra/qtd_dias_mes
															dsr_he = formatnumber(dsr_he*qtd_dias_dsr,2)
														
														response.write(formatnumber(hora_extra,2))
													else
														response.write("0")
													end if%></td>
								<td align="right">&nbsp;</td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>DSR H.E.</b><%'=qtd_dias_mes%></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;<%=dsr_he%></td>
								<td align="right">&nbsp;</td>
							</tr>
			<!-- ------------------------------------------- D E S C O N T O S ------------------------------------------- -->
							<%cor_linha = "#ffcccc"%>
							<%'if IsNumeric(nr_faltas) then%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Faltas</b></td>
								<td align="right">&nbsp;<%=total_faltas_mes%></td>
								<td align="right">&nbsp;<%'=formatnumber(nr_salario_liq,2)%></td>
								<td align="right"><%=formatnumber(nr_valor_falta,2)%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Atrasos</b></td>
								<td align="right">&nbsp;<%if nr_atrasos <> "" Then response.write(nr_atrasos&"h")%></td>
								<td align="right">&nbsp;</td>
								<td align="right"><%if nr_atrasos <> "" Then response.write(formatNumber(nr_desc_atrasos,2))%></td>
								<%nr_salario_liq = nr_salario_liq - nr_desc_atrasos%>
							</tr>
							<%'end if%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Cesta Básica</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;</td>
								<td align="right"><%xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=10 AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
											Set rs_refeic = dbconn.execute(xsql)
											if not rs_refeic.EOF AND nr_salario > 0 then
												cesta_b = rs_refeic("nr_valor")
												response.write(formatnumber(cesta_b,2))
											end if%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Vale Refeição</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;</td>
								<td align="right"><%if int(expediente) > 6 then
														xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=2 AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
															Set rs_refeic = dbconn.execute(xsql)
															if not rs_refeic.EOF then
																valor_refeicao = rs_refeic("nr_valor")
															'	response.write(formatnumber(valor_refeicao))
															'else
															'	valor_refeicao = "0"
															'	response.write(valor_refeicao)
															end if
													else
														valor_refeicao = "0"
														'response.write(valor_refeicao)
													end if
													'valor_refeicao = 3.50%>
													<%if nr_salario <> "" AND nr_vrefeic_cancela = 0 Then
														response.write(valor_refeicao)
													else
														response.write("0")
														valor_refeicao = "0"
													end if%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Vale Transporte</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;</td>
								<td align="right"><%if nr_salario > "0" AND nr_vtransp_cancela <> 1 Then
														v_trasnsp = nr_salario*0.06
														response.write(formatnumber(v_trasnsp))
													else
														response.write("0")
													end if%></td>
							</tr>
							<%base_inss = abs(nr_salario_liq) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(hora_extra) + abs(dsr_he)%></td>
							<%cor_linha = "#ffcccc"%>
							<tr style="background-color:<%=cor_linha%>;">
								<%xsql = "Select * From TBL_inss Where dt_atualizacao <= '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
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
										faixa_inss = pri_aliq
										'response.write(formatnumber(inss,2))
										'response.write("1")
									elseif base_inss <= abs(faixa_2) then
										inss = base_inss * seg_aliq
										faixa_inss =  seg_aliq
										'response.write(formatnumber(inss,2))
										'response.write("2")
									elseif base_inss <= abs(faixa_3) then
										inss = base_inss * ter_aliq
										faixa_inss =  ter_aliq
										'response.write(formatnumber(inss,2))
										'response.write("3")
									end if%>
								<td align="right" colspan="3"><b>INSS</b></td>
								<td align="right">&nbsp;<%if inss > "0" AND nr_salario <> 1 Then response.write(faixa_inss*100&"%")%></td>
								<td align="right">&nbsp;<%'=base_inss%></td>
								<td align="right">&nbsp;<%if inss > "0" AND nr_salario <> 1 Then response.write(formatnumber(inss,2))%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Conv. Méd.</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;<%=nr_conv_medico%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Contribuição Sindical</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;<%=nr_contr_sind%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>Descontos Diversos</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;<%if nr_desc_diversos <> "0" Then response.write(FormatNumber(nr_desc_diversos,2))%></td>
							</tr>
							
							<%
							fgts = nr_salario*0.08
							
							total_vencimento = abs(nr_salario) + abs(nr_aj_custo) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(auxcreche) + abs(hora_extra) + abs(dsr_he)
							total_desconto = abs(nr_valor_falta) + abs(nr_desc_atrasos) + abs(cesta_b) + abs(valor_refeicao) + abs(v_trasnsp) + abs(inss) + abs(nr_conv_medico) + abs(nr_contr_sind)+ abs(nr_desc_diversos)%>
							
							<%base_irrf = abs(nr_salario_liq) + abs(nr_aj_custo) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not)  + abs(hora_extra) + abs(dsr_he)
							'base_irrf = abs(nr_salario) + abs(nr_aj_custo) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(auxcreche) + abs(hora_extra) + abs(dsr_he)
							base_irrf = base_irrf - inss 
							xsql = "Select * From TBL_IRRF Where dt_atualizacao <= '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
								Set rs_irrf = dbconn.execute(xsql)
								if not rs_irrf.EOF then
									deduc_dependente = rs_irrf("deduc_dependente")
									
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
										base_irrf = base_irrf' - deduc_dependente
									end if
								
								if base_irrf = 0 then
									irrf = "0"
									parc_deduzir = 0
									faixa_irrf = 0
								elseif base_irrf <= abs(faixa_1) then
									irrf = base_irrf * (pri_aliq/100)
									parc_deduzir = pri_deducao
									faixa_irrf = "Isento"
								elseif base_irrf <= abs(faixa_2) then
									irrf = base_irrf * (seg_aliq/100)
									parc_deduzir = seg_deducao
									faixa_irrf = seg_aliq&"%"
								elseif base_irrf <= abs(faixa_3) then
									irrf = base_irrf * (ter_aliq/100)
									parc_deduzir = ter_deducao
									faixa_irrf = ter_aliq&"%"
								elseif base_irrf <= abs(faixa_4) then
									irrf = base_irrf * (qua_aliq/100)
									parc_deduzir = qua_deducao
									faixa_irrf = qua_aliq&"%"
								elseif base_irrf > abs(faixa_5) then
									irrf = base_irrf * (qui_aliq/100)
									parc_deduzir = qui_deducao
									faixa_irrf = qui_aliq&"%"
								end if
								
									if irrf <> "" Then 
										irrf_final = irrf - parc_deduzir
										response.write(base_irrf)
									end if
								
									'	if irrf_final < 10 then
									'		irrf_final = 0
									'		faixa_irrf = "Isento"
									'		parc_deduzir = "0"
									'	end if%>
							<%if qtd_dependentes_21 > 1 AND irrf_final <> 0 then%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>IRRF - Dependente</b></td>
								<td align="right" colspan="3">&nbsp;Desc. dependente até 21a: &nbsp; &nbsp; &nbsp; <%=deduc_dependente%></td>
								<!--td align="right">&nbsp;<%'=base_irrf%></td>
								<td align="right">&nbsp;<%'response.write("+"&formatnumber(irrf_final,2))
									'if irrf <> "0" and qtd_dependentes_21 <> "0" then
										'irrf_final = irrf_final - deduc_dependente
										'	if irrf_final < "0" then
										'		irrf_final = "0"
										'	end if
									'end if%>
								</td-->
							</tr>
							<%end if
							cor_linha = "#ffffcc"%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="3"><b>IRRF</b></td>
								<td align="right">&nbsp;<%if parc_deduzir <> "0" then response.write("(-"&parc_deduzir&")")%></td>
								<td align="right">&nbsp;<%=faixa_irrf%></td>
								<td align="right">&nbsp;<%if irrf_final <> "0" then response.write(FormatNumber(irrf_final,2))%></td>
							</tr>
							
							<%total_desconto = total_desconto + irrf_final
							total_liquido = total_vencimento - total_desconto
							total_liquido = total_liquido' + auxcreche%>
							
							<tr>
								<td colspan="6"><img src="../../imagens/px.gif" alt="" width="60" height="3" hspace="0" vspace="0" border="0"></td>
							</tr>
							<%cor_linha  = "silver"%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="center" valign="middle" colspan="4" rowspan="3">Bco: <%=nm_banco%><br>AG.:<%=cd_banco_ag%><br>C/C:<%=cd_banco_conta%></td>
								<td align="center"><b>Total Vencimentos</b></td>
								<td align="center"><b>Total Descontos</b></td>
								</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" style="background-color:#ccffff;">&nbsp;<%if total_vencimento <> "" Then response.write(formatnumber(total_vencimento,2))%></td>
								<td align="right" style="background-color:#ffcccc;">&nbsp;<%if total_desconto <> "" Then response.write(formatnumber(total_desconto,2))%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>Valor Liquido</b></td>
								<td align="right" style="background-color:#ccffcc; font-size:12px;">&nbsp;<b><%if total_liquido <> 0 then response.write(formatnumber(total_liquido,2))%></b></td>
							</tr>
							<tr>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="80" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="80" height="1" hspace="0" vspace="0" border="0"></td>
							</tr>
							<tr>
								<td>Salário Base</td>
								<td>Sal. Contr. INSS</td>
								<td>Base Calc. FGTS</td>
								<td>FGTS do Mês</td>
								<td>Base Calc. IRRF</td>
								<td>Faixa IRRF</td>
							</tr>
							<tr>
								<td><%if not nr_salario = 0 then response.write(formatnumber(nr_salario,2))%></td>
								<td><%if not base_inss = 0 then response.write(formatnumber(base_inss,2))%></td>
								<td><%if nr_salario <> "" Then response.write(formatnumber(nr_salario))%></td>
								<td><%if nr_salario <> "" Then response.write(formatnumber(fgts))%></td>
								<td><%if base_irrf <> "" Then response.write(formatnumber(base_irrf,2))%></td>
								<td><%=faixa_irrf%></td>
							</tr>
						</table>
			
			
			<SCRIPT LANGUAGE="javascript">
				function JsDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o funcionario?"))
				  {
				  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir&protege_exclusao='+cod2+'');
					}
					  }
				
				function JsContatoDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o contato?"))
				  {
				  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_cont='+cod2+'&acao=apaga_contato');
					}
					  }  
					  
				
			
			</SCRIPT>