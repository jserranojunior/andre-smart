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
strvar = request("var")

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
	
	competencia = int(ano_competencia&zero(mes_competencia))-int(year(now)&zero(month(now)))
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
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "nr_salario"; // nome do primeiro campo do site
netscape = "";
ver = navigator.appVersion; len = ver.length;
for(iln = 0; iln < len; iln++) if (ver.charAt(iln) == "(") break;
	netscape = (ver.charAt(iln+1).toUpperCase() != "C");
	function keyDown(DnEvents) {
		// ve quando e o netscape ou IE
		k = (netscape) ? DnEvents.which : window.event.keyCode;
		if (k == 13) { // preciona tecla enter
		if (nextfield == 'done') {
			//alert("viu como funciona?");
			eval('document.form.' + nextfield + '.focus()');
			return false;
			//return true; // envia quando termina os campos
		} else {
			// se existem mais campos vai para o proximo
			eval('document.form.' + nextfield + '.focus()');
			return false;
		}
	}
}
document.onkeydown = keyDown; // work together to analyze keystrokes
if (netscape) document.captureEvents(Event.KEYDOWN|Event.KEYUP);
// End -->
function foco(){
	document.form.ad_noturno.focus(); }
</script>
<body onload="foco();">
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print" style="border-collapse:collapse">
						
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="funcionarios_folhapagamento_cad.asp"  method="POST">
							<input type="hidden" name="tipo" value="cadastro">
							<input type="hidden" name="cod" value="<%=strcod%>">
							<tr>
								<td colspan="6" align="center" class="cabecalho" style="background-color:black; color:white;"><b>FOLHA DE PAGAMENTO - Valores e Variáveis<%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td align="center" style="background-color:silver;"><b>Mês do pagamento:</b></td>
								<td align="center" style="background-color:silver;" colspan="4"><b><%=mesdoano(dt_mes)&"/"&dt_ano%></b>
									<!--select name="dt_mes">
										<%for i = 1 to 12%>
											<option value="<%=i%>" <%if i = int(dt_mes) then%>SELECTED<%end if%>><%=mesdoano(i)%></option>
										<%next%>
									</select> / <input type="text" name="dt_ano" value="<%=dt_ano%>" size="4" maxlength="4">
									<input type="submit" name="ok" value="Ver"--></td>
								<td align="center" style="background-color:silver;"><b>&nbsp;</b></td>						
							</tr>
						</form>
						
			<!-- ---------------------------------------------------------------- -->				
												
						<form name="form" action="../acoes/folha_pagamento_acao.asp"  method="POST">
							<input type="hidden" name="dt_mes" value="<%=mes_competencia%>">
							<input type="hidden" name="dt_ano" value="<%=ano_competencia%>">
							<input type="hidden" name="cd_user" value="<%=cd_user%>">
							<input type="hidden" name="jan" value="1">
							<input type="hidden" name="cd_funcionario" value="<%=strcod%>">
							<%'End if%>
							
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

							'xsql = "up_funcionario_rh_lista_individual @cd_funcionario="&strcod&", @dt_atualizacao='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
							xsql = "up_funcionario_folhapagamento_individual @cd_funcionario="&strcod&", @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&mes_competencia&"/1/"&ano_competencia&"', @dt_f = '"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
							'xsql = "up_funcionario_folhapagamento_individual @cd_funcionario="&strcod&", @dt_pagamento_i='"&mes_pagamento&"/1/"&ano_pagamento&"', @dt_pagamento_f='"&mes_pagamento&"/"&final_pagamento&"/"&ano_pagamento&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'"
								Set rs = dbconn.execute(xsql)
								nm_nome = rs("nm_nome")
								cd_unidade = rs("cd_unidade")
								nm_unidade = rs("nm_unidade")
								nm_sigla = rs("nm_sigla")
								cd_hospital = rs("cd_hospital")
								cd_sexo = rs("cd_sexo")
								
								cd_salario = rs("cd_salario")
								nr_salario = abs(rs("nr_salario"))
									if  not IsNumeric(nr_salario) Then nr_salario = "0"
									nr_salario_dia = nr_salario/30
								
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
											salario_hora = nr_salario_dia/carga_diaria
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
										total_atrasos = nr_atrasos * salario_hora
										nr_salario_liq = nr_salario_liq - total_atrasos
									end if
									if  not IsNumeric(nr_atrasos) Then nr_atrasos = "0"
								
								cd_credito_he = rs("cd_credito_he")
								nr_credito_he = rs("nr_credito_he")
									if  not IsNumeric(nr_credito_he) Then nr_credito_he = "0"
								
								cd_desconto_he = rs("cd_desconto_he")
								nr_desconto_he = rs("nr_desconto_he")
									if  not IsNumeric(nr_desconto_he) Then nr_desconto_he = "0"%>
								
							<tr style="color:white;">
								<td align="center" colspan="6">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês  &nbsp; &nbsp; &nbsp; &nbsp; <a href="javascript:void(0);"return false;"  onClick="window.open('../../empresa/funcionario/funcionarios_folhapagamento_demonstrativo.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','','width=600,height=600,scrollbars=1')">VER</a></td>
							</tr>
							<tr>
								<td align="right"><b>Funcion&aacute;rio</b></td>
								<td colspan="5" onClick="window.open('../../empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&cod=<%=strcod%>','','width=630,height=600,scrollbars=1')"><%=nm_nome%></td>
							</tr>
							<tr>
								<td align="right"><b>Unidade</b></td>
								<td align="center"><%=nm_sigla%></td>
								<td align="center"><%=tipo_expediente%></td>
								<td align="right" colspan="3">&nbsp;</td>
							</tr>
							<tr><td colspan="6">&nbsp;</td></tr>
			<!-- ------------------------------------------- V A R I A V E I S ------------------------------------------- -->							
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>&nbsp;</b></td>
								<td align="center"><b>Adicional Noturno</b><input type="hidden" name="cd_ad_noturno" value="<%=cd_ad_noturno%>"></td>
								<td align="center"><b>Hora<br>Extra</b><%if he_marca = "1" Then response.write("*")%><input type="hidden" name="cd_hextra" value="<%=cd_hextra%>"></td>
								<td align="center"><b>Qtd. Faltas<br>( Dias )</b><input type="hidden" name="cd_faltas" value="<%=cd_faltas%>"></td>
								<td align="center"><b>Qtd. DSR<br>( Semanas )</b><input type="hidden" name="cd_dsr_faltas" value="<%=cd_dsr_faltas%>"></td>
								<td align="center"><b>Atrasos<br>( Horas )</b><input type="hidden" name="cd_atraso" value="<%=cd_atraso%>"></td>
							</tr>
							<%styleinput = "style='font-size:11px; border:1px solid #cdcdcd;'"%>
							<%cor_linha = "#ccffcc"%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>Variáveis</b> <%'=competencia%></td>
								<td align="center">&nbsp;<%if competencia >= 0 then%><input type="text" name="ad_noturno" size="3" maxlength="3" onFocus="nextfield ='qtd_he';" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= 0 then%><input type="text" name="qtd_he" size="3" maxlength="3" onFocus="nextfield ='nr_faltas';" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= 0 then%><input type="text" name="nr_faltas" size="3" maxlength="2" onFocus="nextfield ='nr_dsr_faltas';" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= 0 then%><input type="text" name="nr_dsr_faltas" size="2" maxlength="1" onFocus="nextfield ='nr_atrasos'" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= 0 then%><input type="text" name="nr_atrasos" size="2" maxlength="1" onFocus="nextfield =''" <%=styleinput%>><%else%>-<%end if%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>Quant.</b></td>
								<td align="right"><%=zero(qtd_dia_noturno)%> <%if qtd_dia_noturno = 1 then%> dia<%elseif qtd_dia_noturno > 1 then%> dias<%end if%></td>
								<td align="right"><%if qtd_hextra >= 1 then%><%=zero(qtd_hextra)%>h<%end if%>&nbsp;</td>
								<td align="right">&nbsp;<%=nr_faltas%><%if nr_faltas = 1 then%> dia<%elseif nr_faltas > 1 then%> dias<%end if%></td>
								<td align="right">&nbsp;<%=nr_dsr_faltas%><%if nr_dsr_faltas >= 1 then%> dom.<%end if%></td>
								<td align="right">&nbsp;<%=nr_atrasos%><%'if nr_atrasos >= 1 then%> <%'end if%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>Valor</b></td>
								<td align="right">
												<%if qtd_dia_noturno > 0 then
													ad_noturno = (((nr_salario*0.40)/carga_mensal)*7)*qtd_dia_noturno
													qtd_dias_mes = qtd_dias_comp-qtd_dias_dsr
													dsr_ad_not = ad_noturno/qtd_dias_mes
														dsr_ad_not = formatnumber(dsr_ad_not*qtd_dias_dsr,2)
													
													response.write(formatnumber(ad_noturno,2))
												else
													ad_noturno = "0"
													response.write(ad_noturno)
												end if%></td>
								<td align="right">
									<%if qtd_hextra > 0 then
										hora_extra = ((nr_salario/carga_mensal)*1.9)*qtd_hextra
										qtd_dias_mes = qtd_dias_comp-qtd_dias_dsr
										dsr_he = hora_extra/qtd_dias_mes
											dsr_he = formatnumber(dsr_he*qtd_dias_dsr,2)
										
										response.write(formatnumber(hora_extra,2))
									else
										hora_extra = "0"
										response.write(hora_extra)
										dsr_he = "0"
									end if%></td>
								<td align="right">&nbsp;<%if nr_faltas >=1 then  response.write(FormatNumber(nr_faltas*nr_salario_dia,2))%></td>
								<td align="right">&nbsp;<%if nr_dsr_faltas >=1 then  response.write(FormatNumber(nr_dsr_faltas*nr_salario_dia,2))%></td>
								<td align="right">&nbsp;<%if nr_atrasos >=1 then  response.write(FormatNumber(total_atrasos,2))%></td>
							</tr>
							<%if qtd_dia_noturno <> "" OR qtd_dia_noturno <> "" Then%>
							<tr>
								<td align="right"><b>DSR</b></td>
								<td align="right"><%if ad_noturno <> "0" then response.write(dsr_ad_not)%></td>
								<td align="right"><%if nr_he <> "0" then response.write(dsr_he)%></td>
								<td colspan="3">&nbsp;</td>
							</tr>
							<%end if%>
							<%if strvar <> "1" Then%>
							<tr>
								<td align="right" colspan="2">&nbsp;</td>
								<td align="center"><b>Qtd.</b></td>
								<td align="center"><b>Valores</b></td>
								<td align="center" colspan="2"><b>Alterações</b></td>
							</tr>
			<!-- ------------------------------------------- V E N C I M E N T O S ------------------------------------------- -->
							<%cor_linha = "#ccffff"
							styleinput = "style='font-size:9px; border:1px solid #cdcdcd;'"%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="2"><b>Sal&aacute;rio</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_salario" value="<%=cd_salario%>"></td>
								<td align="right">&nbsp;<%=nr_salario%></td>
								<td align="left" colspan="2">&nbsp;<%if competencia >= 0 then%><input type="text" name="nr_salario" size="9" maxlength="11" onFocus="nextfield ='nr_aj_custo';" <%=styleinput%>><%end if%></td>								
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="2"><b>Ajuda de Custo</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_aj_custo" value="<%=cd_aj_custo%>"></td>
								<td align="right">&nbsp;<%=nr_aj_custo%></td>
								<td align="left" colspan="2">&nbsp;<%if competencia >= 0 then%><input type="text" name="nr_aj_custo" size="9" maxlength="11" onFocus="nextfield ='insalubridade_forca';" <%=styleinput%>><%end if%></td>								
							</tr>
							<%'*** verifica se deve forçar o pagamento de insalubridade ***
							if nr_insalubridade_forca = 1 then cd_hospital = 1
							
								if cd_hospital = 1 AND nr_salario > 0 then
									insalubridade = nr_salminimo * 0.1
								else
									insalubridade = "0"
								end if%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="2"><b>Insalubridade</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_insalubridade_forca" value="<%=cd_insalubridade_forca%>"></td>
								<td align="right">&nbsp;<%if insalubridade <> "0" then response.write(FormatNumber(insalubridade,2))%></td>
								<td align="left" colspan="2">Forçar pagto. &nbsp;
									<%if competencia >= 0 then%>
										<input type="checkbox" name="insalubridade_forca" value="1" <%if nr_insalubridade_forca = 1 Then%>checked<%end if%> onFocus="nextfield ='nr_credito_he'">
									<%else%> 
										<%if nr_insalubridade_forca = 1 Then 
											response.write("<img src='../../imagens/check_ok.gif' alt='' width='25' height='12' border='0'>")
										else
											response.write("<img src='../../imagens/check_inativo.gif' alt='' width='25' height='12' border='0'>")
										end if%>
									<%end if%></td>
							</tr>
							
							
							
							
							<%'xsql_dep = ("SELECT COUNT(nm_nome) AS qtd_dependentes FROM View_dependentes WHERE  (DATEDIFF(m, dt_nascimento, '"&dt_mes&"/"&dia_final&"/"&dt_ano&"') <= 251) AND (cd_funcionario = "&strcod&")")
							'Set rs_dep = dbconn.execute(xsql_dep)
							'if not rs_dep.EOF then
							'	qtd_dependentes = rs_dep("qtd_dependentes")
							'end if%>
							
							<%'xsql_dep = ("SELECT COUNT(nm_nome) AS qtd_dependentes FROM View_dependentes WHERE  (DATEDIFF(m, dt_nascimento, '"&dt_mes&"/"&dia_final&"/"&dt_ano&"') <= 83) AND (cd_funcionario = "&strcod&")")
							'Set rs_dep = dbconn.execute(xsql_dep)
							'if not rs_dep.EOF then
							'	qtd_dependentes = rs_dep("qtd_dependentes")
							'end if%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="2"><b>Crédito H. E.</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_credito_he" value="<%=cd_conv_medico%>"></td>
								<td align="right">&nbsp;<%=nr_credito_he%></td>
								<td align="left" colspan="2">&nbsp;<%if competencia >= 0 then%><input type="text" name="nr_credito_he" size="9" maxlength="11" onFocus="nextfield ='nr_faltas';" <%=styleinput%>><%end if%></td>
							</tr>
							<%cor_linha = "#e1feff"%>
							<%'*** Informa os dados o dependente ***
							xsql_dep = ("SELECT * FROM View_dependentes WHERE (DATEDIFF(m, dt_nascimento, '"&dt_mes&"/"&dia_final&"/"&dt_ano&"') <= 83) AND dt_nascimento <= '"&dt_mes&"/"&dia_final&"/"&dt_ano&"'  AND (cd_funcionario = "&strcod&")")
								Set rs_dep = dbconn.execute(xsql_dep)
								while not rs_dep.EOF
									cd_depend = rs_dep("cd_codigo")
									nm_nome = rs_dep("nm_nome")
									dt_nascimento = rs_dep("dt_nascimento")
										idade = DateDiff("m",day(dt_nascimento)&"/"&month(dt_nascimento)&"/"&year(dt_nascimento),dia_final&"/"&dt_mes&"/"&dt_ano)
										'qtd_dep_idade = qtd_dep_idade + 1	
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
									<tr style="background-color:<%=cor_linha%>;">
										<td colspan="3"> &nbsp; &nbsp; &nbsp; <%=nm_nome%> - <span style="color:<%=auxcreche_color%>;"><b></b></span><!--<br>///<%'=idade%>///--><%=formataidade(idade)%><!--//nasc.:<%'=day(dt_nascimento)&"/"&month(dt_nascimento)&"/"&year(dt_nascimento)%> - data:<%'=dia_final&"/"&dt_mes&"/"&dt_ano%>--></td>
										<td align="left" colspan="2"><input type="hidden" name="list_depend" value="<%=cd_depend%>">
											<input type="hidden" name="cd_auxcreche<%=cd_depend%>" value="<%=cd_auxcreche%>">
										Institui pag.&nbsp;<input type="checkbox" name="auxcreche_institui<%=cd_depend%>" value="1" <%if ok_auxcreche = 1 Then%>checked<%end if%> onFocus="nextfield ='vtransp_cancela'"></td>
									</tr>
								<%rs_dep.movenext
								wend%>
								
							<%cor_linha = "#ccffff"%>
							<%if qtd_dep_auxcreche > 0 Then%>
								<%auxcreche = qtd_dep_auxcreche*nr_auxcreche%>
								<tr style="background-color:<%=cor_linha%>;">
									<td align="right" colspan="2"><b>Auxílio Creche</b></td>
									<td align="right">&nbsp;</td>
									<td align="right">&nbsp;<%=auxcreche%></td>
									<td>&nbsp;</td>
								</tr>									
							<%end if%>
							<%cor_linha = "silver"%>
							<%total_vencimento = abs(nr_salario) + abs(nr_aj_custo) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(auxcreche) + abs(hora_extra) + abs(dsr_he)%>
							<!--<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>Total</b></td>
								<td align="right">&nbsp;</td>
								<td align="right"><i><%=formatnumber(total_vencimento,2)%></i></td>
								<td align="right">&nbsp;</td>
							</tr>-->
			<!-- ------------------------------------------- D E S C O N T O S ------------------------------------------- -->
							<%cor_linha = "#ffffcc"%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="2"><b>Desconto H. E.</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_desconto_he" value="<%=cd_desconto_he%>"></td>
								<td align="right">&nbsp;<%=nr_desconto_he%></td>
								<td align="left" colspan="2">&nbsp;<%if competencia >= 0 then%><input type="text" name="nr_desconto_he" size="9" maxlength="11" onFocus="nextfield ='vrefeic_cancela';" <%=styleinput%>><%end if%></td>
							</tr>
							<%xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=10 AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
								Set rs_refeic = dbconn.execute(xsql)
								if not rs_refeic.EOF AND nr_salario > 0 then
									cesta_b = rs_refeic("nr_valor")
									'response.write(formatnumber(cesta_b,2))
								else
									cesta_b = "0"
									'response.write(cesta_b)
								end if%>
							<tr style="background-color:<%=cor_linha%>;">
								<%if int(expediente) > 6 then
									xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=2 AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
										Set rs_refeic = dbconn.execute(xsql)
										if not rs_refeic.EOF then
											valor_refeicao = formatnumber(rs_refeic("nr_valor"))
										else
											valor_refeicao = "0"
										end if
								end if%>
								<td align="right" colspan="2"><b>Vale Refeição</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_vrefeic_cancela" value="<%=cd_vrefeic_cancela%>"></td>
								<td align="right"><%if nr_salario <> "" AND nr_vrefeic_cancela = 0 Then
														response.write(valor_refeicao)
													else
														response.write("0")
														valor_refeicao = "0"
													end if%></td>
								<td align="left" colspan="2">cancela &nbsp;
									<%if competencia >= 0 then%>
										<input type="checkbox" name="vrefeic_cancela" value="1" <%if nr_vrefeic_cancela = 1 Then%>checked<%end if%> onFocus="nextfield ='vtransp_cancela'">
									<%else%> 
										<%if nr_vrefeic_cancela = 1 Then 
											response.write("<img src='../../imagens/check_ok.gif' alt='' width='25' height='12' border='0'>")
										else
											response.write("<img src='../../imagens/check_inativo.gif' alt='' width='25' height='12' border='0'>")
										end if%>
									<%end if%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="2"><b>Vale Transporte</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_vtransp_cancela" value="<%=cd_vtransp_cancela%>"></td>
								<td align="right"><%if nr_salario <> "" AND nr_vtransp_cancela = 0 Then
														v_trasnsp = nr_salario*0.06
														response.write(formatnumber(v_trasnsp))
													else
														v_trasnsp = "0"
														response.write(v_trasnsp)
													end if%></td>
								<td align="left" colspan="2">cancela &nbsp;
									<%if competencia >= 0 then%>
										<input type="checkbox" name="vtransp_cancela" value="1" <%if nr_vtransp_cancela = 1 Then%>checked<%end if%> onFocus="nextfield ='conv_medico'">
									<%else%>
										<%if nr_vtransp_cancela = 1 Then
											response.write("<img src='../../imagens/check_ok.gif' alt='' width='25' height='12' border='0'>")
										else
											response.write("<img src='../../imagens/check_inativo.gif' alt='' width='25' height='12' border='0'>")
										end if%>
									<%end if%></td>
							</tr>
							<%base_inss = abs(nr_salario) + abs(insalubridade) + abs(ad_noturno) + abs(hora_extra) + abs(dsr_he) + abs(dsr_ad_not)%>
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
								elseif base_inss <= abs(faixa_2) then
									inss = base_inss * seg_aliq
								elseif base_inss <= abs(faixa_3) then
									inss = base_inss * ter_aliq
								end if%>
							<!--<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>INSS</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;<%=formatnumber(inss)%></td>
								<td align="right">&nbsp;</td>
							</tr>-->
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="2"><b>Conv. Méd.</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_conv_medico" value="<%=cd_conv_medico%>"></td>
								<td align="right">&nbsp;<%=nr_conv_medico%></td>
								<td align="left" colspan="2">&nbsp;<%if competencia >= 0 then%><input type="text" name="conv_medico" size="9" maxlength="11" onFocus="nextfield ='contr_sind'" <%=styleinput%>><%end if%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="2"><b>Contribuição Sindical</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_contr_sind" value="<%=cd_contr_sind%>"></td>
								<td align="right">&nbsp;<%=nr_contr_sind%></td>
								<td align="left" colspan="2">&nbsp;<%if competencia >= 0 then%><input type="text" name="contr_sind" size="9" maxlength="11" onFocus="nextfield ='desc_diversos'" <%=styleinput%>><%end if%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right" colspan="2"><b>Desc. Diversos</b></td>
								<td align="right">&nbsp;<input type="hidden" name="cd_desc_diversos" value="<%=cd_desc_diversos%>"></td>
								<td align="right">&nbsp;<%=nr_desc_diversos%></td>
								<td align="left" colspan="2">&nbsp;<%if competencia >= 0 then%><input type="text" name="desc_diversos" size="9" maxlength="11" onFocus="nextfield =''" <%=styleinput%>><%end if%></td>
							</tr>
							<%end if%>
		<!-- ************* FIM DOS DESCONTOS **************** -->
							<%'total_vencimento = abs(nr_salario) + abs(nr_aj_custo) + abs(insalubridade) + abs(ad_noturno) + abs(dsr_ad_not) + abs(auxcreche) + abs(hora_extra) + abs(dsr_he)
							'total_desconto = abs(cesta_b) + abs(valor_refeicao) + abs(v_trasnsp) + abs(inss) + abs(nr_conv_medico) + abs(contr_sind)%>
													
							<%base_irrf = total_vencimento - inss%>
							<%'if base_irrf <> "" Then response.write(formatnumber(base_irrf,2))%></td>
								<%xsql = "Select * From TBL_IRRF Where dt_atualizacao <= '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
									Set rs_irrf = dbconn.execute(xsql)
									if not rs_irrf.EOF then
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
									
									if base_irrf = 0 then
										irrf = "0"
										parc_deduzir = 0
										'response.write(formatnumber(irrf,2))
										'response.write("-")
									elseif base_irrf <= abs(faixa_1) then
										irrf = base_irrf * (pri_aliq/100)
										parc_deduzir = pri_deducao
										'response.write(formatnumber(irrf,2))
										'response.write("Isento")
									elseif base_irrf <= abs(faixa_2) then
										irrf = base_irrf * (seg_aliq/100)
										parc_deduzir = seg_deducao
										'response.write(formatnumber(irrf,2))
										'response.write(seg_aliq&"%")
									elseif base_irrf <= abs(faixa_3) then
										irrf = base_irrf * (ter_aliq/100)
										parc_deduzir = ter_deducao
										'response.write(formatnumber(irrf,2))
										'response.write(ter_aliq&"%")
									elseif base_irrf <= abs(faixa_4) then
										irrf = base_irrf * (qua_aliq/100)
										parc_deduzir = qua_deducao
										'response.write(formatnumber(irrf,2))
										'response.write(qua_aliq&"%")
									elseif base_irrf > abs(faixa_5) then
										irrf = base_irrf * (qui_aliq/100)
										parc_deduzir = qui_deducao
										'response.write(formatnumber(irrf,2))
										'response.write(qui_aliq&"%")
									end if%>							
							<%if irrf <> "" Then
								irrf_final = irrf - parc_deduzir
							end if%>
							<!--<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>IRRF</b></td>
								<td align="right">&nbsp;</td>
								<td align="right">&nbsp;<%'if irrf_final <> "0" Then response.write(formatnumber(irrf_final,2))%></td>
								<td align="right">&nbsp;</td>
							</tr>-->
							<%cor_linha = "silver"%>
							<%total_desconto = abs(cesta_b) + abs(valor_refeicao) + abs(v_trasnsp) + abs(inss) + abs(nr_conv_medico) + abs(nr_desc_diversos) + abs(contr_sind) + abs(irrf_final)%>
							<!--<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>Total</b></td>
								<td align="right">&nbsp;</td>
								<td align="right"><i><%=formatnumber(total_desconto,2)%></i></td>
								<td align="right">&nbsp;</td>
							</tr>-->
				<!-- ------------------------------------------- T O T A L		L Í Q U I D O ------------------------------------------- -->
							<%cor_linha = "white"%>
							<%if strvar = "1" Then%>
							<tr>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
							</tr>
							
							
							<tr>
								<!--td align="right" colspan="3">&nbsp;</td-->
								<td align="center" colspan="6">&nbsp;<%if competencia >= 0 then%><input type="image" src="../../imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0"><%end if%></td>
							</tr>
							<%else%>
							<tr>
								<td colspan="2"><img src="../../imagens/px.gif" alt="" width="90" height="1" hspace="0" vspace="0" border="0"></td>
								
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
							</tr>
							<%total_liquido = total_vencimento - total_desconto
							'total_liquido = total_liquido - irrf_final%>
							<tr>
								<!--td align="right" colspan="2">&nbsp;</td>
								<td align="right">&nbsp;<b><%'if total_liquido <> 0 then response.write(formatnumber(total_liquido,2))%></b></td-->
								<td align="center" colspan="6">&nbsp;<%if competencia >= 0 then%><input type="image" src="../../imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0"><%end if%></td>
							</tr>
							<%end if%>
							</form>
							<%If strcod <> "" then%>							
								<!--tr>
									<td colspan="4">&nbsp;
									<form name="form_ex" action="empresa/acoes/empr_funcionarios_acao.asp" id="forms" method="post"  enctype="multipart/form-data">
										<input type="hidden" name="acao" value="excluir">
										<input type="hidden" name="cod" value="<%=cd_funcionario%>">
									</form></td>
								</tr-->
							<%End if%>
							<!--tr>
								<td colspan="4">
									&nbsp;&nbsp; * 10% do salário Mínimo vigente, apenas para quem trabalha em hospital;<br>
									&nbsp;** atualizado automaticamente pelo cadastro de funcionários;<br>
									***6% do salário bruto</td>
							</tr-->
						</table>
			
			
</body>