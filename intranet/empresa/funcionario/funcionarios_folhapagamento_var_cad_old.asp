<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%
cd_user = session("cd_codigo")

strcod = request("cod")
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strcat = request("cat")
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
	document.form.qtd_plantoes.focus();}
</script>
<body onload="foco();" onunload="window.opener.data.submit(true);">
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print" style="border-collapse:collapse">
						
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="funcionarios_folhapagamento_cad.asp"  method="POST">
							<input type="hidden" name="tipo" value="cadastro">
							<input type="hidden" name="cod" value="<%=strcod%>">
							<tr>
								<td colspan="10" align="center" class="cabecalho" style="background-color:black; color:white;"><b>FECHAMENTO MENSAL - VARIÁVEIS<%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td align="center" style="background-color:silver;"><b>Mês do pagamento:</b></td>
								<td align="center" style="background-color:silver;" colspan="8"><b><%=mesdoano(dt_mes)&"/"&dt_ano%></b>
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
										'qtd_hextra = "36"
										he_marca = "1"
									end if
								
								cd_qtd_plantoes = rs("cd_qtd_plantoes")
								qtd_plantoes = rs("qtd_plantoes")
								
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
								
								cd_faltas_justif = rs("cd_faltas_justif")
								nr_faltas_justif = rs("nr_faltas_justif")
									if  not IsNumeric(nr_faltas_justif) Then nr_faltas_justif = "0"
								
								cd_faltas_injust = rs("cd_faltas_injust")
								nr_faltas_injust = rs("nr_faltas_injust")
									if  not IsNumeric(nr_faltas_injust) Then nr_faltas_injust = "0"
								
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
									if  not IsNumeric(nr_desconto_he) Then nr_desconto_he = "0"
								
								cd_vr_extra = rs("cd_vr_extra")
								nr_vr_extra = rs("nr_vr_extra")
									if  not IsNumeric(nr_vr_extra) Then nr_vr_extra = "0"
								
								cd_vt_extra = rs("cd_vt_extra")
								nr_vt_extra = rs("nr_vt_extra")
									if  not IsNumeric(nr_vt_extra) Then nr_vt_extra = "0"%>
								
							<tr style="color:white;">
								<td align="center" colspan="10">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês  &nbsp; &nbsp; &nbsp; &nbsp; <a href="javascript:void(0);"return false;"  onClick="window.open('../../empresa/funcionario/funcionarios_folhapagamento_demonstrativo.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','','width=600,height=600,scrollbars=1')">VER</a></td>
							</tr>
							<tr>
								<td align="center"><b>Funcion&aacute;rio</b></td>
								<td colspan="9" onClick="window.open('../../empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&cod=<%=strcod%>','','width=630,height=600,scrollbars=1')"><%=nm_nome%></td>
							</tr>
							<tr>
								<td align="center"><b>Unidade</b></td>
								<td align="center"><%=nm_sigla%></td>
								<td align="center"><%=tipo_expediente%></td>
								<td align="right" colspan="7">&nbsp;</td>
							</tr>
							<tr><td colspan="10">&nbsp;</td></tr>
			<!-- ------------------------------------------- V A R I A V E I S ------------------------------------------- -->							
							<tr style="background-color:<%=cor_linha%>;">
								<td align="right"><b>&nbsp;</b><br><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>Qtd. Plant&otilde;es</b><input type="hidden" name="cd_qtd_plantoes" value="<%=cd_qtd_plantoes%>"><br><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>Adicional Noturno</b><input type="hidden" name="cd_ad_noturno" value="<%=cd_ad_noturno%>"></td>
								<td align="center"><b>Hora<br>Extra</b><%if he_marca = "1" Then response.write("*")%><input type="hidden" name="cd_hextra" value="<%=cd_hextra%>"><br></td>
								<td align="center"><b>Faltas<br>( Justif.)</b><input type="hidden" name="cd_faltas_justif" value="<%=cd_faltas_justif%>"><br></td>
								<td align="center"><b>Faltas<br>( Injust )</b><input type="hidden" name="cd_faltas_injust" value="<%=cd_faltas_injust%>"><br></td>
								<%if strcat = "Andre" then%><td align="center"><b>Qtd. DSR<br>( Semanas )</b><input type="hidden" name="cd_dsr_faltas" value="<%=cd_dsr_faltas%>"><br></td><%end if%>
								<td align="center"><b>Atrasos<br>( Horas )</b><input type="hidden" name="cd_atraso" value="<%=cd_atraso%>"><br></td>
								<td align="center"><b>VR Extra</b><input type="hidden" name="cd_vr_extra" value="<%=cd_vr_extra%>"><br></td>
								<td align="center"><b>VT Extra</b><input type="hidden" name="cd_vt_extra" value="<%=cd_vt_extra%>"><br></td>
							</tr>
							<%styleinput = "style='font-size:11px; border:1px solid #cdcdcd;'"%>
							<%cor_linha = "#ccffcc"%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="center"><b>Variáveis</b> <%'=competencia%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="qtd_plantoes" size="3" maxlength="3" onFocus="nextfield ='ad_noturno';" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="ad_noturno" size="3" maxlength="3" onFocus="nextfield ='qtd_he';" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="qtd_he" size="3" maxlength="3" onFocus="nextfield ='nr_faltas';" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="nr_faltas_justif" size="3" maxlength="2" onFocus="nextfield ='nr_dsr_faltas_justif';" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="nr_faltas_injust" size="3" maxlength="2" onFocus="nextfield ='nr_dsr_faltas_injust';" <%=styleinput%>><%else%>-<%end if%></td>
								<%if strcat = "Andre" then%><td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="nr_dsr_faltas" size="2" maxlength="1" onFocus="nextfield ='nr_atrasos'" <%=styleinput%>><%else%>-<%end if%></td><%end if%>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="nr_atrasos" size="2" maxlength="1" onFocus="nextfield =''" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="nr_vr_extra" size="2" maxlength="3" onFocus="nextfield =''" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="nr_vt_extra" size="2" maxlength="3" onFocus="nextfield =''" <%=styleinput%>><%else%>-<%end if%></td>
							</tr>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="center"><b>Quant.</b></td>
								<td align="center"><%=zero(qtd_plantoes)%></td>
								<td align="center"><%=zero(qtd_dia_noturno)%> <%if qtd_dia_noturno = 1 then%> dia<%elseif qtd_dia_noturno > 1 then%> dias<%end if%></td>
								<td align="center"><%if qtd_hextra >= 1 then%><%=zero(qtd_hextra)%>h<%end if%>&nbsp;</td>
								<td align="center">&nbsp;<%=nr_faltas_justif%><%if nr_faltas_justif = 1 then%> dia<%elseif nr_faltas_justif > 1 then%> dias<%end if%></td>
								<td align="center">&nbsp;<%=nr_faltas_injust%><%if nr_faltas_injust = 1 then%> dia<%elseif nr_faltas_injust > 1 then%> dias<%end if%></td>
								<%if strcat = "Andre" then%><td align="center">&nbsp;<%=nr_dsr_faltas%><%if nr_dsr_faltas >= 1 then%> dom.<%end if%></td><%end if%>
								<td align="center">&nbsp;<%=nr_atrasos%><%'if nr_atrasos >= 1 then%> <%'end if%></td>
								<td align="center">&nbsp;<%=nr_vr_extra%></td>
								<td align="center">&nbsp;<%=nr_vt_extra%></td>
							</tr>
							
						
							
							
			
			
							
							
							<tr>
								<!--td align="right" colspan="3">&nbsp;</td-->
								<td align="center" colspan="9">&nbsp;<%if competencia >= -1 then%><input type="image" src="../../imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0"><%end if%></td>
							</tr>
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