<html>
<head>
	<title>RCM</title>
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

dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

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
</head>
<body onload="foco();" onunload="window.opener.data.submit(true);">
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print" style="border-collapse:collapse">
						
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="funcionarios_folhapagamento_cad.asp"  method="POST">
							<input type="hidden" name="cod" value="<%=strcod%>">
							<tr>
								<td colspan="9" align="center" class="cabecalho" style="background-color:black; color:white;"><b>RCM - Inclusão<%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td align="center" style="background-color:silver;"><b>Mês de competencia:</b></td>
								<td align="center" style="background-color:silver;" colspan="8"><b><%=mesdoano(mes_competencia)&"/"&ano_competencia%></b>															
							</tr>
						</form>
						
			<!-- ---------------------------------------------------------------- -->				
												
						<form name="form" action="../acoes/rcm_acao.asp"  method="post">
							<input type="hidden" name="dt_mes" value="<%=mes_competencia%>">
							<input type="hidden" name="dt_ano" value="<%=ano_competencia%>">
							<input type="hidden" name="cd_user" value="<%=cd_user%>">
							<input type="hidden" name="acao" value="rcm">
							<input type="hidden" name="jan" value="1">
							<input type="hidden" name="cd_funcionario" value="<%=strcod%>">
							<%'End if%>
							
							<%xsql = "up_funcionario_folhapagamento_individual @cd_funcionario="&strcod&", @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&mes_competencia&"/1/"&ano_competencia&"', @dt_f = '"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
								Set rs = dbconn.execute(xsql)
								nm_nome = rs("nm_nome")
								cd_unidade = rs("cd_unidade")
								nm_unidade = rs("nm_unidade")
								nm_sigla = rs("nm_sigla")
								
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
									
								cd_vrefeic_cancela = rs("cd_vrefeic_cancela")
								nr_vrefeic_cancela = rs("nr_vrefeic_cancela")
									if  not IsNumeric(nr_vrefeic_cancela) Then nr_vrefeic_cancela = "0"
								
								cd_vtransp_cancela = rs("cd_vtransp_cancela")
								nr_vtransp_cancela = rs("nr_vtransp_cancela")
									if  not IsNumeric(nr_vtransp_cancela) Then nr_vtransp_cancela = "0"
								
								%>
								
							<!--tr style="color:white;">
								<td align="center" colspan="10">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês  &nbsp; &nbsp; &nbsp; &nbsp; <a href="javascript:void(0);"return false;"  onClick="window.open('../../empresa/funcionario/funcionarios_folhapagamento_demonstrativo.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','','width=600,height=600,scrollbars=1')">VER</a></td>
							</tr-->
							<tr>
								<td align="center"><b>Funcion&aacute;rio</b></td>
								<td colspan="5" onClick="window.open('../../empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&cod=<%=strcod%>','','width=630,height=600,scrollbars=1')"><%=nm_nome%></td>
								<td align="center"><%=nm_sigla%></td>
								<td align="center"><%=tipo_expediente%></td>
								<td align="center">&nbsp;</td>
							</tr>							
							<tr>
								<td colspan="9" align="center">&nbsp;
									<table border="1" style="border-collapse:collapse;">
										<%strsql = "SELECT * FROM View_funcionario_rcm WHERE cd_funcionario="&strcod&" AND dt_rcm BETWEEN '"&mes_competencia&"/1/"&ano_competencia&"' AND '"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
										Set rs_rcm = dbconn.execute(strsql)%>
										<tr valign="bottom" style="background-color:silver;">
											<td align="center"><b>Dia</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
											<td align="center"><b>Horário</b><br><img src="../../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
											<td align="center"><b>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
											<td align="center"><b>Ad. Noturno</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
											<td align="center"><b>VR</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
											<td align="center"><b>VT</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
											<td align="center"><b>Falta justf.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
											<td align="center"><b>Falta injust.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
											<td align="center"><b>Atraso</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
											<td align="center"><b>Obs.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
											
										</tr>
										<%while not rs_rcm.EOF
										
										dt_rcm = day(rs_rcm("dt_rcm"))
										
										nr_ad_noturno = rs_rcm("nr_ad_noturno")
											if nr_ad_noturno > 0 then total_ad_noturno = int(total_ad_noturno) + nr_ad_noturno
											
										nr_falta_justif = rs_rcm("nr_falta_justif")
										nr_falta_injust = rs_rcm("nr_falta_injust")
											if nr_falta_justif = 0 AND nr_falta_injust = 0 then
												hora_inicio = rs_rcm("hora_inicio")
												hora_i = zero(hour(hora_inicio))&":"&zero(minute(hora_inicio))
												
												hora_fim = rs_rcm("hora_fim")
												hora_f = zero(hour(hora_fim))&":"&zero(minute(hora_fim))
											else
												hora_i = " - - "
												hora_f = " - - "
											end if
										qtd_he = rs_rcm("qtd_he")
											if qtd_he > 0 then total_he = int(total_he) + qtd_he
											
										nr_vr_extra = rs_rcm("nr_vr_extra")
											if nr_vr_extra > 0 then total_vr_extra = int(total_vr_extra) + int(nr_vr_extra)
										nr_vt_extra = rs_rcm("nr_vt_extra")
											if nr_vt_extra > 0 then total_vt_extra = int(total_vt_extra) + int(nr_vt_extra)
										
										nr_falta_justif = rs_rcm("nr_falta_justif")
											if nr_falta_justif > 0 then total_falta_justif = int(total_falta_justif) + int(nr_falta_justif)
										nr_falta_injust = rs_rcm("nr_falta_injust")
											if nr_falta_injust > 0 then total_falta_injust = int(total_falta_injust) + int(nr_falta_injust)
										
										nr_atraso = rs_rcm("nr_atraso")
											if nr_atraso > 0 then total_atraso = int(total_atraso) + int(nr_atraso)
										
										nm_obs = rs_rcm("nm_obs")
										%>
										<tr>
											<td align="center"><%=dt_rcm%></td>
											<td align="center"><%=hora_i%>&nbsp;-&nbsp;<%=hora_f%></td>											
											<td align="center"><%=qtd_he/60%></td>
											<td align="center"><%if nr_ad_noturno > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center"><%if nr_vr_extra > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center"><%if nr_vt_extra > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center"><%if nr_falta_justif > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center"><%if nr_falta_injust > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center"><%if nr_atraso > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center" style="color:orange;"><%if nm_obs <> "" then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center"><img src="../../imagens/ic_editar.gif" alt="" width="13" height="14" border="0"></td>
											<td align="center"><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0"></td>
										</tr>
										<%dt_rcm = ""
										hora_i = ""
										hora_f = ""
										qtd_he = "0"
										nr_vr_extra = ""
										nr_vt_extra = ""
										nr_falta_justif = "NULL"
										nr_falta_injust = "Null"
										
										rs_rcm.movenext
										wend%>
										<tr><td colspan="11"><img src="../../imagens/blackdot.gif" alt="" width="490" height="1" border="0"></td></tr>
										<tr bgcolor="silver">
											<td align="center"><b>TOTAL</b></td>
											<td align="center">&nbsp;</td>											
											<td align="center">&nbsp;<b><%=total_he/60%></b></td>
											<td align="center">&nbsp;<b><%=total_ad_noturno%></b></td>
											<td align="center">&nbsp;<b><%=total_vr_extra%></b></td>
											<td align="center">&nbsp;<b><%=total_vt_extra%></b></td>
											<td align="center">&nbsp;<b><%=total_falta_justif%></b></td>
											<td align="center">&nbsp;<b><%=total_falta_injust%></b></td>
											<td align="center">&nbsp;<b><%=total_atraso%></b></td>
											<td align="center">&nbsp;</td>
											<td align="center">&nbsp;</td>
											<td align="center">&nbsp;</td>
										</tr>
									</table>
							
							
							
							</td></tr>
			<!-- ------------------------------------------- V A R I A V E I S ------------------------------------------- -->							
							<tr style="background-color:<%=cor_linha%>;">
								<td align="center"><b>Data</b><br><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>H. Inicio</b><br><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>H. Fim</b><br><img src="../../imagens/px.gif" alt="" width="60" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>Ad. Noturno</b></td>
								<td align="center"><b>VR Extra</b></td>
								<td align="center"><b>VT Extra</b></td>
								<td align="center"><b>Faltas<br>( Justif.)</b></td>
								<td align="center"><b>Faltas<br>( Injust )</b></td>
								<%if strcat = "Andre" then%><td align="center"><b>Qtd. DSR<br>( Semanas )</b></td><%end if%>
								<td align="center"><b>Atrasos<br>( Horas )</b></td>
								
							</tr>
							<%styleinput = "style='font-size:11px; border:1px solid #cdcdcd;'"%>
							<%cor_linha = "#ccffcc"%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="dt_dia" size="2" maxlength="2" value="<%if dia_sel <> "" Then response.write(zero(dia_sel))%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" <%=styleinput%>><!--/<b><%'=zero(mes_competencia)%></b>/<b><%'=ano_competencia%></b>--><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="dt_hora_i" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" <%=styleinput%>>:<input type="text" name="dt_min_i" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="text" name="dt_hora_f" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" <%=styleinput%>>:<input type="text" name="dt_min_f" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="checkbox" name="nr_ad_noturno" value="1" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="checkbox" name="nr_vr_extra" value="1" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="checkbox" name="nr_vt_extra" value="1" <%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="radio" name="nr_faltas" value="1"<%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="radio" name="nr_faltas" value="2"<%=styleinput%>><%else%>-<%end if%></td>
								<td align="center">&nbsp;<%if competencia >= -1 then%><input type="checkbox" name="nr_atrasos" value="1" <%=styleinput%>><%else%>-<%end if%></td>
								
							</tr>
							<!--tr style="background-color:<%=cor_linha%>;">
								<td align="center"><%=zero(qtd_plantoes)%></td>
								<td align="center"><%=zero(qtd_dia_noturno)%> <%if qtd_dia_noturno = 1 then%> dia<%elseif qtd_dia_noturno > 1 then%> dias<%end if%></td>
								<td align="center"><%if qtd_hextra >= 1 then%><%=zero(qtd_hextra)%>h<%end if%>&nbsp;</td>
								<td align="center">&nbsp;<%=nr_faltas_justif%><%if nr_faltas_justif = 1 then%> dia<%elseif nr_faltas_justif > 1 then%> dias<%end if%></td>
								<td align="center">&nbsp;<%=nr_faltas_injust%><%if nr_faltas_injust = 1 then%> dia<%elseif nr_faltas_injust > 1 then%> dias<%end if%></td>
								<%if strcat = "Andre" then%><td align="center">&nbsp;<%=nr_dsr_faltas%><%if nr_dsr_faltas >= 1 then%> dom.<%end if%></td><%end if%>
								<td align="center">&nbsp;<%=nr_atrasos%><%'if nr_atrasos >= 1 then%> <%'end if%></td>
								<td align="center">&nbsp;<%=nr_vr_extra%></td>
								<td align="center">&nbsp;<%=nr_vt_extra%></td>
							</tr-->
							
						
							
							
			
			
							
							
							<tr>
								<td align="center" colspan="3" bgcolor="<%=cor_linha%>">Obs: <input type="text" name="nm_obs" size="30" maxlength="500" <%=styleinput%>></td>
								<td align="center" colspan="6">&nbsp;<%if competencia >= -1 then%><input type="image" src="../../imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0"><%end if%></td>
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
							
						</table>		
			
</body>