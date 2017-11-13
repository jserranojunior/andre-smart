<html>
<head>
	<title>RCM</title>
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>

<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../css/geral.htm"-->
<%
'usuario_autorizado = 4

cd_user = session("cd_codigo")
pasta_loc = "empresa\funcionario\"
arquivo_loc = "funcionarios_folhapagamento_rcm_cad.asp"

if cd_user = 4 then
	usuario_autorizado = 4
elseif cd_user = 46 then
	usuario_autorizado = 46
elseif cd_user = 5 then
	usuario_autorizado = 5
elseif cd_user = 67 then
	usuario_autorizado = 67
end if

nao_fecha_janela = request("nao_fecha_janela")

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
	
	mes_ant = dt_mes - 1
	ano_ant = dt_ano
		if mes_ant < 1 then
			mes_ant = 12
			ano_ant = ano_ant - 1
		end if
	
dia_final = ultimodiames(dt_mes,dt_ano)
	
	mesanoatual=year(now)&zero(month(now))
	mesanosel=dt_ano&zero(dt_mes)
		
	mes_competencia = dt_mes' - 1
		'if mes_competencia < 1 then
		'	mes_competencia = 12
		'	ano_competencia = dt_ano - 1
		'else
			ano_competencia = dt_ano
		'end if
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
	
	ano_hoje = year(now)
	mes_hoje = month(now)
	dia_hoje = day(now)
	
	'*****************************************************************
	'*** Verificação para travamento da inserção de dados (dia 25) ***
	'*****************************************************************
	competencia = int(ano_competencia&zero(mes_competencia)&int("25"))-int(ano_hoje&zero(mes_hoje)&zero(dia_hoje))
	'*****************************************************************
'response.write("<br>"&ano_competencia&zero(mes_competencia))&"25<br>"&year(now)&zero(month(now))&zero(day(now)&"<br>")

			if ordem_total <> "" Then
				ordem_funcionarios = ordem_total
			else
				'ordem_funcionarios = "cd_contrato,nm_nome,nm_sigla"
				ordem_funcionarios = "cd_contrato,nm_nome"
			end if
			
data_atual = final_competencia&"/"&mes_competencia&"/"&ano_competencia%>
<!--#include file="../../includes/feriados.asp"-->
<%qtd_dias_dsr = qtd_domingos_comp + qtd_feriados - qtd_feriado_dsr%>



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
	document.form.dt_dia.focus();
	}
	
function ad_noturno(){
	hora_ini = document.form.dt_hora_i.value;
	hora_fin = document.form.dt_hora_f.value;
		if (hora_ini >= '22' && hora_fin > '22'){
			diz = "noturno";}
		else{
			
		} 
	window.alert (diz);}	
</script>
</head>


<%if strtipo = "cadastro" then atualiza_origem = "onunload='window.opener.data.submit(true);'"%>
<body onload="foco();" <%=atualiza_origem%>>
<!--#include file="../../includes/arquivo_loc.asp"-->
					<table border="1" cellspacing="0" cellpadding="2" width="50" align="center" style="border-collapse:collapse">
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="funcionarios_folhapagamento_cad.asp"  method="POST">
							<input type="hidden" name="cod" value="<%=strcod%>">
							<tr>
								<td colspan="10" align="center" class="cabecalho" style="background-color:black; color:white;"><b>RCM - <%=strtipo%><%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td colspan="3" align="center" style="background-color:silver;"><b>Mês de competencia:</b></td>
								<td colspan="7" align="center" style="background-color:silver;"><b><%=mesdoano(mes_ant)&" - "&mesdoano(mes_competencia)&"/"&ano_competencia%></b>															
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
							
							<%outras_variaveis = " AND cd_codigo="&strcod
							
							ano_competencia_i = ano_competencia
							mes_competencia_i = mes_competencia
							dia_competencia_i = 25
							
							ano_competencia_f = ano_competencia
							mes_competencia_f = mes_competencia
							final_competencia_f = 1'final_competencia'30
							
							'xsql = "up_funcionario_contrato_lista3 @dt_data_i='"&mes_competencia&"/20/"&ano_competencia&"', @dt_data_f='"&mes_competencia&"/20/"&ano_competencia&"', @dt_atualizacao = '"&mes_ant&"/21/"&ano_ant&"', @ordem_funcionarios='"&ordem_funcionarios&"', @outras_variaveis='"&outras_variaveis&"'"
							'xsql = "up_funcionario_contrato_lista4 @dt_data_i='"&mes_competencia_i&"/"&dia_competencia_i&"/"&ano_competencia_i&"', @dt_data_f='"&mes_competencia_f&"/"&final_competencia_f&"/"&ano_competencia_f&"', @dt_atualizacao = '"&mes_ant&"/"&dia_final&"/"&ano_ant&"', @ordem_funcionarios='"&ordem_funcionarios&"', @outras_variaveis='"&outras_variaveis&"'"
							
							xsql = "up_funcionario_contrato_lista4 @dt_data_i='"&mes_competencia_i&"/"&dia_competencia_i&"/"&ano_competencia_i&"', @dt_data_f='"&mes_competencia_f&"/"&final_competencia_f&"/"&ano_competencia_f&"', @dt_atualizacao = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @ordem_funcionarios='"&ordem_funcionarios&"', @outras_variaveis='"&outras_variaveis&"'"
							Set rs = dbconn.execute(xsql)
								
								'if NOT rs.EOF then
								nm_nome = rs("nm_nome")
								cd_unidade = rs("cd_unidade")
								nm_unidade = rs("nm_unidade")
								nm_sigla = rs("nm_sigla")
								
								hr_entrada = rs("hr_entrada")
								hr_saida = rs("hr_saida")
									periodo = hour(hr_entrada) - hour(hr_saida)
										if periodo < 0 then
											nm_periodo = "diurno"
											cd_periodo = 1
										elseif periodo > 0 then
											nm_periodo = "noturno"
											cd_periodo = 2
										end if
										horario_trabalho = zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))&" até "&zero(hour(hr_saida))&":"&zero(minute(hr_saida))
										
								'> expediente = rs("expediente")
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
									
								'>cd_vrefeic_cancela = rs("cd_vrefeic_cancela")
								'>nr_vrefeic_cancela = rs("nr_vrefeic_cancela")
									if  not IsNumeric(nr_vrefeic_cancela) Then nr_vrefeic_cancela = "0"
								
								'>cd_vtransp_cancela = rs("cd_vtransp_cancela")
								'>nr_vtransp_cancela = rs("nr_vtransp_cancela")
									if  not IsNumeric(nr_vtransp_cancela) Then nr_vtransp_cancela = "0"
								'end if
								%>
								
							<!--tr style="color:white;">
								<td align="center" colspan="10">&nbsp;
									<%=qtd_dias_comp%> dias &nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_domingos_comp%> dom.&nbsp; &nbsp; &nbsp; &nbsp; <%=qtd_feriados%> feriado<%if qtd_feriados > 1 then%>s<%end if%> no mês  &nbsp; &nbsp; &nbsp; &nbsp; <a href="javascript:void(0);"return false;"  onClick="window.open('../../empresa/funcionario/funcionarios_folhapagamento_demonstrativo.asp?tipo=cadastro&cod=<%=strcod%>&dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','','width=600,height=600,scrollbars=1')">VER</a></td>
							</tr-->
							<!--tr><td colspan="10">
								xsql = "up_funcionario_contrato_lista4 @dt_data_i='<%=mes_competencia_i%>/<%=dia_competencia_i%>/<%=ano_competencia_i%>', @dt_data_f='<%=mes_competencia_f%>/<%=final_competencia_f%>/<%=ano_competencia_f%>', @dt_atualizacao = '<%=dt_mes%>/<%=dia_final%>/<%=dt_ano%>', @ordem_funcionarios='<%=ordem_funcionarios%>', @outras_variaveis='<%=outras_variaveis%>'"
							</td></tr-->
							<tr>
								<td colspan="2" align="center">
									<b>Funcion&aacute;rio</b></td>
								<td colspan="4" onClick="window.open('../../empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&cod=<%=strcod%>','','width=630,height=600,scrollbars=1')">
									<%=nm_nome%></td>
								<td align="center">
									<%=nm_sigla%></td>
								<!--td align="center"><%'=tipo_expediente%></td-->
								<td colspan="2" align="center">
									<%=horario_trabalho%></td>
								<td align="center"><%if cd_periodo = 1 then%>
														<img src="../../imagens/ic_periodo_1.gif" alt="Diurno" width="13" height="14" border="0">
													<%elseif cd_periodo = 2 then%>
														<img src="../../imagens/ic_periodo_2.gif" alt="Noturno" width="13" height="14" border="0">
													<%end if%></td>
							</tr>							
							<tr>
								<td colspan="10" align="center">&nbsp;
									
									<%'if x="" then%>
									
									<table border="1" width="100" style="border-collapse:collapse;">
										<%strsql = "SELECT * FROM View_funcionario_rcm WHERE cd_funcionario="&strcod&" AND dt_rcm BETWEEN '"&mes_ant&"/21/"&ano_ant&"' AND '"&mes_competencia&"/20/"&ano_competencia&"' order by dt_rcm"
										Set rs_rcm = dbconn.execute(strsql)%>
										<tr valign="bottom" style="background-color:silver;">
											<td align="center"><b>Dia</b><br><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
											<td align="center"><b>Horário H.E.</b><br><img src="../../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
											<td align="center"><b>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
											<td align="center"><b>Ad. Noturno</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
											<td align="center"><b>VR</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
											<td align="center"><b>VT</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
											<td align="center"><b>Folga<br />enfer.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
											<td align="center"><b>Falta justf.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
											<td align="center"><b>Falta injust.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
											<td align="center"><b>Atraso</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
											<td align="center"><b>Obs.</b><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
											<td colspan="3" class="xno_print">&nbsp;</td>
										</tr>
										<%while not rs_rcm.EOF
											cd_rcm = rs_rcm("cd_codigo")
											dt_rcm = rs_rcm("dt_rcm")
												mes_rcm = month(dt_rcm)
												dt_rcm = day(rs_rcm("dt_rcm"))
											
											nr_ad_noturno = rs_rcm("nr_ad_noturno")
												if nr_ad_noturno > 0 then total_ad_noturno = int(total_ad_noturno) + nr_ad_noturno
												
											nr_folga_enf = rs_rcm("nr_folga_enf")
												if nr_folga_enf <> 0 then 
													total_folga_enf = total_folga_enf + nr_folga_enf
												end if
                                            nr_falta_justif = rs_rcm("nr_falta_justif")
											nr_falta_injust = rs_rcm("nr_falta_injust")
												if nr_falta_justif = 0 AND nr_falta_injust = 0 then
													hora_inicio = rs_rcm("hora_inicio")
													if not isnull(hora_inicio) then
														hora_i = zero(hour(hora_inicio))&":"&zero(minute(hora_inicio))
													else
														hora_i = " - - "
													end if
													 
													hora_fim = rs_rcm("hora_fim")
													if not isnull(hora_fim) then
														hora_f = zero(hour(hora_fim))&":"&zero(minute(hora_fim))
													else
														hora_i = " - - "
													end if
												else
													hora_i = " - - "
													hora_f = " - - "
												end if
												
												
											qtd_he = rs_rcm("qtd_he")
												'qtd_he = qtd_he/60 
												if qtd_he > 0 then total_he = total_he + qtd_he
												
												
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
											
											'#### Avisa se o funcionario está estorando HE. ###
											if total_he < 1200 then
												cor_aviso = 1
												cor_linha_aviso = "#ccffff"
												cor_aviso = "black"
												img_aviso = "Aviso_1.gif"
											elseif total_he <= 2160 then
												cor_aviso = 2
												cor_linha_aviso = "orange"
												cor_aviso = "black"
												img_aviso = "Aviso_2.gif"
											elseif total_he > 2160 then
												cor_aviso = 3
												cor_linha_aviso = "red"
												cor_aviso = "white"
												img_aviso = "Aviso_3.gif"
											end if
											'###################################################
											%>
											<%if ultimo_mes_rcm <> mes_rcm then%>
												<tr><td colspan="14"><img src="../../imagens/blackdot.gif" alt="" width="600" height="1" border="0"></td></tr>
											<%end if%>
										<tr>
											<td align="center" valign="top"><%=dt_rcm%></td>
											<td align="center" valign="top"><%=hora_i%>&nbsp;-&nbsp;<%=hora_f%></td>											
											<td align="center" valign="top"><%=sinal_arredonda%><%=formatahora(qtd_he)%></td>
											<td align="center" valign="top"><%if nr_ad_noturno > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center" valign="top"><%if nr_vr_extra > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center" valign="top"><%if nr_vt_extra > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center" valign="top"><%if nr_folga_enf > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
                                            <td align="center" valign="top"><%if nr_falta_justif > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center" valign="top"><%if nr_falta_injust > 0 then%><b>X</b><%else%>&nbsp;<%end if%></td>
											<td align="center" valign="top"><%if nr_atraso > 0 then%><%=nr_atraso%><%else%>&nbsp;<%end if%></td>
											<td align="center" valign="top" class="xno_print"><%if nm_obs <> "" then%><a href="javascript:void(0);" return false;" onmouseout="hideandshow_a(<%=cd_rcm%>);"><div class="<%=cd_rcm%>" style="position:static; left:460px; top:100px; background-color:white; border:2px solid white; display:none;"><%=nm_obs%></div></a><a href="javascript:void(0);" return false;" onmouseover="hideandshow_a(<%=cd_rcm%>);"><b style="color:orange; display:block;" class="<%=cd_rcm%>">X</b></a><%else%>&nbsp;<%end if%></td>
												
												<td align="center" valign="top" class="xok_print"><%if nm_obs <> "" then%><div class="<%=cd_rcm%>"style="position:static; left:460px; top:100px; background-color:white; border:2px solid white; display:block;"><%=nm_obs%></div><%end if%></td>
											<td align="center" valign="top" class="xno_print">&nbsp;<!--img src="../../imagens/ic_editar.gif" alt="" width="13" height="14" border="0"--></td>
											<td align="center" valign="top" class="xno_print"><%if strtipo = "cadastro" then%><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsDelete('<%=cd_rcm%>','<%=strcod%>','<%=dt_rcm%>','<%=dt_mes%>','<%=dt_ano%>');"><%else%>&nbsp;<%end if%></td>
											
										</tr>
										<%ultimo_mes_rcm = mes_rcm
										
										dt_rcm = ""
										hora_i = ""
										hora_f = ""
										qtd_he = "0"
										nr_vr_extra = ""
										nr_vt_extra = ""
										nr_falta_justif = "NULL"
										nr_falta_injust = "Null"
										
										rs_rcm.movenext
										wend%>
										<tr><td colspan="14"><img src="../../imagens/blackdot.gif" alt="" width="600" height="1" border="0"></td></tr>
										<tr bgcolor="silver">
											<td align="center"><b>TOTAL</b></td>
											<td align="center">&nbsp;<%if total_he > 1200 Then%><img src="../../imagens/<%=img_aviso%>"><%end if%></td>											
											<td align="center">&nbsp;<b><%if total_he > 0 then%><%=total_mes(total_he)%><%'=(total_he)%><%end if%></b></td>
											<td align="center">&nbsp;<b><%=total_ad_noturno%></b></td>
											<td align="center">&nbsp;<b><%=total_vr_extra%></b></td>
											<td align="center">&nbsp;<b><%=total_vt_extra%></b></td>
											<td align="center">&nbsp;<b><%=total_folga_enf%></b></td>
											<td align="center">&nbsp;<b><%=total_falta_justif%></b></td>
											<td align="center">&nbsp;<b><%=total_falta_injust%></b></td>
											<td align="center">&nbsp;<b><%=total_atraso%></b></td>
											<td align="center">&nbsp;</td>
											<td align="center" class="xno_print">&nbsp;</td>
											<td align="center" class="xno_print">&nbsp;</td>
										</tr>
									</table>
							
							<%'else%>
							<!--&nbsp;-->
							<%'end if%>
							
								</td>
							</tr>
						<%if strtipo = "cadastro" then%>
			<!-- ------------------------------------------- V A R I A V E I S ------------------------------------------- -->							
						<%if competencia > -1 AND competencia < 100 AND cd_user = usuario_autorizado then
							img_espacador = "px"%>
							<tr style="background-color:<%=cor_linha%>;">
								<td align="center"><b>Dia</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="40" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>H. Inicio</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="85" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>H. Fim</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="85" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>Ad. Noturno</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="50" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>VR Extra</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="50" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>VT Extra</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="50" height="1" hspace="0" vspace="0" border="0"></td>
                                <td align="center"><b>Folga<br />Enfer.</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="50" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>Faltas<br>( Justif.)</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="50" height="1" hspace="0" vspace="0" border="0"></td>
								<td align="center"><b>Faltas<br>( Injust )</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="50" height="1" hspace="0" vspace="0" border="0"></td>
								<%if strcat = "Andre" then%><td align="center"><b>Qtd. DSR<br>( Semanas )</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="50" height="1" hspace="0" vspace="0" border="0"></td><%end if%>
								<td align="center"><b>Atrasos<br>( Horas )</b><br><img src="../../imagens/<%=img_espacador%>.gif" alt="" width="50" height="1" hspace="0" vspace="0" border="0"></td>
								
							</tr>
							<%styleinput = "style='font-size:11px; border:1px solid #cdcdcd;'"%>
							<%cor_linha = "#ccffcc"
							'competencia = 1%>
							<tr style="color=<%=cor_aviso%>;background-color:<%=cor_linha_aviso%>;">
								<td align="center">&nbsp;<input type="text" name="dt_dia" size="2" maxlength="2" value="<%if dia_sel <> "" Then response.write(zero(dia_sel))%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" <%=styleinput%>>
								<td align="center">&nbsp;<input type="text" name="dt_hora_i" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" <%=styleinput%>>:<input type="text" name="dt_min_i" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" <%=styleinput%>></td>
								<td align="center">&nbsp;<input type="text" name="dt_hora_f" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" <%=styleinput%>>:<input type="text" name="dt_min_f" size="2" maxlength="2" <%=styleinput%>></td>
								<td align="center">&nbsp;<input type="checkbox" name="nr_ad_noturno" value="1" <%=styleinput%>></td>
								<td align="center">&nbsp;<input type="checkbox" name="nr_vr_extra" value="1" <%=styleinput%>></td>
								<td align="center">&nbsp;<input type="checkbox" name="nr_vt_extra" value="1" <%=styleinput%>></td>
								<td align="center">&nbsp;<input type="checkbox" name="nr_folga_enf" value="1" <%=styleinput%> <%if nr_folga_enf >="1" then%> disabled <%end if%> ></td>
								<td align="center">&nbsp;<input type="radio" name="nr_faltas" value="1"<%=styleinput%>></td>
								<td align="center">&nbsp;<input type="radio" name="nr_faltas" value="2"<%=styleinput%>></td>
								<td align="center">&nbsp;<input type="text" name="nr_atrasos" size="2" maxlength="2" <%=styleinput%>></td>
								
							</tr>
						<%end if%>
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
							
						
							
							
			
			
							
							<%if competencia > -1 AND competencia < 100 AND cd_user = usuario_autorizado then%>	
								<tr style="color=<%=cor_aviso%>;background-color=<%=cor_linha_aviso%>;">
									<td colspan="7" lign="center">
										Obs: <input type="text" name="nm_obs" size="30" maxlength="500" <%=styleinput%>></td>
									<td colspan="3" align="center">&nbsp;
										<input type="image" src="../../imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0">
									</td>
								</tr>	
							<%elseif competencia < 0 AND cd_user = usuario_autorizado then%>
								<tr>
									<td align="center" colspan="1" bgcolor="<%=cor_linha%>">
										<b style="color:red;">O período selecionado já está encerrado.</b>
									</td>
								</tr>
							<%'elseif competencia > 100 AND cd_user = usuario_autorizado then
							elseif competencia > 100 AND cd_user = usuario_autorizado then%>
								<tr>
									<td align="center" colspan="1" bgcolor="<%=cor_linha%>">
										<b style="color:red;">ATENÇÃO! O período selecionado é futuro <%=competencia%>.</b>
									</td>
								</tr>
							<%end if%>
							
							<%'if competencia > -1 AND competencia < 100 AND cd_user = usuario_autorizado then
							if competencia > -1 AND competencia < 203 AND cd_user = usuario_autorizado then%>
								<%'=competencia%>							
								<tr>
									<td colspan="9"><input type="checkbox" name="nao_fecha_janela" value="1" <%if nao_fecha_janela = 1 then response.write("Checked")%> > Impede o fechamento desta janela.&nbsp;<!--* hora extra arredondada--></td>
									<td align="center"><img src="../../imagens/ic_print.gif" alt="imprimir" width="16" height="16" border="0" onclick="visualizarImpressao();"></td>
								</tr>
								</form>
							<%else%>
								<tr>
									<td colspan="1">&nbsp;</td>
									<td align="center"><img src="../../imagens/ic_print.gif" alt="imprimir" width="16" height="16" border="0" onclick="visualizarImpressao();"></td>
								</tr>
							<%end if%>
						<%End if%>
					</table>
						
<SCRIPT LANGUAGE="javascript">
	function JsDelete(cod,cod1,cod2,cod3,cod4)
		{//alert(cod);
			if (confirm ("Tem certeza que deseja excluir o registro?"))
				{
					location.href = '../acoes/rcm_acao.asp?cd_rcm='+cod+'&cd_funcionario='+cod1+'&dt_dia='+cod2+'&dt_mes='+cod3+'&dt_ano='+cod4+'&acao=rcm_excluir';
				}
				
		}
</script>
			
</body>