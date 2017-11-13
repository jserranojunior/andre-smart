<!--#include file="../../../includes/util.asp"-->
<!--#include file="../../../includes/inc_open_connection.asp"-->
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<%
cd_user = session("cd_codigo")
pasta_loc = "empresa\funcionario\ausencias\"
arquivo_loc = "funcionarios_cad_ausencias.asp"

strcod = request("cod")
	if strcod = "" Then
		strcod = request("id")
	end if
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")
strcd_ferias = request("cd_ferias")

dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

dia_atual = day(now)
mes_atual = month(now)
ano_atual = year(now)

if dia_sel = "" Then dia_sel = dia_atual
if mes_sel = "" Then mes_sel = mes_atual
if ano_sel = "" Then ano_sel = ano_atual

If strcod <> "" And strmsg = "" Then
	
'xsql ="up_Funcionario_Lista_por_cd_codigo @cd_codigo="&strcod&""
xsql =" VIEW_funcionario_contrato_lista where cd_funcionario="&strcod&""
Set rs = dbconn.execute(xsql)

	If NOT rs.EOF Then
		'str_funcionario = rs("cd_codigo")
		str_funcionario = rs("cd_funcionario")
		strregime_trabalho = rs("cd_contrato")
		strmatricula = rs("cd_matricula") 
		strnome = rs("nm_nome")
		str_sexo = rs("cd_sexo")
		strfoto = rs("nm_foto")
		'strregime_contrato = rs("cd_regime_trabalho")
		
		cd_numreg = rs("cd_numreg")
		str_numreg = rs("cd_numreg")
	End If

else

End if
%>

<script language="JavaScript">



function preenche_cid(cid){
	document.Form.cd_cid.value=cid;
}
</script>
<script language="JavaScript">
function validar_ausencia(shipinfo){
	if (shipinfo.cd_motivo_falta.value.length==""){window.alert ("Informe o motivo da ausencia.");shipinfo.cd_motivo_falta.focus();return (false);}
	if (shipinfo.per_dia_f.value.length==""){window.alert ("O campo dia final não foi preenchido.");shipinfo.per_dia_f.focus();return (false);}
	if (shipinfo.per_mes_f.value.length==""){window.alert ("O campo mês final não foi preenchido.");shipinfo.per_mes_f.focus();return (false);}
	if (shipinfo.per_ano_f.value.length==""){window.alert ("O campo ano final não foi preenchido.");shipinfo.per_ano_f.focus();return (false);}
	return (true);}
</script>
<!--include file="../../../js/ajax.js"-->
<!--#include file="../../../js/ajax2.asp"-->
<!--#include file="js/ausencias.js"-->

<!--#include file="../../../includes/arquivo_loc.asp"-->

					<% If strmsg <> "" then%>
						<table  border="0" cellspacing="0" cellpadding="0" class="txt_cinza">
							<tr class="no_print">
							 <td>
								<b><%=strmsg%></b>
							 </td>
							</tr>
							<tr class="no_print">
							 <td align=center><br><br><a href="coop_cooperados_lista.asp"><img src="imagens/bt_lista.gif" alt="" width="119" height="22" border="1"></a></td>
							</tr>
						</table>
					<%else%>
					
						<table border="1" cellspacing="0" cellpadding="2" align="center"><!-- OK 2-->
									
								<form name="Form" action="../../acoes/funcionarios_acao.asp"  method="POST"  onsubmit="return validar_ausencia(document.Form);">
								<input type="hidden" name="acao" value="<%=stracao%>">
								<input type="hidden" name="cod" value="<%=strcod%>">
								<input type="hidden" name="cd_ferias" value="<%=strcd_ferias%>">
								<input type="hidden" name="cd_user" value="<%=cd_user%>">
								<input type="hidden" name="cd_funcionario" value="<%=str_funcionario%>">
								<%if strfoto = "" then
									strfoto = "Vdlap.gif"
								end if%>
									<tr>
										<td colspan="4" align="center" class="cabecalho"><b>REGISTRO DE FÉRIAS DE COLABORADOR</b></td>
									</tr>
									<tr>
										<td rowspan="5" align="left" valign="top" style="color:#7d7d7d;"><img src="../../../foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"></td>
										<td colspan="3" style="background-color:gray; color:white;"><b>Dados do colaborador</b></td>
									</tr>
									<tr>
										<td style="background-color:silver; color:white;"><b>Matrícula</b></td>
										<td style="background-color:silver; color:white;" colspan="2"><b>Empresa</b></td>
									</tr>
									<tr class="no_print">
										<td><b><%=strmatricula%></b></td>
										<td colspan="2"><%
											strsql = "Select * From TBL_tipo_contrato Where cd_codigo ="&strregime_trabalho&""
						  					Set rs_contr = dbconn.execute(strsql)%>
											<%Do While Not rs_contr.eof
											nm_regime_trab = rs_contr("nm_regime_trab")
											cd_empresa= rs_contr("cd_codigo")%>
											<input type="hidden" name="cd_empresa" value="<%=cd_empresa%>">	
											<b><%=nm_regime_trab%></b>
											<%rs_contr.movenext
											loop
											
												if strcd_ferias <> "" then
													strsql = "SELECT * From TBL_funcionario_ausencias where cd_codigo='"&strcd_ferias&"'"
						  							Set rs_contr = dbconn.execute(strsql)
														while not rs_contr.EOF 
															str_afastamento_i = rs_contr("dt_afastamento_i")
																str_dia_i = zero(day(str_afastamento_i))
																str_mes_i = zero(month(str_afastamento_i))
																str_ano_i = year(str_afastamento_i)
															str_afastamento_f = rs_contr("dt_afastamento_f")
																str_dia_f = zero(day(str_afastamento_f))
																str_mes_f = zero(month(str_afastamento_f))
																str_ano_f = year(str_afastamento_f)
															str_ferias = datediff("d",str_afastamento_i,str_afastamento_f&" 23:59")+1
															
															str_pagamento_ferias = rs_contr("dt_pagamento_ferias")
																str_dia_pag = zero(day(str_pagamento_ferias))
																str_mes_pag = zero(month(str_pagamento_ferias))
																str_ano_pag = year(str_pagamento_ferias)
															str_valor = rs_contr("valor_pag")
																if isnumeric(str_valor) then
																	strvalor_pag = formatNumber(str_valor,2)
																else
																	strvalor_pag = "0,00"
																end if
															str_obs = rs_contr("nm_obs")
														rs_contr.movenext
														wend
												else ' *** Se o codigo das férias não for informado, preenche a data inicial ***
													str_dia_i = zero(dia_sel)
													str_mes_i = zero(mes_sel)
													str_ano_i = ano_sel
												end if
											
											rs_contr.close
											Set rs_contr = nothing%>
										</td>								
									</tr>
									<tr>
										<td style="background-color:silver; color:white;"><b>Nome</b></td>
										<td style="background-color:silver; color:white;" colspan="2"><b>Unidade</b></td>
									</tr>
									<tr>
										<td><b><%=strnome & strnm_sobrenome%></b></td>
										<td colspan="2">
											<%strsql ="SELECT * From View_funcionario_unidade where cd_funcionario = "&strcod&" and cd_suspenso <> 1 order by dt_atualizacao desc"
										  	Set rs_unidade = dbconn.execute(strsql)
											
											do while not rs_unidade.EOF
											cd_unidade = rs_unidade("cd_unidade")
											nm_unidade = rs_unidade("nm_unidade")
											dt_atualizacao = rs_unidade("dt_atualizacao")
											
											rs_unidade.movenext
											loop%>
											<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
											 <b><%=nm_unidade%></b></a>
											<%rs_unidade.close
											Set rs_unidade = nothing%>
										</td>							
									</tr>
								</table>
								<br>
								<table border="1" cellspacing="0" cellpadding="2" align="center">
									<tr>
										<td style="width: 90px; background-color:silver; color:white;"><b>Tempo</b></td>
										<td style="width: 150px; background-color:silver; color:white;"><b>Início</b></td>
										<td style="width: 150px; background-color:silver; color:white;"><b>Fim</b></td>
									</tr>
									
									<tr>
										<td valign="top">
											<span id='divDias'>
												<input type="text" name="qtd_dias_d" id="qtd_dias_d" size="3" maxlength="4" value="<%=str_ferias%>" onKeyup="calcula_data(this.value, document.getElementById('dt_dia').value, document.getElementById('dt_mes').value, document.getElementById('dt_ano').value);"> dia(s)
											</span>
										</td>	
										<td valign="top">
												<input type="text" name="dt_dia" id="dt_dia" size="2" maxlength="2" value="<%=str_dia_i%>" onchange="calcula_data(document.getElementById('qtd_dias_d').value, this.value, document.getElementById('dt_mes').value, document.getElementById('dt_ano').value);" value="<%=zero(dia_sel)%>"; onKeyPress="ChecarTAB();" onKeyup="Mostra(this, 2)" onFocus="PararTAB(this);">
												<input type="text" name="dt_mes" id="dt_mes" size="2" maxlength="2" value="<%=str_mes_i%>" onchange="calcula_data(document.getElementById('qtd_dias_d').value, document.getElementById('dt_dia').value, this.value, document.getElementById('dt_ano').value);" value="<%=zero(mes_sel)%>"; onKeyPress="ChecarTAB();" onKeyup="Mostra(this, 2)" onFocus="PararTAB(this);">
												<input type="text" name="dt_ano" id="dt_ano" size="4" maxlength="4" value="<%=str_ano_i%>" onchange="calcula_data(document.getElementById('qtd_dias_d').value, document.getElementById('dt_dia').value, document.getElementById('dt_mes').value, this.value);" value="<%=ano_sel%>"></td>
										<td valign="top">
											<span id='divTfinal'>
												<input type="text" name="ajax_dia_f" size="2" maxlength="2" value="<%=str_dia_f%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(per_dia_f)%>"  onChange="calcula_dias(document.getElementById('dt_dia').value, document.getElementById('dt_mes').value, document.getElementById('dt_ano').value, this.value, document.getElementById('per_mes_f').value, document.getElementById('per_ano_f').value);">
												<input type="text" name="ajax_mes_f" size="2" maxlength="2" value="<%=str_mes_f%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(per_mes_f)%>"  onChange="calcula_dias(document.getElementById('dt_dia').value, document.getElementById('dt_mes').value, document.getElementById('dt_ano').value, document.getElementById('per_dia_f').value, this.value, document.getElementById('per_ano_f').value);">
												<input type="text" name="ajax_ano_f" size="4" maxlength="4" value="<%=str_ano_f%>" onKeyup="calcula_dias(document.getElementById('dt_dia').value, document.getElementById('dt_mes').value, document.getElementById('dt_ano').value, document.getElementById('per_dia_f').value, document.getElementById('per_mes_f').value, this.value);">
											</span></td>
									</tr>
									<tr>
										<td style="width: 150px; background-color:silver; color:white;"><b>Data pagamento</b></td>
										<td style="width: 100px; background-color:silver; color:white;"><b>Valor</b></td>
										<td style="width: 100px; background-color:silver; color:white;"><b>Observações</b></td>
									</tr>
									<tr>	
										<td>
											<input type="text" name="dt_dia_pgto" size="2" maxlength="2" value="<%=str_dia_pag%>" onKeyPress="ChecarTAB();" onKeyup="Mostra(this, 2)" onFocus="PararTAB(this);">/
											<input type="text" name="dt_mes_pgto" size="2" maxlength="2" value="<%=str_mes_pag%>" onKeyPress="ChecarTAB();" onKeyup="Mostra(this, 2)" onFocus="PararTAB(this);">/
											<input type="text" name="dt_ano_pgto" size="4" maxlength="4" value="<%=str_ano_pag%>" onKeyPress="ChecarTAB();" onKeyup="Mostra(this, 4)" onFocus="PararTAB(this);">
										</td>
										<td>R$ <input type="text" name="valor_pgto_ferias" size="8" maxlength="8" value="<%=strvalor_pag%>"></td>
										<td rowspan="3"><textarea cols="15" rows="3" name="nm_obs"><%=str_obs%></textarea></td>
									</tr>										
									<tr>
										<!--td colspan="4" valign="top"-->
											<input type="hidden" name="per_dia_i" id="per_dia_i" size="2" maxlength="2" value="<%=str_dia_i%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(dia_sel)%>" readonly="1" style="border-style:none;">
											<input type="hidden" name="per_mes_i" id="per_mes_i" size="2" maxlength="2" value="<%=str_mes_i%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(mes_sel)%>" readonly="1" style="border-style:none;">
											<input type="hidden" name="per_ano_i" id="per_ano_i" size="4" maxlength="4" value="<%=str_ano_i%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="<%=ano_sel%>" readonly="1" required="." style="border-style:none;">
											<!--input type="text" name="per_hor_i" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">:<input type="text" name="per_min_i" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value=""-->
														
										<!--	Até:--> 
											<input type="hidden" name="per_dia_f" id="per_dia_f" size="2" maxlength="2" value="<%=str_dia_f%>" readonly="1" required="." style="border-style:none;">
											<input type="hidden" name="per_mes_f" id="per_mes_f" size="2" maxlength="2" value="<%=str_mes_f%>" readonly="1" required="." style="border-style:none;">
											<input type="hidden" name="per_ano_f" id="per_ano_f" size="4" maxlength="4" value="<%=str_ano_f%>" readonly="1" required="." style="border-style:none;">
											
											<input type="hidden" name="qtd_dias" id="qtd_dias" size="2" maxlength="3" value="<%=qtd_dias%>" readonly="1" required="." style="border-style:none;">
										<!--/td-->
									</tr>
									<tr>	
										<td><input type="submit" value="Ok"></td>
										<td><input type="Button" value="Limpar" onmouseup=></td>
									</tr>
									<!--tr class="no_print">
										<td colspan="4"><span id='divCid'></span></td>
									</tr-->
									</form>
							</table>
								<br>
								<br>
								<table align="center" border="1">
									<tr>
										<td colspan="7" align="center" style="background-color:gray; color:white;">Registro de férias</td>
									</tr>
									<tr style="background-color:silver;">
										<td align="center" colspan="3">Periodo de férias</td>
										<td>Data Pagto.</td>
										<td>Valor R$</td>
										<td>Obs.</td>
									</tr>
									<%'strsql ="SELECT * From View_funcionario_ausencias where cd_funcionario = "&str_funcionario&" order by dt_ausencia"
									strsql ="SELECT TOP 6 * From TBL_funcionario_ausencias where cd_funcionario = "&str_funcionario&" order by dt_afastamento_i desc"
									  	Set rs_falta = dbconn.execute(strsql)
										
											do while not rs_falta.EOF
												cd_ferias = rs_falta("cd_codigo")
												cd_motivo = rs_falta("cd_motivo")
												'subcategoria = rs_falta("subcategoria")
												subcategoria = rs_falta("nm_cid")
												
												dt_afastamento_i = rs_falta("dt_afastamento_i")
												dt_afastamento_f = rs_falta("dt_afastamento_f")
												dt_pagamento_ferias = rs_falta("dt_pagamento_ferias")
												valor_pag = rs_falta("valor_pag")
													if isnumeric(valor_pag) then
														valor_pag = formatNumber(valor_pag,2)
													end if
												nm_obs = rs_falta("nm_obs")
												
												'per_ausencia = rs_falta("per_ausencia")
												per_ausencia = datediff("n",dt_afastamento_i,dt_afastamento_f&" 23:59")
													per_ausencia = per_ausencia
													periodo_ausencia = per_ausencia 
														
														if int(per_ausencia) =< 59 then 'retorna Minutos
															atrasos = atrasos + per_ausencia
															dias_faltas = dias_faltas
																per_ausencia = "<b>"&per_ausencia&" m</b>"
																
														elseif int(per_ausencia) < 1439 then 'retorna Horas
															atrasos = atrasos + per_ausencia
															dias_faltas = dias_faltas
																'if per_ausencia = 59 then
																per_ausencia = per_ausencia' + 1
																
																per_ausencia = "<b>"&int((per_ausencia)/60)&" h "&(per_ausencia mod 60)&"m</b>"
																
														elseif int(per_ausencia) => 1439 then 'retorna dias
															per_ausencia = ((per_ausencia+1)/60)/24
															atrasos = atrasos
																dias_faltas = dias_faltas + per_ausencia
															
															ausencia = (per_ausencia)
																if per_ausencia > (1440/59)/24 then
																	plural= "s"
																else
																	plural = ""
																end if
																	per_ausencia = "<b>"&int(per_ausencia)&" dia"&plural&"</b>"
															
														Else
															per_ausencia = per_ausencia
															atrasos = atrasos
															dias_faltas = dias_faltas
														end if%>
									<tr>
										<td><%=zero(day(dt_afastamento_i))&"/"&zero(month(dt_afastamento_i))&"/"&right(year(dt_afastamento_i),2)%></td>
										<td><%=zero(day(dt_afastamento_f))&"/"&zero(month(dt_afastamento_f))&"/"&right(year(dt_afastamento_f),2)%></td>
										<td align="right"><%=per_ausencia%>&nbsp;</td>
										<td align="center"><%=zero(day(dt_pagamento_ferias))&"/"&zero(month(dt_pagamento_ferias))&"/"&right(year(dt_pagamento_ferias),2)%></td>
										<td><%=valor_pag%></td>
										<td><%=left(nm_obs,25)%><%if len(nm_obs) > 25 Then response.write(" ...")%></td>
										<td><a href="funcionarios_cad_ausencias.asp?cod=<%=strcod%>&cd_ferias=<%=cd_ferias%>&acao=update_ferias"><img src="../../../imagens/ic_editar.gif" alt="" width="13" height="14" border="0"></a></td>
										<td><img src="../../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onmouseup="JsausenciaDelete(<%=cd_ferias%>,<%=str_funcionario%>);"></td>
									</tr>
											<%rs_falta.movenext
											loop%>									
								</table>
						<br>
						<br>
							
					
									 								
							<%End if%>
		<SCRIPT LANGUAGE="javascript">
			function JsausenciaDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse registro de férias?"))
				  		{
				 			 //alert("teste"+cod2);
							 window.location='../../acoes/funcionarios_acao.asp?cd_ferias='+cod+'&cd_funcionario='+cod2+'&acao=delete_ferias';
						}
				   } 
		</SCRIPT>
	