<!--include file="../../../includes/util.asp"-->
<!--include file="../../../includes/inc_open_connection.asp"-->
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<%
cd_user = session("cd_codigo")

strcod = request("cod")
	if strcod = "" Then
		strcod = request("id")
	end if
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")

dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

dia_atual = day(now)
mes_atual = month(now)
ano_atual = year(now)

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


<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<script language="JavaScript">
function preenche_dia(dia){
	dia = document.Form.dt_dia.value
	document.Form.per_dia_i.value=dia;}
	
function preenche_mes(mes){
	mes = document.Form.dt_mes.value
	document.Form.per_mes_i.value=mes;}
	
function preenche_ano(ano){
	ano = document.Form.dt_ano.value
	document.Form.per_ano_i.value=ano;}

function preenche_data(){		
		//h_mes = parseInt(h_mes) - 1;
		dia_falta = (document.Form.per_dia_i.value);
		//alert(dia_falta)
			dia_falta = parseInt(dia_falta) + parseInt(document.Form.qtd_dias.value);
				if (dia_falta <= 9){
					dia_falta = "0"+dia_falta
				}	
				
				document.Form.per_dia_f.value=dia_falta;		
		}
function preenche_cid(cid){
	document.Form.cd_cid.value=cid;
}
</script>
<!--#include file="../../../js/ajax.js"-->
<!--#include file="js/ausencias.js"-->

					<% If strmsg <> "" then%>
						<table  border="0" cellspacing="0" cellpadding="0" class="txt_cinza">
							<tr id="no_print">
							 <td>
								<b><%=strmsg%></b>
							 </td>
							</tr>
							<tr id="no_print">
							 <td align=center><br><br><a href="coop_cooperados_lista.asp"><img src="imagens/bt_lista.gif" alt="" width="119" height="22" border="1"></a></td>
							</tr>
						</table>
					<%else%>
					<table align="center" border="0" rules="groups" width="100%">
					<tr id="no_print">
						<td>&nbsp;&nbsp;<a href="empresa.asp?tipo=ausencia&func=ativo">Listagem</a> 
						&nbsp;&nbsp;&nbsp; <a href="empresa.asp?tipo=novo">Novo</a></td>
					</tr></table>
					<br>
						<table border="0" cellspacing="0" cellpadding="2" align="center"><!-- OK 2-->
									
									<form name="Form" action="empresa/funcionario/ausencias/acoes/ausencias_acao.asp"  method="POST"  onsubmit="return validar_cad(document.forms);">
									<input type="hidden" name="cd_sexo" value="<%=str_sexo%>">
									<input type="hidden" name="acao" value="insert">
									<input type="hidden" name="cod" value="<%=strcod%>">
									<input type="hidden" name="cd_user" value="<%=cd_user%>">
									<input type="hidden" name="cd_funcionario" value="<%=str_funcionario%>">
									<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
										<tr>
											<td colspan="7" align="center" class="cabecalho"><b>REGISTRO DE AUSÊNCIAS DE COLABORADOR</b></td>
											<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
										</tr>
										<tr><td colspan="100%"><img src="../imagens/blackdot.gif" width="100%" height="1"><br>&nbsp;</td></tr>
										<tr id="no_print">
											<td rowspan="100%"><img src="../imagens/px.gif" width="35" height="1"></td>
											<td rowspan="6" align="left" valign="top" style="color:#7d7d7d;"><img src="foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"></td>
											<td colspan="3" style="background-color:gray; color:white;"><b>DADOS DO COLABORADOR</b></td>
											<td rowspan="100%"><img src="../imagens/px.gif" width="35" height="1"></td>
										</tr>
										<tr>
											<td id="no_print" style="background-color:silver; color:white;"><b>Matrícula</b></td>
											<td id="no_print" style="background-color:silver; color:white;" colspan="2"><b>Empresa</b></td>
										</tr>
										<tr id="no_print">
											<td><b><%=strmatricula%></b></td>
											<td colspan="2"><%'strsql = "Select * From TBL_tipo_contrato Where cd_codigo ="&strregime_trabalho&""
											strsql = "Select * From TBL_tipo_contrato Where cd_codigo ="&strregime_trabalho&""
						  					Set rs_contr = dbconn.execute(strsql)%>
											<%Do While Not rs_contr.eof
											nm_regime_trab = rs_contr("nm_regime_trab")
											cd_empresa= rs_contr("cd_codigo")%>
											<input type="hidden" name="cd_empresa" value="<%=cd_empresa%>">	
											<b><%=nm_regime_trab%></b>
											<%rs_contr.movenext
											loop
											rs_contr.close
											Set rs_contr = nothing%>
											</td>								
										</tr>
										<tr>
											<td rowspan="100%" id="ok_print"><img src="../imagens/px.gif" width="40" height="1"></td>
											<td rowspan="6" align="left" valign="top" id="ok_print"><img src="foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"></td>
											<td style="background-color:silver; color:white;"><b>Nome</b></td>
											<td  style="background-color:silver; color:white;" colspan="2"><b>Unidade</b></td>
										</tr>
										<tr>
											<td><b><%=strnome & strnm_sobrenome%></b></td>
											<td colspan="2">
												<%strsql ="SELECT * From View_funcionario_unidade where cd_funcionario = "&strcod&" and cd_suspenso <> 1 order by dt_atualizacao asc"
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
										<tr id="no_print"><td colspan="4"><b></b></td></tr>
										<tr><td colspan="4"><hr  noshade></td></tr>
										<tr>
											<td style="background-color:silver; color:white;"><b>Data</b></td>
											<td style="background-color:silver; color:white;"><b>Motivo</b></td>
											<td style="background-color:silver; color:white;" colspan="2"><b>Obs.</b></td>
											
										</tr>
										<tr>
											<td valign="top"><input type="text" name="dt_dia" size="3" maxlength="2" onKeyUp="preenche_dia();" value="<%=zero(dia_sel)%>">/
															<input type="text" name="dt_mes" size="3" maxlength="2" onKeyUp="preenche_mes();" value="<%=zero(mes_sel)%>">/
															<input type="text" name="dt_ano" size="5" maxlength="4" onKeyUp="preenche_ano();" value="<%=ano_sel%>"></td>
											<td valign="top">
															<select name="cd_motivo_falta">
																<%strsql = "SELECT * From TBL_funcionario_motivo"
															  	Set rs_motivo = dbconn.execute(strsql)
																
																do while not rs_motivo.EOF
																	cd_motivo = rs_motivo("cd_codigo")
																	nm_motivo = rs_motivo("nm_motivo")%>
																	<option value="<%=cd_motivo%>"><%=nm_motivo%></option>
																	<%rs_motivo.movenext
																loop%>
															</select>
											</td>
											<td colspan="2" valign="top"><textarea cols="34" rows="3" name="obs_ausencia"></textarea></td>
										</tr>										
										<tr>
											<td style="background-color:silver; color:white;"><b>CID</b></td>
											<td style="background-color:silver; color:white;" colspan="3"><b>Período de ausência</b></td>
										</tr>
										<tr id="no_print">
											<td valign="top">
															<input type="text" name="cd_cid" size="15" maxlength="20" onKeyup="mostra_cid();">
											</td>
											<td valign="middle">
												Tempo: <input type="text" name="qtd_dias" size="4" maxlength="4" onKeyup="calcula_T()" value=""> dia(s)
												<br>
												<br>
												De: &nbsp; <input type="text" name="per_dia_i" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(dia_sel)%>" readonly="1">/
															<input type="text" name="per_mes_i" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(mes_sel)%>" readonly="1">/
															<input type="text" name="per_ano_i" size="4" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="<%=ano_sel%>" readonly="1" required=".">&nbsp;
															<input type="text" name="per_hor_i" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">:
															<input type="text" name="per_min_i" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">
															<br>
															<!--/td>
											<td colspan="2" valign="middle"-->
											<span id='divTfinal'>Até: <input type="text" name="per_dia_f" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">/
															<input type="text" name="per_mes_f" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">/
															<input type="text" name="per_ano_f" size="4" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="">&nbsp;
															<input type="text" name="per_hor_f" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">:
															<input type="text" name="per_min_f" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">
											</span>
											</td>
											<td><input type="submit" value="ok"></td>				
										</tr>
										<tr id="no_print">
											<td colspan="4"><span id='divCid'></span></td>
										</tr>							
										<tr>
											<td><img src="../imagens/px.gif" width="135" height="1"></td>
											<td><img src="../imagens/px.gif" width="210" height="1"></td>
											<td><img src="../imagens/px.gif" width="135" height="1"></td>
											<td><img src="../imagens/px.gif" width="125" height="1"></td>
										</tr>							
										</form>
								</table><br>
								<br>
								
								<table align="center" border="1">
									<tr>
										<td colspan="7" align="center" style="background-color:gray; color:white;">Quantidade de faltas no mês atual</td>
									</tr>
									<tr style="background-color:silver;">
									<td>Data</td>
									<td>Motivo</td>
									<td>CID</td>
									<td colspan="3">Periodo de ausencia</td>
									<td>Obs.</td>
									</tr>
									<%'strsql ="SELECT * From View_funcionario_ausencias where cd_funcionario = "&str_funcionario&" order by dt_ausencia"
									strsql ="SELECT * From TBL_funcionario_ausencias where cd_funcionario = "&str_funcionario&" AND month(dt_afastamento_i) = "&mes_atual&" order by dt_ausencia"
									  	Set rs_falta = dbconn.execute(strsql)
										
											do while not rs_falta.EOF
												cd_falta = rs_falta("cd_codigo")
												dt_ausencia = rs_falta("dt_ausencia")
												cd_motivo = rs_falta("cd_motivo")
												'subcategoria = rs_falta("subcategoria")
												subcategoria = rs_falta("nm_cid")
												
												dt_afastamento_i = rs_falta("dt_afastamento_i")
												dt_afastamento_f = rs_falta("dt_afastamento_f")
												
												nm_obs = rs_falta("nm_obs")
												
												'per_ausencia = rs_falta("per_ausencia")
												per_ausencia = datediff("n",dt_afastamento_i,dt_afastamento_f)
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
										<td align="right"><%=zero(day(dt_ausencia))&"/"&zero(month(dt_ausencia))&"/"&year(dt_ausencia)%>&nbsp;</td>
										<td><%strsql = "SELECT * From TBL_funcionario_motivo where cd_codigo="&cd_motivo&""
											  	Set rs_motivo = dbconn.execute(strsql)
												
												do while not rs_motivo.EOF
													cd_motivo = rs_motivo("cd_codigo")
													nm_motivo = rs_motivo("nm_motivo")%>
													
													<%rs_motivo.movenext
												loop%>
										
										<%=nm_motivo%>&nbsp;</td>
										<td><%=ucase(mid(subcategoria,1,3))%><%if len(trim(subcategoria)) > 3 then response.write("."&mid(subcategoria,4,4))%>&nbsp;</td>
										<td><%=zero(day(dt_afastamento_i))&"/"&zero(month(dt_afastamento_i))&"/"&year(dt_afastamento_i)%>&nbsp;</td>
										<td><%=zero(day(dt_afastamento_f))&"/"&zero(month(dt_afastamento_f))&"/"&year(dt_afastamento_f)%></td>
										<td align="right"><%=per_ausencia%> <%'=" - "&periodo_ausencia%>&nbsp;</td>
										<td><%=nm_obs%>&nbsp;</td>
									</tr>
											<%faltas = faltas + round(periodo_ausencia)	
											
											rs_falta.movenext
											loop%>
									<tr>
										<td colspan="3">&nbsp;</td>
										<td colspan="2" align="right">Atrasos</td>
										<td align="right"><%=int(atrasos/60)%>h <%=(atrasos mod 60)%>m</td>
										<td>&nbsp;</td>
									</tr>
									<tr>
										<td colspan="3">&nbsp;</td>
										<td colspan="2" align="right">Faltas</td>
										<td><%=dias_faltas%> dia<%if dias_faltas > 1 Then%>s<%end if%></td>
										<td>&nbsp;</td>
									</tr>
									<tr><td colspan="7">&nbsp;</td></tr>
									<tr>
										<td colspan="3">&nbsp;</td>
										<td colspan="2" align="right"><b>AUSÊNCIA TOTAL</b></td>
										<td><b><%=int((faltas/60)/24)%> dia<%if dias_faltas > 1 Then%>s<%end if%></b></td>
										<td>&nbsp;</td>
									</tr>
									
								</table>
						<br>
						<br>
							
					
									 								
							<%End if%>
		<SCRIPT LANGUAGE="javascript">							
			function JsausenciaDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse registro de ausencia?"))
				  {
				  document.location.href('empresa/acoes/ausencias_acao.asp?cod='+cod+'&cod_emp='+cod2+'&acao=apaga_emprego');
					}
					  } 
		</SCRIPT>
	