<!--#include file="../../../includes/util.asp"-->
<!--#include file="../../../includes/inc_open_connection.asp"-->
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<%   
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

If strcod <> "" And strmsg = "" Then
	
xsql ="up_Funcionario_Lista_por_cd_codigo @cd_codigo="&strcod&""
Set rs = dbconn.execute(xsql)

	If NOT rs.EOF Then
		str_funcionario = rs("cd_codigo")
		strregime_trabalho = rs("cd_regime_trabalho")
		strmatricula = rs("cd_matricula") 
		strnome = rs("nm_nome")
		str_sexo = rs("cd_sexo")
		strfoto = rs("nm_foto")
		strregime_contrato = rs("cd_regime_trabalho")
		
		cd_numreg = rs("cd_numreg")
		str_numreg = rs("cd_numreg")
	End If

else

End if
%>


<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<script language="JavaScript">
VerifiqueTAB=true;
function Mostra(quem, tammax) {
	if ( (quem.value.length == tammax) && (VerifiqueTAB) ) { 
	var i=0,j=0, indice=-1;
	for (i=0; i<document.forms.length; i++) { 
	for (j=0; j<document.forms[i].elements.length; j++) { 
	if (document.forms[i].elements[j].name == quem.name) { 
	indice=i;
	break;
	} 
	} 
	if (indice != -1) break; 
	} 
	for (i=0; i<=document.forms[indice].elements.length; i++) { 
	if (document.forms[indice].elements[i].name == quem.name) { 
	while ( (document.forms[indice].elements[(i+1)].type == "hidden") &&
	(i < document.forms[indice].elements.length) ) { 
	i++;
	} 
	document.forms[indice].elements[(i+1)].focus();
	VerifiqueTAB=false;
	break;
	} 
	} 
	} 
	} 

function PararTAB(quem) { VerifiqueTAB=false; } 
function ChecarTAB() { VerifiqueTAB=true; } 
</script>
<!--include file="../../../js/ajax.js"-->
<!--include file="js/cid.js"-->

				<table align="center" border="0" rules="groups" width="100%">
					<tr id="no_print">
						<td>&nbsp;&nbsp;<a href="empresa.asp?tipo=ausencia&func=ativo">Listagem</a> 
						&nbsp;&nbsp;&nbsp; <a href="empresa.asp?tipo=novo">Novo</a></td>
					</tr>
				</table>
					<!--table align="center" border="1" rules="groups">
					<tr><td-->
								<table border="0" cellspacing="0" cellpadding="2" align="center"><!-- OK 2-->
									
									<form name="Form" action="empresa/acoes/funcionarios_acao.asp"  method="POST"  onsubmit="return validar_cad(document.forms);">
									<input type="hidden" name="cd_sexo" value="<%=str_sexo%>">
									<input type="hidden" name="acao" value="insert">
									<input type="hidden" name="cod" value="<%=strcod%>">
									<input type="hidden" name="cd_funcionario" value="<%=str_funcionario%>">
									<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
										<tr>
											<td colspan="7" align="center" class="cabecalho"><b>REGISTRO DE FALTAS DE COLABORADOR</b></td>
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
											<td colspan="2"><%strsql = "Select * From TBL_tipo_contrato Where cd_codigo ="&strregime_trabalho&""
						  					Set rs_contr = dbconn.execute(strsql)%>
											<%Do While Not rs_contr.eof
											nm_regime_trab = rs_contr("nm_regime_trab")
											cd_cod= rs_contr("cd_codigo")%>
												
											<b><%=nm_regime_trab%></b>
											<%rs_contr.movenext
											loop
											rs_contr.close
											Set rs_contr = nothing%>
											</td>								
										</tr>
										<tr id="no_print"><td colspan="3"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
										<tr>
											<td rowspan="100%" id="ok_print"><img src="../imagens/px.gif" width="40" height="1"></td>
											<td rowspan="6" align="left" valign="top" id="ok_print"><img src="foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"><br></td>
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
												 <b><%=nm_unidade%></b></a>
												<%rs_unidade.close
												Set rs_unidade = nothing%>
											</td>							
										</tr>
										<tr id="no_print"><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
										<tr id="no_print"><td colspan="4"><b></b></td></tr>
										<tr>
											<td style="background-color:silver; color:white;"><b>Data da falta</b></td>
											<td style="background-color:silver; color:white;"><b>Motivo</b></td>
											<td style="background-color:silver; color:white;" colspan="2"><b>Obs.</b></td>
											
										</tr>
										<tr>
											<td valign="top"><input type="text" name="dt_dia" size="3" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(dia_sel)%>">/
															<input type="text" name="dt_mes" size="3" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(mes_sel)%>">/
															<input type="text" name="dt_ano" size="5" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="<%=ano_sel%>"></td>
											<td valign="top">
															<select name="cd_motivo_falta">
																<option value="1">Sem motivo</option>
																<option value="2">Consulta médica atestado</option>
																<option value="3">Atestado de acompanhante</option>
																<option value="0">Outros</option>
															</select>
											</td>
											<td colspan="2"><textarea cols="34" rows="3" name="obs_ausencia"></textarea></td>
										</tr>
										<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
										
										
										<tr>
											<td><img src="../imagens/px.gif" width="125" height="1"></td>
											<td><img src="../imagens/px.gif" width="130" height="1"></td>
											<td><img src="../imagens/px.gif" width="125" height="1"></td>
											<td><img src="../imagens/px.gif" width="115" height="1"></td>
										</tr>
										<tr>
											<td style="background-color:silver; color:white;"><b>CID</b></td>
											<td style="background-color:silver; color:white;" colspan="3"><b>Período de afastamento</b></td>
										</tr>
										<tr id="no_print">
											<td valign="top">
															<input type="text" name="cd_cid" size="15" maxlength="20" onKeyup="mostra_cid();">
											</td>
											<td valign="middle">
														De: <input type="text" name="per_dia_i" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">/
															<input type="text" name="per_mes_i" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">/
															<input type="text" name="per_ano_i" size="4" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="">&nbsp;
															<input type="text" name="per_hor_i" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">:
															<input type="text" name="per_min_i" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">
															</td>
											<td colspan="2" valign="middle">
														até: <input type="text" name="per_dia_f" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">/
															<input type="text" name="per_mes_f" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">/
															<input type="text" name="per_ano_f" size="4" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="">&nbsp;
															<input type="text" name="per_hor_f" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">:
															<input type="text" name="per_min_f" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">
															</td>				
										</tr>
										<tr id="no_print">
											<td colspan="4"><!--span id='divCid'></span--></td>								
										</tr>							
																	
										</form>
									</table>
								<!--/td>
							</tr-->
						</table>
						<br>
						<br>
							
					
									 								
							<%'End if%>