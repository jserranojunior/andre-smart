<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>

<%   
strcod = request("cod")
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")




If strcod <> "" And strmsg = "" Then
	
xsql ="up_Funcionario_Lista_por_cd_codigo @cd_codigo="&strcod&""
Set rs = dbconn.execute(xsql)

If NOT rs.EOF Then
strregime_trabalho = rs("cd_regime_trabalho")
strmatricula = rs("cd_matricula") 
strnome = rs("nm_nome")
strnm_sobrenome = rs("nm_sobrenome")
strfoto = rs("nm_foto")
Strdt_contratado = rs("dt_contratado")
strdt_nascimento = rs("dt_nascimento")
strnm_email = rs("nm_email")
strnm_rg = rs("nm_rg")
strnm_cpf = rs("nm_cpf")
strmae = rs("nm_mae")
strpai = rs("nm_pai")
strnm_ddd = rs("nm_ddd")
strnm_fone = rs("nm_fone")
strnm_ddd_cel = rs("nm_ddd_cel")
strnm_cel = rs("nm_cel")
strocontato = rs("nm_o_contato")
strnm_endereco = rs("nm_endereco")
strnr_numero = rs("nr_numero")
strnm_complemento = rs("nm_complemento")
strnm_cidade = rs("nm_cidade")
strnm_estado = rs("nm_estado")
strnm_cep = rs("nm_cep")
strregime_contrato = rs("cd_regime_trabalho")
strcd_funcao = rs("cd_funcao")
strstatus = rs("cd_status")
strunidades = rs("cd_unidades")
'str_demissao = rs("dt_demissao")
str_desliga = rs("dt_desliga")
strbairro = rs("nm_bairro")
str_nacionalidade = rs("nm_nacionalidade")
str_escolaridade = rs("nm_escolaridade")
str_pis = rs("cd_pis")
str_ctps = rs("cd_ctps")
str_ctps_serie = rs("cd_ctps_serie")
strqtd_dependentes = rs("cd_qtd_dependentes")
strqtd_filhos = rs("cd_qtd_filhos")
str_estado_civil = rs("cd_estado_civil")
str_assistencia_medica = rs("cd_assistencia_medica")
str_assistencia_medica_matricula = rs("cd_assistencia_medica_matricula")
str_assistencia_medica_tipo = rs("cd_assistencia_medica_tipo")
str_carga_horaria = rs("cd_carga_horaria")
str_carga_diaria = rs("cd_carga_diaria")
str_vr = rs("cd_vr")
str_cmt = rs("cd_cmt")
str_sptrans = rs("cd_sptrans")

str_banco = rs("nm_banco")
str_banco_ag = rs("cd_banco_ag")
str_banco_conta = rs("cd_banco_conta")
str_banco_conta_tipo = rs("cd_banco_conta_tipo")

'strcarga_horaria = rs("cd_carga_horaria")

	'If Str_demissao = "300012" Then
	'	Str_demissao = ""
	'Elseif str_demissao <> "" Then
	'str_demissao_ano = left(str_demissao,4)
	'str_demissao_mes = right(str_demissao,2)
	'str_demissao_dia = 
	'str_demissao = str_demissao_mes &"/"& str_demissao_ano 
	'end if
End If
else
'response.redirect "../../empr_cadastro_lista.asp?msg="&strmsg&""
End if
%>


<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<script language="JavaScript">
function validar_cad(shipinfo){
	if (shipinfo.cd_regime_trabalho.value.length==""){window.alert ("O item Contrato deve ser preenchido.");shipinfo.cd_regime_trabalho.focus();return (false);
	} else {
	}return (true);}
</script>	
			<table width="100%" border="1" cellspacing="0" cellpadding="0">
				<tr>
					<td height="30"><img src="px.gif" alt="" width="1" height="30" border="0"></td>
				</tr>   
				<tr>		
					<td class="txt_cinza">
					<% If strmsg = "" then%>
					&nbsp; <b>Lista &raquo;  - <A href="coop_cooperados_lista.asp">Retorna</b></a><br><br>
					<%End if%>
					&nbsp; <b><%If strcod = "" then%>Cadastro<%else%>Altera<%End if%> &raquo; - <font color="red">Funcion�rio</font></b></td>
				</tr>
				<tr>
					
					<td height="30"><img src="px.gif" alt="" width="1" height="30" border="0"></td>
				</tr>
				<tr>
					<td valign="top" align=center>
					<% If strmsg <> "" then%>
						<table  border="0" cellspacing="0" cellpadding="0" class="txt_cinza">
							<tr>
							 <td>
								<b><%=strmsg%></b>
							 </td>
							</tr>
							<tr>
							 <td align=center><br><br><a href="coop_cooperados_lista.asp"><img src="imagens/bt_lista.gif" alt="" width="119" height="22" border="1"></a></td>
							</tr>
						</table>
					<%else%>
			
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>			
								<td valign="top" align=center>
							
									<table width="100%" border="0" cellspacing="1" cellpadding="1">
									<form name="form" action="empresa/funcionarios/empr_funcionarios_insert.asp"  method="POST"  onsubmit="return validar_cad(document.form);">
									<!--form name="frmSend" method="POST" enctype="multipart/form-data" action="upload/uploadTester.asp" onSubmit="return onSubmitForm();"-->
									<%If strcod = "" then%>
									  <input type="hidden" name="acao" value="insert">
									<%ElseIf strcod <> ""  then %>
			  						  <input type="hidden" name="acao" value="altera">
									  <input type="hidden" name="cod" value="<%=strcod%>">
									<%End if%>
										<%if strcod <> "" Then%>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza" valign="top">Foto :</td>
											<td><%if strfoto <> "" then%><img src="foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0">
												<%else%><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0">
												<%end if%><br>
											<font color="#7d7d7d"><a href="#" onClick="window.open('upload/formulario.asp?cd_codigo=<%=strcod%>','upload','width=400,height=150')" alt="Inserir/mudar foto">(Adicionar foto)</a></font>
												<input name="foto_h" type="hidden" value="<%=strfoto%>">
												<!--input name="foto" type="file" value=""> (<a href="../foto/funcionarios/<%=strfoto%>" target="_blank"><%=strfoto%></a>)--> 
											</td>
										</tr>
										<%end if%>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Contrato:</td>
											<td>
											<%	strsql = "Select * From TBL_tipo_contrato Where cd_codigo=1 or cd_codigo = 2"
										  Set rs_contr = dbconn.execute(strsql)%>
											<select name="cd_regime_trabalho" class="inputs">
												<option value="">Selecione</option>
											<%Do While Not rs_contr.eof
											nm_regime_trab = rs_contr("nm_regime_trab")
											cd_cod= rs_contr("cd_codigo")
											
											if cd_cod = strregime_trabalho then%><%ckcontr = "selected"%><%Else%><%ckcontr=""%><%end if%>	
												<option value="<%=rs_contr("cd_codigo")%>" <%=ckcontr%>><%=nm_regime_trab%></option>
											<%rs_contr.movenext
												loop
			
												rs_contr.close
												Set rs_contr = nothing
											  %>
											</select> *
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Matricula:</td>
											<td><input type="text" name="cd_matricula" value="<%=strmatricula%>" size="10" maxlength="4" class="inputs"></td>
										</tr>
										<tr><td>&nbsp;<!--salto de linha--></td></tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Nome:</td>
											<td><input type="text" name="nm_nome" value="<%=strnome%>" size="55" maxlength="50" class="inputs"> *</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Sobrenome</td>
											<td><input type="text" name="nm_sobrenome" value="<%=strnm_sobrenome%>" size="55" maxlength="50" class="inputs"> *</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Endere�o:</td>
											<td><input type="text" name="nm_endereco" value="<%=strnm_endereco%>" size="50" maxlength="100" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">N�mero:</td>
											<td><input type="text" name="nr_numero" value="<%=strnr_numero%>" size="5" maxlength="8" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Complemento:</td>
											<td><input type="text" name="nm_complemento" value="<%=strnm_complemento%>" size="30" maxlength="30" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Bairro:</td>
											<td><input type="text" name="nm_bairro" value="<%=strbairro%>" size="30" maxlength="100" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Cep:</td>
											<td><input type="text" name="nm_cep" value="<%=strnm_cep%>" size="10" maxlength="9" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Cidade:</td>
											<td><input type="text" name="nm_cidade" value="<%=strnm_cidade%>" size="30" maxlength="30" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Estado:</td>
											<td><input type="text" name="nm_estado" value="<%=strnm_estado%>" size="1" maxlength="2" class="inputs"></td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Fone:</td>
											<td><input type="text" name="nm_ddd" value="<%=strnm_ddd%>" size="2" maxlength="2" class="inputs">-<input type="text" name="nm_fone" value="<%=strnm_fone%>" size="10" maxlength="9" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Celular:</td>
											<td><input type="text" name="nm_ddd_cel" value="<%=strnm_ddd_cel%>" size="2" maxlength="2" class="inputs">-<input type="text" name="nm_cel" value="<%=strnm_cel%>" size="10" maxlength="9" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Outros contatos:</td>
											<td><input type="text" name="nm_o_contato" value="<%=strocontato%>" size="55" maxlength="100" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Email</td>
											<td><input type="text" name="nm_email" value="<%=strnm_email%>" size="55" maxlength="50" class="inputs"></td>
										</tr>
										
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">M�e:</td>
											<td><input type="text" name="nm_mae" value="<%=strmae%>" size="55" maxlength="100" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Pai:</td>
											<td><input type="text" name="nm_pai" value="<%=strpai%>" size="55" maxlength="100" class="inputs"></td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Data de Nascimento</td>
											<td><input type="text" name="dt_nascimento" value="<%=strdt_nascimento%>" size="12" maxlength="12" class="inputs"> *</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Nacionalidade</td>
											<td><input type="text" name="nm_nacionalidade" value="<%=str_nacionalidade%>" size="12" maxlength="25" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Escolaridade</td>
											<%if str_escolaridade = 1 then%><%ck_escola1="selected"%>
											<%elseif str_escolaridade = 2 then%><%ck_escola2="selected"%>
											<%elseif str_escolaridade = 3 then%><%ck_escola3="selected"%>
											<%elseif str_escolaridade = 4 then%><%ck_escola4="selected"%>
											<%elseif str_escolaridade = 5 then%><%ck_escola5="selected"%>
											<%elseif str_escolaridade = 6 then%><%ck_escola6="selected"%>
											<%elseif str_escolaridade = 7 then%><%ck_escola7="selected"%>
											<%elseif str_escolaridade = 8 then%><%ck_escola8="selected"%>
											<%elseif str_escolaridade = 9 then%><%ck_escola9="selected"%>
											<%else%><%ck_escola0="selected"%>
											<%end if%>
											<td><select name="nm_escolaridade" class="inputs">
													<option value="" <%=ck_escola0%>></option>
													<option value="1" <%=ck_escola1%>>1� Grau Cursando</option>
													<option value="2" <%=ck_escola2%>>1� Grau completo</option>
													<option value="3" <%=ck_escola3%>>2� Grau cursando</option>
													<option value="4" <%=ck_escola4%>>2� Grau completo</option>
													<option value="5" <%=ck_escola5%>>3� Grau cursando</option>
													<option value="6" <%=ck_escola6%>>3� Grau completo</option>
													<option value="7" <%=ck_escola7%>>Especializa��o</option>
													<option value="8" <%=ck_escola8%>>Mestrado</option>
													<option value="9" <%=ck_escola9%>>Doutorado</option>
												</select>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">PIS/PASEP</td>
											<td><input type="text" name="cd_pis" value="<%=str_pis%>" size="15" maxlength="25" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">RG</td>
											<td><input type="text" name="nm_rg" value="<%=strnm_rg%>" size="15" maxlength="13" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">CPF</td>
											<td><input type="text" name="nm_cpf" value="<%=strnm_cpf%>" size="15" maxlength="14" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Estado Civil</td>
											<td><select name="cd_estado_civil" class="inputs">
												<%if str_estado_civil = 1 then%><%ck_estciv1="selected"%>
												<%elseif str_estado_civil = 2 then%><%ck_estciv2="selected"%>
												<%elseif str_estado_civil = 3 then%><%ck_estciv3="selected"%>
												<%elseif str_estado_civil = 4 then%><%ck_estciv4="selected"%>
												<%elseif str_estado_civil = 5 then%><%ck_estciv5="selected"%>
												<%elseif str_estado_civil = 6 then%><%ck_estciv6="selected"%>
												<%else%><%ck_estciv0="selected"%><%end if%>
													<option value="" <%=ck_estciv%>>Selecione</option>
													<option value="1" <%=ck_estciv1%>>Solteiro(a)</option>
													<option value="2" <%=ck_estciv2%>>Casado(a)</option>
													<option value="3" <%=ck_estciv3%>>Separado(a)</option>
													<option value="4" <%=ck_estciv4%>>Divorciado(a)</option>
													<option value="5" <%=ck_estciv5%>>Vi�vo(a)</option>
													<option value="6" <%=ck_estciv6%>>Uni�o est�vel</option>
												</select>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Carteira Profissional</td>
											<td  class="txt_cinza"><input type="text" name="cd_ctps" value="<%=str_ctps%>" size="20" maxlength="20" class="inputs"> S�rie <input type="text" name="cd_ctps_serie" size="5" class="inputs" maxlength="5" value="<%=str_ctps_serie%>"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">N� de dependentes</td>
											<td  class="txt_cinza"><input type="text" name="cd_qtd_dependentes" value="<%=strqtd_dependentes%>" size="5" maxlength="5" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">N� de filhos</td>
											<td  class="txt_cinza"><input type="text" name="cd_qtd_filhos" value="<%=strqtd_filhos%>" size="5" maxlength="5" class="inputs"></td>
										</tr>
										<tr>
											<td colspan="2" align="right">-------------------</td>
											<td colspan="2" align="left">-------------------</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Assistencia M�dica</td>
										<%	strsql = "TBL_convenio order by nm_convenio"
										  Set rs_conv = dbconn.execute(strsql)%><td>
											<select name="cd_assistencia_medica" class="inputs"> 
											<option value="">Selecione</option>
											<%Do While Not rs_conv.eof
											cd_conv = rs_conv("cd_codigo")
											nm_convenio = rs_conv("nm_convenio")
											%>
											<option value="<%=cd_conv%>" <%if str_assistencia_medica = cd_conv then%>selected<%End if%>><%=nm_convenio%></option>
											  <%
												rs_conv.movenext
												loop
												rs_conv.close
												Set rs_conv = nothing
											  %>
												</select>								
											</td>
											<!--td><input type="text" name="cd_funcao" value="<%=strcd_funcao%>" size="15" maxlength="50" class="inputs"></td-->
										</tr>
										
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Matricula</td>
											<td><input type="text" name="cd_assistencia_medica_matricula" value="<%=str_assistencia_medica_matricula%>" size="20" maxlength="20" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Modalidade</td>
											<%if str_assistencia_medica_tipo = 1 then%><%ck_assist_m1 = "checked"%>
											<%Elseif str_assistencia_medica_tipo = 2 then%><%ck_assist_m2 = "checked"%>
											<%end if%>
											<td><input type="radio" name="cd_assistencia_medica_tipo" value="1" <%=ck_assist_m1%>> Standard<br>
												<input type="radio" name="cd_assistencia_medica_tipo" value="2" <%=ck_assist_m2%>> Pleno</td>
										</tr
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Cart�o SPTRANS</td>
											<td><input type="text" name="cd_sptrans" value="<%=str_sptrans%>" size="20" maxlength="20" class="inputs"></td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">CMT</td>
											<td><input type="text" name="cd_cmt" value="<%=str_cmt%>" size="20" maxlength="20" class="inputs"></td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Vale refei��o</td>
											<td><input type="text" name="cd_vr" value="<%=str_vr%>" size="20" maxlength="20" class="inputs"></td>
										</tr>
										<tr>
											<td colspan="2" align="right">-------------------</td>
											<td colspan="2" align="left">-------------------</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Banco<%=str_banco%></td>
											<td><input type="text" name="nm_banco" value="<%=str_banco%>" size="20" maxlength="30" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Agencia</td>
											<td><input type="text" name="cd_banco_ag" value="<%=str_banco_ag%>" size="20" maxlength="20" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Conta</td>
											<td><input type="text" name="cd_banco_conta" value="<%=str_banco_conta%>" size="20" maxlength="20" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td align="right" class="txt_cinza" class="txt_cinza">tipo - Conta</td>
											<%if str_banco_conta_tipo = 1 Then%><%ck_banco_tipo1 = "checked"%>
											<%elseif str_banco_conta_tipo = 2 Then%><%ck_banco_tipo2 = "checked"%><%end if%>
											<td class="txt_cinza"><input type="radio" name="cd_banco_conta_tipo" value="1" <%=ck_banco_tipo1%>>C/C<br>
											<input type="radio" name="cd_banco_conta_tipo" value="2" <%=ck_banco_tipo2%>>C/P</td>
										</tr>
										<tr>
											<td colspan="2" align="right">-------------------</td>
											<td colspan="2" align="left">-------------------</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Carga hor�ria</td>
											<td><input type="text" name="cd_carga_horaria" value="<%=str_carga_horaria%>" size="5" maxlength="5" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Carga di�ria</td>
											<td><input type="text" name="cd_carga_diaria" value="<%=str_carga_diaria%>" size="5" maxlength="5" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Data de Contrata��o</td>
											<td><input type="text" name="dt_contratado" value="<%=Strdt_contratado%>" size="12" maxlength="12" class="inputs"> *</td>
										</tr>
										<%If strcod <> ""    Then%>
										<!--tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Mes/Ano de demiss&atilde;o</td>
											<td><input type="text" name="dt_demissao" value="<%=str_demissao%>" size="12" maxlength="12" class="inputs"></td>
										</tr-->
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Data de desligamento</td>
											<td><input type="text" name="dt_desliga" value="<%=str_desliga%>" size="12" maxlength="12" class="inputs"></td>
										</tr>
										<%End If%>
										
										
										
										<tr><td>&nbsp;<!--salto de linha --></td></tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Funcao:</td>
										<%	strsql = "TBL_funcoes"
										  Set rs_func = dbconn.execute(strsql)%><td>
											<select name="cd_funcao" class="inputs"> 
											<option value="">Selecione</option>
											<%Do While Not rs_func.eof%>
											<option value="<%=rs_func("cd_codigo")%>"<%if strcd_funcao = CInt(rs_func("cd_codigo")) then%>selected<%End if%>><%=rs_func("nm_funcao")%></option>
											  <%
												rs_func.movenext
												loop
			
												rs_func.close
												Set rs_func = nothing
											  %>
												</select> *
											</td>
											<!--td><input type="text" name="cd_funcao" value="<%=strcd_funcao%>" size="15" maxlength="50" class="inputs"></td-->
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td align="right" class="txt_cinza">Status:</td>
										<%
										  strsql ="up_ADM_Status_lista"
										  Set rs_cat = dbconn.execute(strsql)
										%>
											<td><select name="cd_status" class="inputs"> 
											<option value="">Selecione</option>
											  <%	
												Do While Not rs_cat.eof%>
											  %>
												<option value="<%=rs_cat("cd_codigo")%>" <%if strstatus = CInt(rs_cat("cd_codigo")) then%>selected<%End if%>  ><%=rs_cat("nm_status")%></option>
											  <%
												rs_cat.movenext
												  loop
												rs_cat.close
												Set rs_cat = nothing
											  %>
												</select> *
											</td>
										</tr>
									
										
			
										
										<tr>
											<td>&nbsp;</td>
											<td align="right" class="txt_cinza">Unidades:</td>
										<%
										  strsql ="up_ADM_unidades_lista"
										  Set rs_uni = dbconn.execute(strsql)
										%>
											
											<td><select name="cd_unidades" class="inputs"> 
											<option value="">Selecione</option>
											  <%	
												Do While Not rs_uni.eof%>
											  %>
												<option value="<%=rs_uni("cd_codigo")%>" <%if strunidades = CInt(rs_uni("cd_codigo")) then%>selected<%End if%>  ><%=rs_uni("nm_Unidade")%></option>
											  <%
												rs_uni.movenext
												  loop
												rs_uni.close
												Set rs_uni = nothing
											  %>
												</select> *
											</td>
										</tr>
																
								<!--/td-->
								</tr>
									<td>&nbsp;</td>
									<td  align="right" colspan=2>
											<table width=100%>
											  <tr>
											   <td  align="right"><A href="empresa_funcionarios_lista.asp?pag=<%=strpag%>&tipo=<%=strtipo%>&busca=<%=strbusca%>"><img src="imagens/bt_lista.gif" alt="Listar" width="119" height="22" border="1"></a></td>
												<td  align="left"><input type="image" src="imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0"></td>
												<td align="left"><input onclick="javascript:window.open('empresa/funcionarios/coop_cadastro_impressao.asp?cod=<%=strcod%>','Impressao','width=700,height=185','location=1');" type="button" value="Imprimir" title="Imprimir">
												<input onclick="javascript:JsDelete('<%=strcod%>')" type="button" value="Apagar" title="Apagar"></td>
												</form>
												<%If strcod <> "" then%>
												<form name="form" action="/empresa/funcionarios/empr_funcionarios_Insert.asp" id="forms" method="post"  enctype="multipart/form-data">
												
												<!--td  align="left"--> </td>
												<input type="hidden" name="acao" value="excluir">
												<input type="hidden" name="cod" value="<%=strcod%>">
												</form>
												<%End if%>
											  </tr>
											 </table>
											</td>
										</tr>
										       
									</table> 
									 <SCRIPT LANGUAGE="javascript">
			 {
			   	MaskInput(document.form.dt_nascimento, "99/99/9999");
			    MaskInput(document.form.dt_contratado, "99/99/9999");
				MaskInput(document.form.nm_cpf, "999.999.999-99");
			   	MaskInput(document.form.nm_rg, "99.999.999-a");
			   	MaskInput(document.form.nm_ddd, "99");
			   	MaskInput(document.form.nm_fone, "9999-9999");
			   	MaskInput(document.form.nm_ddd_cel, "99");
			   	MaskInput(document.form.nm_cel, "9999-9999");
			   	MaskInput(document.form.nm_cep, "99999-999");
				MaskInput(document.form.dt_demissao, "99/9999");
			 }
			</SCRIPT>
			
							<%End if%>
								</td>
							</tr>
							<tr>
								<td height="26"><img src="px.gif" alt="" width="1" height="26" border="0"></td>
							</tr>
							<tr class="textopequeno">
								<td colspan="100%" align="center">* - Indica campo obrigat�rio</td>
							</tr>
							<tr>
								<td height="26"><img src="px.gif" alt="" width="1" height="26" border="0"></td>
							</tr>                                                   
						</table>  
			
			
			
			<SCRIPT LANGUAGE="javascript">
			
			
				function JsDelete(cod)
				 
				  {
					if (confirm ("Tem certeza que deseja excluir?"))
				  {
				  document.location.href('empresa/funcionarios/empr_funcionarios_insert.asp?cod='+cod+'&acao=excluir');
					}
					  }
			
				
			</SCRIPT>
			
			