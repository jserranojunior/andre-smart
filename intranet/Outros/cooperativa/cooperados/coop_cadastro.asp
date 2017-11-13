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
strcd_funcao = rs("cd_funcao")
strstatus = rs("cd_status")
strunidades = rs("cd_unidades")
str_demissao = rs("dt_demissao")

	If Str_demissao = "300012" Then
		Str_demissao = ""
	Elseif str_demissao <> "" Then
	str_demissao_ano = left(str_demissao,4)
	str_demissao_mes = right(str_demissao,2)
	str_demissao = str_demissao_mes &"/"& str_demissao_ano 
	end if
End If
%>


<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<script language="JavaScript">
function validar_cad(shipinfo){
	if (shipinfo.nm_nome.value.length==""){window.alert ("O item Nome deve ser preenchido.");shipinfo.nm_nome.focus();return (false);
	} else {
	if (shipinfo.nm_sobrenome.value.length==""){window.alert ("O item Sobrenome deve ser preenchido.");shipinfo.nm_sobrenome.focus();return (false);
	} else {
	if (shipinfo.dt_contratado.value.length==""){window.alert ("O item Data de Contratação deve ser preenchido.");shipinfo.dt_contratado.focus();return (false);
	} else {
	if (shipinfo.dt_nascimento.value.length==""){window.alert ("O item Data de Nascimento deve ser preenchido.");shipinfo.dt_nascimento.focus();return (false);
	} else {	
	if (shipinfo.nm_rg.value.length==""){window.alert ("O item RG deve ser preenchido.");shipinfo.nm_rg.focus();return (false);
	} else {
	if (shipinfo.nm_cpf.value.length==""){window.alert ("O item CPF deve ser preenchido.");shipinfo.nm_cpf.focus();return (false);
	} else {
	if (shipinfo.nm_ddd.value.length==""){window.alert ("O item DDD deve ser preenchido.");shipinfo.nm_ddd.focus();return (false);
	} else {
	if (shipinfo.nm_fone.value.length==""){window.alert ("O item Fone deve ser preenchido.");shipinfo.nm_fone.focus();return (false);
	} else {
	if (shipinfo.nm_endereco.value.length==""){window.alert ("O item Endereço deve ser preenchido.");shipinfo.nm_endereco.focus();return (false);
	} else {
	if (shipinfo.nr_numero.value.length==""){window.alert ("O item Núemro deve ser preenchido.");shipinfo.nr_numero.focus();return (false);
	} else {
	if (shipinfo.nm_cidade.value.length==""){window.alert ("O item Cidade deve ser preenchido.");shipinfo.nm_cidade.focus();return (false);
	} else {
	if (shipinfo.nm_estado.value.length==""){window.alert ("O item Estado deve ser preenchido.");shipinfo.nm_estado.focus();return (false);
	} else {
	if (shipinfo.nm_cep.value.length==""){window.alert ("O item CEP deve ser preenchido.");shipinfo.nm_cep.focus();return (false);
	} else {
	if (shipinfo.cd_funcao.value.length==""){window.alert ("O item Função deve ser preenchido.");shipinfo.cd_funcao.focus();return (false);
	} else {
	if (shipinfo.cd_status.value.length==""){window.alert ("O item Status deve ser preenchido.");shipinfo.cd_status.focus();return (false);
	} else {
	if (shipinfo.cd_unidades.value.length==""){window.alert ("O item Unidade deve ser preenchido.");shipinfo.cd_unidades.focus();return (false);
	} else {
	}}}}}}}}}}}}}}}}
	return (true);
	}
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
					&nbsp; <b><%If strcod = "" then%>Cadastro<%else%>Altera<%End if%> &raquo; - <font color="red">Colaborador</font></b></td>
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
									<form name="form" action="cooperativa/cooperados/funcionarios_insert.asp"  method="POST"  onsubmit="return validar_cad(document.form);">
									<!--form name="frmSend" method="POST" enctype="multipart/form-data" action="upload/uploadTester.asp" onSubmit="return onSubmitForm();"-->
									<%If strcod = "" then%>
									  <input type="hidden" name="acao" value="insert">
									<%ElseIf strcod <> ""  then %>
			  						  <input type="hidden" name="acao" value="altera">
									  <input type="hidden" name="cod" value="<%=strcod%>">
									<%End if%>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Matricula:</td>
											<td><input type="text" name="cd_matricula" value="<%=strmatricula%>" size="10" maxlength="4" class="inputs"></td>
										</tr>
										<%if strcod <> "" Then%>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza" valign="top">Foto :</td>
											<td><img src="foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"><br>
											<font color="#7d7d7d"><a href="#" onClick="window.open('upload/formulario.asp?cd_codigo=<%=strcod%>','upload','width=400,height=150')" alt="Inserir/mudar foto">(Adicionar foto)</a></font>
												<input name="foto_h" type="hidden" value="<%=strfoto%>">
												<!--input name="foto" type="file" value=""> (<a href="../foto/funcionarios/<%=strfoto%>" target="_blank"><%=strfoto%></a>)--> 
											</td>
										</tr>
										<%end if%>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Nome:</td>
											<td><input type="text" name="nm_nome" value="<%=strnome%>" size="55" maxlength="50" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Sobrenome</td>
											<td><input type="text" name="nm_sobrenome" value="<%=strnm_sobrenome%>" size="55" maxlength="50" class="inputs"></td>
										</tr>	
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Data de Contratação</td>
											<td><input type="text" name="dt_contratado" value="<%=Strdt_contratado%>" size="12" maxlength="12" class="inputs"></td>
										</tr>
									<%If strcod <> ""    Then%>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Mes/Ano de demiss&atilde;o</td>
											<td><input type="text" name="dt_demissao" value="<%=str_demissao%>" size="12" maxlength="12" class="inputs"></td>
										</tr>
									<%End If%>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Data de Nascimento</td>
											<td><input type="text" name="dt_nascimento" value="<%=strdt_nascimento%>" size="12" maxlength="12" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Email</td>
											<td><input type="text" name="nm_email" value="<%=strnm_email%>" size="55" maxlength="50" class="inputs"></td>
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
											<td  align="right" class="txt_cinza">Mãe:</td>
											<td><input type="text" name="nm_mae" value="<%=strmae%>" size="55" maxlength="100" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Pai:</td>
											<td><input type="text" name="nm_pai" value="<%=strpai%>" size="55" maxlength="100" class="inputs"></td>
										</tr>	
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
											<td  align="right" class="txt_cinza">Endereço:</td>
											<td><input type="text" name="nm_endereco" value="<%=strnm_endereco%>" size="50" maxlength="100" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Número:</td>
											<td><input type="text" name="nr_numero" value="<%=strnr_numero%>" size="5" maxlength="8" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Complemento:</td>
											<td><input type="text" name="nm_complemento" value="<%=strnm_complemento%>" size="30" maxlength="30" class="inputs"></td>
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
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Cep:</td>
											<td><input type="text" name="nm_cep" value="<%=strnm_cep%>" size="10" maxlength="9" class="inputs"></td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td  align="right" class="txt_cinza">Funcao:</td>
										<%	strsql = "TBL_funcoes"
										  Set rs_func = dbconn.execute(strsql)%><td>
											<select name="cd_funcao" class="inputs"> 
											<option value="">Selecione</option>
											<%Do While Not rs_func.eof%>
											<option value="<%=rs_func("cd_codigo")%>"<%if strcd_funcao = CInt(rs_func("cd_codigo")) then%>selected<%End if%> ><%=rs_func("nm_funcao")%></option>
											  <%
												rs_func.movenext
												loop
			
												rs_func.close
												Set rs_func = nothing
											  %>
												</select>								
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
												</select>								
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
												</select>								
											</td>
										</tr>
																
								<!--/td-->
								</tr>
									<td>&nbsp;</td>
									<td  align="right" colspan=2>
											<table width=100%>
											  <tr>
											   <td  align="right"><A href="coop_cooperados_lista.asp?pag=<%=strpag%>&tipo=<%=strtipo%>&busca=<%=strbusca%>"><img src="imagens/bt_lista.gif" alt="" width="119" height="22" border="1"></a></td>
												<td  align="left"><input type="image" src="imagens/bt.gif" alt="" width="119" height="22" border="0"></td>
												</form>
												<%If strcod <> "" then%>
												<form name="form" action="/cooperativa/cooperados/FUncionarios_Insert.asp" id="forms" method="post"  enctype="multipart/form-data">
												<input onclick="javascript:window.open('cooperativa/cooperados/coop_cadastro_impressao.asp?cod=<%=strcod%>','Impressao','width=700,height=185');" type="button" value="Imprimir" title="Apagar">
												<td  align="right"><input onclick="javascript:JsDelete('<%=strcod%>')" type="button" value="A<%=strcod%>" title="Apagar"> </td>
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
			    MaskInput(document.form.dt_contratado, "99/99/9999");
				MaskInput(document.form.dt_demissao, "99/9999");
			   	MaskInput(document.form.dt_nascimento, "99/99/9999");   
				MaskInput(document.form.nm_cpf, "999.999.999-99");
			   	MaskInput(document.form.nm_rg, "99.999.999-a");
			   	MaskInput(document.form.nm_ddd, "99");
			   	MaskInput(document.form.nm_fone, "9999-9999");
			   	MaskInput(document.form.nm_ddd_cel, "99");
			   	MaskInput(document.form.nm_cel, "9999-9999");
			   	MaskInput(document.form.nm_cep, "99999-999");
			 }
			</SCRIPT>
			
							<%End if%>
								</td>
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
				  document.location.href('cooperativa/cooperados/funcionarios_insert.asp?cd_codigo='+cod+'&acao=excluir');
					}
					  }
			
				
			</SCRIPT>
			
			