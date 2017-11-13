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
str_funcionario = rs("cd_codigo")
strregime_trabalho = rs("cd_regime_trabalho")
strmatricula = rs("cd_matricula") 
strnome = rs("nm_nome")
strnm_sobrenome = rs("nm_sobrenome")
strfoto = rs("nm_foto")
Strdt_contratado = rs("dt_contratado")
strdt_nascimento = rs("dt_nascimento")
strnm_rg = rs("nm_rg")
strnm_cpf = rs("nm_cpf")
strnm_tit_eleitor = rs("nm_tit_eleitor")
strnm_reservista = rs("nm_reservista")
strmae = rs("nm_mae")
strpai = rs("nm_pai")
strnm_endereco = rs("nm_endereco")
strnr_numero = rs("nr_numero")
strnm_complemento = rs("nm_complemento")
strnm_cidade = rs("nm_cidade")
strnm_estado = rs("nm_estado")
strnm_cep = rs("nm_cep")
strregime_contrato = rs("cd_regime_trabalho")
str_desliga = rs("dt_desliga")
strbairro = rs("nm_bairro")
str_nacionalidade = rs("nm_nacionalidade")
str_pis = rs("cd_pis")
str_ctps = rs("cd_ctps")
str_ctps_serie = rs("cd_ctps_serie")
str_estado_civil = rs("cd_estado_civil")
str_assistencia_medica_matricula = rs("cd_assistencia_medica_matricula")
str_assistencia_odonto_matricula = rs("cd_assistencia_odonto_matricula")
str_vr = rs("cd_vr")
str_vale_alimentacao = rs("cd_vale_alimentacao")
str_cmt_bom = rs("cd_cmt_bom")
str_sptrans = rs("cd_sptrans")

str_banco = rs("nm_banco")
str_banco_ag = rs("cd_banco_ag")
str_banco_conta = rs("cd_banco_conta")

cd_numreg = rs("cd_numreg")

str_conspro = rs("cd_conspro")
str_numreg = rs("cd_numreg")
str_rgpro_concessao = rs("rgpro_concessao")
str_rgpro_status = rs("cd_rgpro_status")
str_rgpro = rs("obs_rgpro")
str_dt_rgproinscr = rs("dt_rgproinscr")
str_dt_rgpropag = rs("dt_rgpropag")
str_dt_rgproval = rs("dt_rgproval")

str_endereco_alt = rs("nm_endereco_alt")

str_sexo = rs("cd_sexo")


End If
else

End if
%>


<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<script language="JavaScript">
function validar_cad(shipinfo){
	if (shipinfo.cd_regime_trabalho.value.length==""){window.alert ("O item Contrato deve ser preenchido.");shipinfo.cd_regime_trabalho.focus();return (false);}
	if (shipinfo.nm_nome.value.length==""){window.alert ("O nome deve ser preenchido.");shipinfo.nm_nome.focus();return (false);}
	if (shipinfo.dt_nascimento.value.length==""){window.alert ("A data de nascimento deve ser preenchida.");shipinfo.dt_nascimento.focus();return (false);}
	//if (shipinfo.dt_contratado.value.length==""){window.alert ("A data admissão deve ser preenchida.");shipinfo.dt_contratado.focus();return (false);}
	//if (shipinfo.dt_contratado.value.length==""){window.alert ("A data admissão deve ser preenchida.");shipinfo.dt_contratado.focus();return (false);}
	//if (shipinfo.dt_contratado.value.length==""){window.alert ("A data admissão deve ser preenchida.");shipinfo.dt_contratado.focus();return (false);}
	if (shipinfo.cd_cargo.value.length==""){window.alert ("O cargo deve ser preenchido.");shipinfo.cd_cargo.focus();return (false);}
	if (shipinfo.cd_status.value.length==""){window.alert ("Informe a situação do funcionário.");shipinfo.cd_status.focus();return (false);}
	if (shipinfo.cd_unidade.value.length==""){window.alert ("Informe a unidade do funcionário.");shipinfo.cd_unidade.focus();return (false);}
	if (shipinfo.hr_entrada.value.length==""){window.alert ("Informe o horário de entrada do funcionário.");shipinfo.hr_entrada.focus();return (false);}
	if (shipinfo.nm_intervalo.value.length==""){window.alert ("Informe o tipo de intervalo do funcionário.");shipinfo.nm_intervalo.focus();return (false);}
	if (shipinfo.hr_saida.value.length==""){window.alert ("Informe o horário de saída do funcionário.");shipinfo.hr_saida.focus();return (false);}	
	return (true);}
</script>
<!--#include file="../../js/ajax.js"-->
<script language="Javascript">
<!-- 
//-------------------------------------------------------------------
function rep_space(string) {
	while (string.indexOf('a') != -1) {
 		string = string.replace('a','%20');
	}
	return string;
}
//--------------------------------------------------------------------

	
function adiciona_dependente(a,b,c,d,e,dependentes_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	dependentes_total=dependentes_total;
	//nextfield ="cd_material_1";
		
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}
			while (b.indexOf(' ') != -1) {
	 		b = b.replace(' ','%20');}
			while (c.indexOf(' ') != -1) {
	 		c = c.replace(' ','%20');}
			while (d.indexOf(' ') != -1) {
	 		d = d.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		e = e.replace(' ','%20');}
			
			//contatos_total = contatos_total.replace(' ','%20');}
			
			while (dependentes_total.indexOf('$$') != -1) {
	 		dependentes_total = dependentes_total.replace('$$','$');}
						
		if (dependentes_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	dependentes_total = dependentes_total + separador
		if (a != ""){
			dependentes_total = dependentes_total + a + '#' + c + '#' + d + '#' + e + '#';
			document.form.dependentes_result.value=a;} //adiciona				
						
		if (b != ""){
			dependentes_total = dependentes_total.replace(b,'');
			document.form.dependentes_total.value=dependentes_total;} //subtrai			

	//document.dependente.dependentes_result.value=a; //a
	//document.dependentecontato.nm_nome_contato.value=b; //b
	document.form.dependentes_total.value=(dependentes_total.replace("$$", "$"));	
	document.form.dependentes_total.value=(dependentes_total); //c
	
	document.form.nm_dependente_nome.value="";
	document.form.cd_dependente_parentesco.value="";
	document.form.dt_dependente_nascimento.value="";
	document.form.cd_dependente_obs.value="";
	document.form.nm_dependente_nome.focus();
	//chama Ajax	
	{
		var oHTTPRequest_dependentes = createXMLHTTP(); 
		oHTTPRequest_dependentes.open("post", "empresa/funcionario/dependentes/js/ajax/ajax_dependentes_lista.asp", true);
		oHTTPRequest_dependentes.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_dependentes.onreadystatechange=function(){
		if (oHTTPRequest_dependentes.readyState==4){
	    document.all.divDependente_lista.innerHTML = oHTTPRequest_dependentes.responseText;}}
	    oHTTPRequest_dependentes.send("dependentes_total=" + form.dependentes_total.value);
	}
 
}
//---------------------------------------------------------------------->
</script>

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
						<td>&nbsp;&nbsp;Listagem: <a href="empresa.asp?tipo=lista&func=ativo">Listagem</a>
						&nbsp;&nbsp;&nbsp; <a href="empresa.asp?tipo=novo">Novo</a></td>
					</tr></table>
					<table align="center" border="1" rules="groups">
					<tr><td>
					<table border="0" cellspacing="0" cellpadding="2" align="center"><!-- OK 2-->
						
						<form name="form" action="empresa/acoes/funcionarios_acao.asp"  method="POST"  onsubmit="return validar_cad(document.forms);">
						<%If strcod = "" then%>
						<input type="hidden" name="acao" value="insert">
						<%ElseIf strcod <> ""  then %>
						<input type="hidden" name="acao" value="altera">
						<input type="hidden" name="cod" value="<%=strcod%>">
						<input type="hidden" name="cd_funcionario" value="<%=str_funcionario%>">
						<input name="foto_h" type="hidden" value="<%=strfoto%>">
						<%End if%>
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
							<tr>
								<td colspan="7" align="center" class="cabecalho"><b>FICHA CADASTRAL DO COLABORADOR <%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr><td colspan="100%"><img src="../imagens/blackdot.gif" width="100%" height="1"><br>&nbsp;</td></tr>
							<tr id="no_print">
								<td rowspan="100%"><img src="../imagens/px.gif" width="35" height="1"></td>
								<td rowspan="7" align="left" valign="top" style="color:#7d7d7d;">									
										<img src="foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"><br>
									<%if strcod <> "" then%>
										<a href="#" onClick="window.open('empresa/janelas/funcionarios_foto.asp?cd_codigo=<%=strcod%>','upload','width=400,height=150')" alt="Inserir/mudar foto">(Alterar foto)</a>
									<%end if%>
								</td>
								<%'if 
								'end if%>
								<td id="no_print"><b>Matrícula</b></td>
								<td id="no_print">&nbsp;</td>
								<td>&nbsp;</td>
								<td rowspan="100%"><img src="../imagens/px.gif" width="35" height="1"></td>
							</tr>
							<tr id="no_print">
								<td><input type="text" name="cd_matricula" value="<%=strmatricula%>" size="10" maxlength="4" class="inputs"></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr id="no_print"><td colspan="3"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td rowspan="100%" id="ok_print"><img src="../imagens/px.gif" width="40" height="1"></td>
								<td rowspan="6" align="left" valign="top" id="ok_print"><img src="foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"><br></td>
								<td colspan="2" style="color:red;"><b>Nome</b></td>
								<td style="color:red;"><b>Sexo</b></td>
							</tr>
							<tr>
								<td colspan="2"><input type="text" name="nm_nome" value="<%=strnome & strnm_sobrenome%>" size="50" maxlength="70" class="inputs"></td>
								<td>
									<select name="cd_sexo">
											<option value=""></option>
											<option value="1" <%if str_sexo = 1 then response.write("SELECTED")%>>Masculino</option>
											<option value="2" <%if str_sexo = 2 then response.write("SELECTED")%>>Feminino</option>
									</select>								
								</td>
							</tr>
							<tr>
								<td colspan="3"><b>Endereço / Nº</b></td>								
							</tr>
							<tr>
								<td colspan="3"><input type="text" name="nm_endereco" value="<%=strnm_endereco%>" size="65" maxlength="100" class="inputs">&nbsp;
								<input type="text" name="nr_numero" value="<%=strnr_numero%>" size="4" maxlength="8" class="inputs"></td>
							</tr>
							<tr>
								<td><b>Bairro</b></td>
								<td colspan="3"><b>Complemento</b></td>
							</tr>
							<tr>
								<td><input type="text" name="nm_bairro" value="<%=strbairro%>" size="20" maxlength="100" class="inputs"></td>
								<td colspan="3"><input type="text" name="nm_complemento" value="<%=strnm_complemento%>" size="50" maxlength="30" class="inputs"></td>
							</tr>
							<tr>
								<td id="ok_print">&nbsp;</td>
								<td><b>CEP</b></td>
								<td colspan="2"><b>Cidade</b></td>
								<td><b>UF</b></td>
							</tr>
							<tr>
								<td id="ok_print">&nbsp;</td>
								<td><input type="text" name="nm_cep" value="<%=strnm_cep%>" size="10" maxlength="9" class="inputs"></td>
								<td colspan="2"><input type="text" name="nm_cidade" value="<%=strnm_cidade%>" size="30" maxlength="30" class="inputs"></td>
								<td><input type="text" name="nm_estado" value="<%=strnm_estado%>" size="1" maxlength="2" class="inputs"></td>
							</tr>
							<tr>
								<td colspan="4"><b>Filiação</b></td>
							</tr>
							<tr>
								<td colspan="4"><input type="text" name="nm_mae" value="<%=strmae%>" size="86" maxlength="100" class="inputs"> <b>Mãe</b></td>
							</tr>
							<tr>
								<td colspan="4"><input type="text" name="nm_pai" value="<%=strpai%>" size="86" maxlength="100" class="inputs"> <b>Pai</b></td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td><b>Endereço alternativo</b></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td colspan="4"><input type="text" name="nm_endereco_alt" value="<%=str_endereco_alt%>" size="100 maxlength="150" class="inputs"></td>
							</tr>
							<tr>
								<td colspan="100%">
								<!-- ****** CONTATOS!!!! ****** -->
									<%'id_origem = str_funcionario
									'cd_origem = 4
									'pag_retorno = ".."&mem_posicao
									'variaveis = "../../empresa.asp?tipo=cadastro$cod="&str_funcionario%>
									<!--include file="../../contatos/default.asp"-->
								<!-- ****** CONTATOS!!!! ****** -->							
								</td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr><td colspan="4"><b>Contatos</b></td></tr>
							<tr>
								<!--td><b>Nome</b></td-->
								<td><b>Tipo</b></td>
								<td><b>Tel./Cel/E-mail/Nextel e outros</b></td>
								<td><b>Observações</b></td>
							</tr>
							<tr>
								<input type="hidden" name="cd_contato_origem" value="4">
								<input type="hidden" name="id_contato_origem" value="<%=str_funcionario%>">
								<!--td><input type="text" name="nm_contato_nome" value="<%=str_contato_nome%>" size="20" maxlength="100" class="inputs"></td-->
								<td><!--input type="text" name="nm_cargo" value="<%'=%>" size="20" maxlength="100" class="inputs"-->
								
								<select name="cd_contato_categoria" class="inputs">
									<option value=""></option>
									<%strsql = "Select * From TBL_contato_cat "'Where cd_codigo=1 or cd_codigo = 2"
									Set rs_contato = dbconn.execute(strsql)%>
									<%Do While Not rs_contato.eof
											categoria_contato = rs_contato("categoria_contato")
											cd_contato_cat = rs_contato("cd_codigo")
											'if cd_contato_cat = cat_ then%><%'ckcontr = "selected"%><%'Else%><%'ckcontr=""%><%'end if%>	
											<option value="<%=cd_contato_cat%>" <%=ckcontr%>><%=categoria_contato%></option>
											<%rs_contato.movenext
										loop
											rs_contato.close
											Set rs_contato = nothing%>						
											
								</select></td>
								<td><input type="text" name="cd_contato_ddd" value="<%=str_contato_ddd%>" size="2" maxlength="2" class="inputs">&nbsp;
									<input type="text" name="nm_contato_numero" value="<%=str_contato%>" size="20" maxlength="100" class="inputs"></td>
								<td><input type="text" name="nm_contato_obs" value="<%=str_contato_obs%>" size="25" maxlength="100" class="inputs"></td>
							</tr>
							<tr>
								<td colspan="100%">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_contato Where cd_origem=4 and id_origem='"&str_funcionario&"'"
											Set rs_contato = dbconn.execute(strsql)
											
											'if not rs_contato.EOF then
											while not rs_contato.EOF
											nm_contato_nome = rs_contato("nm_nome")
											cd_contato_categoria = rs_contato("cd_categoria")
											'nm_cargo = rs_contato("nm_contato")
											cd_contato_ddd = rs_contato("cd_ddd")
											nm_contato_numero = rs_contato("nm_numero")
											nm_contato_obs = rs_contato("nm_obs")%>	
										<%if cabeca_contato = 0 Then%>
										<tr bgcolor="#808080" style="color:white;">
											<!--td>Nome</td-->
											<td>&nbsp;<b>Tipo</b></td>
											<td>&nbsp;<b>Tel/Cel/Email</b></td>
											<td>&nbsp;<b>Observações</b></td>
										</tr>
										<%end if%>
										<tr>
											<!--td>&nbsp;<%'=nm_contato_nome%></td-->
											
											<%strsql = "Select * From TBL_contato_cat Where cd_codigo='"&cd_contato_categoria&"'"'Where cd_codigo=1 or cd_codigo = 2"
											Set rs_contato_cat = dbconn.execute(strsql)%>
											<%Do While Not rs_contato_cat.eof
												categoria_contato = rs_contato_cat("categoria_contato")
												cd_contato_cat = rs_contato_cat("cd_codigo")
											rs_contato_cat.movenext
											loop
											rs_contato_cat.close
											Set rs_contato_cat = nothing%>
											
											
											<td>&nbsp;<%=categoria_contato%></td>
											<td>&nbsp;<%if cd_contato_ddd <> "" Then response.write(cd_contato_ddd &" - ")%> <%=nm_contato_numero%></td>
											<td>&nbsp;<%=nm_contato_obs%></td>
										</tr>
										<%cabeca_contato = cabeca_contato + 1
										rs_contato.movenext
										wend%>									
									</table>
								</td>
							</tr>							
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td style="color:red;"><b>Data de nascimento</b></td>
								<td><b>Nacionalidade</b></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td><input type="text" name="dt_nascimento" value="<%=strdt_nascimento%>" size="20" maxlength="10" class="inputs">
								<!--input type="text" name="dt_nascimento_day" value="<%=zero(day(strdt_nascimento))%>" size="3" maxlength="2" class="inputs">/
								<input type="text" name="dt_nascimento_month" value="<%=zero(month(strdt_nascimento))%>" size="3" maxlength="2" class="inputs">/
								<input type="text" name="dt_nascimento_year" value="<%=year(strdt_nascimento)%>" size="6" maxlength="4" class="inputs"></td-->
								<td><input type="text" name="nm_nacionalidade" value="<%=str_nacionalidade%>" size="12" maxlength="25" class="inputs"></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr>								
								<td><b>PIS</b></td>
								<td><b>Carteira Profissional</b></td>
								<td><b>Série</b></td>								
								<td>&nbsp;</td>
							</tr>
							<tr>								
								<td><input type="text" name="cd_pis" value="<%=str_pis%>" size="20" maxlength="25" class="inputs"></td>
								<td><input type="text" name="cd_ctps" value="<%=str_ctps%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_ctps_serie" size="5" class="inputs" maxlength="5" value="<%=str_ctps_serie%>"></td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td><b>RG</b></td>
								<td><b>CPF</b></td>
								<td><b>Tit. Eleitor</b></td>
								<td><b>Reservista</b></td>
							</tr>
							<tr>
								<td><input type="text" name="nm_rg" value="<%=strnm_rg%>" size="20" maxlength="15" class="inputs"></td>
								<td><input type="text" name="nm_cpf" value="<%=strnm_cpf%>" size="20" maxlength="14" class="inputs"></td>
								<td><input type="text" name="nm_tit_eleitor" value="<%=strnm_tit_eleitor%>" size="25" maxlength="14" class="inputs"></td>
								<td><input type="text" name="nm_reservista" value="<%=strnm_reservista%>" size="25" maxlength="14" class="inputs"></td>
							</tr>
							<!--*****************-->
							<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
							<tr>
								<td><b>Dependentes</b></td>
							</tr>
							<tr><td>Nome</td><td>Parentesco</td><td>Data nasc.</td><td>Observações.</td><td>&nbsp;</td></tr>
							<!--form name="formdependente"-->
							<tr>
								<td><input type="text" name="nm_dependente_nome" size="20" maxlength="30" class="inputs"></td>
								<td><select name="cd_dependente_parentesco" class="inputs">
										<option value=""></option>
										<%strsql = "Select * From TBL_dependente_parentesco "'Where cd_codigo=1 or cd_codigo = 2"
										Set rs_parentesco = dbconn.execute(strsql)%>
										<%Do While Not rs_parentesco.eof
												nm_parentesco = rs_parentesco("nm_parentesco")
												cd_dependente_parentesco = rs_parentesco("cd_codigo")
												'if cd_dependente_parentesco = cat_ then%><%'ckparent = "selected"%><%'Else%><%'ckparent=""%><%'end if%>	
												<option value="<%=cd_dependente_parentesco%>" <%=ckparent%>><%=nm_parentesco%></option>
												<%rs_parentesco.movenext
											loop
												rs_parentesco.close
												Set rs_parentesco = nothing%>						
												
									</select></td>
								<td><input type="text" name="dt_dependente_nascimento" value="" size="10" maxlength="10"></td>
								<td><input type="text" name="cd_dependente_obs" value="" size="15"></td>
								<td>
								<!--input type="button" name="inclui_dependente" value="+" onFocus="adiciona_dependente(document.forms.nm_dependente_nome.value,'',document.form.cd_dependente_parentesco.value,document.form.dt_dependente_nascimento.value,document.form.cd_dependente_obs.value,document.form.dependentes_total.value)"-->
								<input type="button" name="inclui_dependente" value="+" onFocus="adiciona_dependente(document.form.nm_dependente_nome.value,'',document.form.cd_dependente_parentesco.value,document.form.dt_dependente_nascimento.value,document.form.cd_dependente_obs.value,document.form.dependentes_total.value)"></td>
							</tr>
							<!--tr><td colspan="100%"><br-->
							<input type="hidden" name="dependentes_total" value="" size="100"><br>	
							<input type="hidden" name="dependentes_result" size="70"><!--/td></tr-->	
							<!--/form-->
							<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
							<tr><td colspan="100%"><span id='divDependente_lista'> &nbsp;</span></td></tr>
							<!--*****************-->
							
							
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr><td><b>DEPENDENTES</b></td></tr>
							<tr>
								<td>&nbsp;<b>Nome</b></td>
								<td>&nbsp;<b>Parentesco</b></td>
								<td>&nbsp;<b>Data de nasc.</b></td>								
							</tr>
							<tr>
								<td colspan="100%">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_dependente Where cd_funcionario='"&str_funcionario&"' order by dt_nascimento"
											Set rs_dependente = dbconn.execute(strsql)
											
											'if not rs_contato.EOF then
											conta_dependentes = 1
											while not rs_dependente.EOF
											cd_dependente = rs_dependente("cd_codigo")
											nm_dependente_nome = rs_dependente("nm_nome")
											cd_dependente_parentesco = rs_dependente("cd_parentesco")
											dt_dependente_nascimento = rs_dependente("dt_nascimento")
											nm_dependente_obs = rs_dependente("nm_obs")%>
										<%if cabeca_dependente = "" Then%>
										<tr bgcolor="#808080" style="color:white;">
											<td><b>&nbsp;</b></td>
											<td>&nbsp;<b>Nome</b>&nbsp;</td>
											<td>&nbsp;<b>Parentesco</b>&nbsp;</td>
											<td>&nbsp;<b>Nascimento</b>&nbsp;</td>
											<td>&nbsp;<b>Idade</b>&nbsp;</td>
											<td>&nbsp;<b>Observações</b>&nbsp;</td>
										</tr>
										<%end if%>
										<tr>
											<td><b><%=conta_dependentes%></b></td>
											<td>&nbsp;<%=nm_dependente_nome%></td>											
											<%strsql = "Select * From TBL_dependente_parentesco Where cd_codigo="&cd_dependente_parentesco&""
											Set rs_dependente_parent = dbconn.execute(strsql)%>
											<%While Not rs_dependente_parent.eof
												nm_dependente_parentesco = rs_dependente_parent("nm_parentesco")
												'cd_dependente_parent = rs_dependente_parent("cd_codigo")
											rs_dependente_parent.movenext
											wend
											'rs_dependente_parent.close
											'Set rs_dependente_parent = nothing%>
											
											<td>&nbsp;<%=nm_dependente_parentesco%></td>
											<td>&nbsp;<%=dt_dependente_nascimento%></td>
										<%'idade = datediff("m",day(dt_dependente_nascimento)&"/"&month(dt_dependente_nascimento)&"/"&year(dt_dependente_nascimento),now)
										idade = datediff("m",dt_dependente_nascimento,now)
											if idade >= 12 then
												if idade = 12 then
													tempo = "ano"
												else
													tempo = "anos"
												end if
												idade = int(idade/12)&" "&tempo
											else 
												if idade <= 1 then
													tempo = "mes"
												else
													tempo = "meses"
												end if
												idade = idade&" "&tempo
											end if%>
											<td>&nbsp;<%=idade%></td>
											<td>&nbsp;<%=nm_dependente_obs%></td>
											<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="javascript:JsDeleteDependente('<%=cd_dependente%>','<%=str_funcionario%>')"><%=str_funcionario%></td>
										</tr>
										<%cabeca_dependente = cabeca_dependente + 1
										conta_dependentes = conta_dependentes + 1
										rs_dependente.movenext
										wend%>									
									</table>
								</td></tr>	
									
								
								
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td><b>Registro Profissional</b></td>
								<td><b>Número</b></td>
								<td><b>Tipo de concessão / Situação</b></td>
								<td><b>Observações</b></td>
							</tr>
							<tr>
								<td>
								<select name="cd_conspro" class="inputs">
										<option value=""></option>
								<%strsql = "Select * From TBL_cons_profissional "
			  					Set rs_conspro = dbconn.execute(strsql)
									While not rs_conspro.EOF
										cod_conspro = rs_conspro("cd_codigo")
										nm_conspro = rs_conspro("nm_conspro")
											if cod_conspro = int(str_conspro) then
												conspro_sel = "Selected"
											else
												conspro_sel = ""
											end if%>
																
										<option value="<%=cod_conspro%>" <%=conspro_sel%>><%=nm_conspro%></option>
										<%rs_conspro.movenext
									wend%>
								</select></td>
								<td><input type="text" name="cd_numreg" value="<%=str_numreg%>" size="20" maxlength="25" class="inputs"></td>
								<td valign="top">
									<select name="rgpro_concessao" class="inputs">
											<%if str_rgpro_concessao = "" then
												concessao_sel_0 =	"SELECTED"
											elseif str_rgpro_concessao = 1 then
												concessao_sel_1 =	"SELECTED"
											elseif str_rgpro_concessao = 2 then
												concessao_sel_2 =	"SELECTED"
											end if%>
											<option value="" <%=concessao_sel_0%>></option>
											<option value="1" <%=concessao_sel_1%>>Definitivo</option>
											<option value="2" <%=concessao_sel_2%>>Provisório</option>
									</select>&nbsp;									
									<select name="rgpro_status" class="inputs">
											<%if str_rgpro_status = "" then
												rgpro_status_sel_0 =	"SELECTED"
											elseif str_rgpro_status = 1 then
												rgpro_status_sel_1 =	"SELECTED"
											elseif str_rgpro_status = 2 then
												rgpro_status_sel_2 =	"SELECTED"
											end if%>
											<option value="" <%=rgpro_status_sel_0%>></option>
											<option value="1" <%=rgpro_status_sel_1%>>Ativo</option>
											<option value="2" <%=rgpro_status_sel_2%>>Inativo</option>
									</select></td>
								<td valign="top" rowspan="3"><textarea cols="25" rows="4" name="obs_rgpro" class="inputs"><%=str_rgpro%></textarea></td>
							</tr>
							<tr>
								<td><b>Data de inscrição</b></td>
								<td><b>Data de pagamento</b></td>
								<td><b>Validade</b></td>
								<td></td>
							</tr>
							<tr>
								<td valign="top"><input type="text" name="dt_rgproinscr" size="20" maxlength="10" class="inputs" value="<%=str_dt_rgproinscr%>"></td>
								<td valign="top"><input type="text" name="dt_rgpropag" size="20" maxlength="10" class="inputs" value="<%=str_dt_rgpropag%>"></td>
								<td valign="top"><input type="text" name="dt_rgproval" size="25" maxlength="12" class="inputs" value="<%=str_dt_rgproval%>"></td>
								<!--td><input type="text" name="dia_valid" size="3" maxlength="2" class="inputs">/<input type="text" name="mes_valid" size="3" maxlength="2" class="inputs">/<input type="text" name="ano_valid" size="5" maxlength="4" class="inputs"></td-->
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td><b>Cartão SPTRANS</b></td>
								<td><b>Cartão BOM</b></td>
								<td><b>Vale Refeição</b></td>
								<td><b>Cesta Básica</b></td>
							</tr>
							<tr>
								<td><input type="text" name="cd_sptrans" value="<%=str_sptrans%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_cmt_bom" value="<%=str_cmt_bom%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_vr" value="<%=str_vr%>" size="25" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_vale_alimentacao" value="<%=str_vale_alimentacao%>" size="25" maxlength="20" class="inputs"></td>
							</tr>
							<tr>
								<td><b>Banco</b></td>
								<td><b>Agência</b></td>
								<td><b>Conta</b></td>
								<td>&nbsp;</td>
							</tr>
							<tr><%if str_banco = "" Then%><%str_banco = "Real"%><%end if%>
								<%if str_banco_ag = "" Then%><%str_banco_ag = "1342"%><%end if%>
								<td><input type="text" name="nm_banco" value="<%=str_banco%>" size="20" maxlength="30" class="inputs"></td>
								<td><input type="text" name="cd_banco_ag" value="<%=str_banco_ag%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_banco_conta" value="<%=str_banco_conta%>" size="20" maxlength="20" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td>&nbsp;<b>Assistência Médica</b></td>
								<td>&nbsp;<b>Assistência Odontológica</b></td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td><input type="text" name="cd_assistencia_medica_matricula" value="<%=str_assistencia_medica_matricula%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="cd_assistencia_odonto_matricula" value="<%=str_assistencia_odonto_matricula%>" size="20" maxlength="20" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<%if strcod = "" Then%>
							<tr>
								<td><b>Contrato</b> &nbsp;</td>
								<td style="color:red;"><b>Data de admissão</b></td>
								<%If strcod <> "" Then%>
								<td><b>Data de desligamento</b></td><%else%><td>&nbsp;</td><%End If%>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td><%strsql = "Select * From TBL_tipo_contrato "'Where cd_codigo=1 or cd_codigo = 2"
			  					Set rs_contr = dbconn.execute(strsql)%>
								<select name="cd_regime_trabalho" class="inputs">
								<option value="">Selecione</option>
								<%Do While Not rs_contr.eof
								nm_regime_trab = rs_contr("nm_regime_trab")
								cd_cod = rs_contr("cd_codigo")
								if cd_cod = strregime_trabalho then%><%ckcontr = "selected"%><%Else%><%ckcontr=""%><%end if%>	
								<option value="<%=cd_cod%>" <%=ckcontr%>><%=cd_cod%>-<%=nm_regime_trab%></option>
								<%rs_contr.movenext
								loop
								rs_contr.close
								Set rs_contr = nothing%>
								</select> </td>
								<td><!--input type="text" name="dt_contratado" value="<%=Strdt_contratado%>" size="20" maxlength="10" class="inputs"-->
									<input type="text" name="cd_dia" value="" size="3" maxlength="2" class="inputs">/
									<input type="text" name="cd_mes" value="" size="3" maxlength="2" class="inputs">/
									<input type="text" name="cd_ano" value="" size="6" maxlength="4" class="inputs"></td>								
								<td>&nbsp;</td>								
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<%else%>						
							<tr>
								<td><b>Contrato</b> &nbsp;</td>
								<td style="color:red;"><b>Data de admissão</b></td>
								<%If strcod <> "" Then%>
								<td><b>Data de desligamento</b></td><%else%><td>&nbsp;</td><%End If%>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							
								<%'strsql = "Select * From VIEW_funcionario_contrato_lista Where cd_funcionario = "&str_funcionario&" order by dt_admissao desc"
								strsql = "Select * From VIEW_funcionario_contrato_lista Where cd_funcionario = "&str_funcionario&" order by dt_admissao desc"
			  					Set rs_contrato = dbconn.execute(strsql)%>
								<%while not rs_contrato.EOF
								cd_funcionario = rs_contrato("cd_funcionario")
								cd_regime_trabalho = rs_contrato("cd_codigo")
								nm_regime_trab = rs_contrato("nm_regime_trab")
								dt_admissao = rs_contrato("dt_admissao")
								dt_demissao = rs_contrato("dt_demissao")
								%>
								<tr>
									<td>&nbsp;<%=nm_regime_trab%></td>
									<td>&nbsp;<%=zero(day(dt_admissao))%>/<%=mesdoano(month(dt_admissao))%>/<%=year(dt_admissao)%></td>
									<td>&nbsp;<%=dt_demissao%></td>
									<td>&nbsp;&nbsp;
									<%if dt_demissao <> "" then
										if nr_contrato = "" then%>
											<a href="javascript:;" onclick="window.open('empresa/janelas/funcionarios_contrato.asp?cd_funcionario=<%=str_funcionario%>', 'janela_vdlap', 'width=350, height=400'); return false;"><b>*** Novo Contrato ***</b></a>
										<%end if%>
									<%else%>
											<a href="javascript:;" onclick="window.open('empresa/janelas/funcionarios_recisao.asp?cd_contrato=<%=cd_regime_trabalho%>','principal','width=350, height=400'); return false;">Dispensar?</a>&nbsp;
									<%end if%>
									<a href="javascript:;" onclick="window.open('empresa/janelas/funcionarios_contrato.asp?cd_funcionario=<%=str_funcionario%>', 'janela_vdlap', 'width=350, height=400'); return false;">Alterar</a></td>
								</tr>
								<%nr_contrato = nr_contrato + 1
								rs_contrato.movenext
								wend%>
							
							<%end if%>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td colspan="2" style="color:red;"><b>Cargo</b></td>
								<td style="color:red;"><b>Situação</b></td>
								<td style="color:red;"><b>Unidade</b></td>
							</tr>
							<tr>
								<%strsql = "TBL_cargos where cd_especificidade = 3"
								Set rs_func = dbconn.execute(strsql)%>
								<%if strcod = "" Then%>
								<td colspan="2">
									<select name="cd_cargo" class="inputs"> 
									<option value="">Selecione</option>
									<%Do While Not rs_func.eof%>
									<option value="<%=rs_func("cd_codigo")%>"<%if strcd_funcao = CInt(rs_func("cd_codigo")) then%>selected<%End if%>><%=rs_func("nm_cargo")%></option>
									<%rs_func.movenext
									loop
									rs_func.close
									Set rs_func = nothing%>
									</select> *
								<%Else
									xsql_cargo = "select * From View_funcionario_cargo Where cd_codigo='"&strcod&"'"
									Set rs_cargo = dbconn.execute(xsql_cargo)
									if not rs_cargo.EOF Then
									nm_cargo = rs_cargo("nm_cargo")
								End if%>
								
								<%strsql = "select * from VIEW_funcionario_cargo where cd_funcionario='"&strcod&"' and cd_suspenso <> 1 order by dt_atualizacao desc"
								Set rs_cargo = dbconn.execute(strsql)%>
								<%if not rs_cargo.eof then%><%nm_cargo=rs_cargo("nm_cargo")%><%end if%>
								<td class="txt_cinza" colspan="2" <%=tipo%> style="background: #c7c7c7;"><a href="javascript:;" onclick="window.open('empresa/janelas/funcionarios_cargo.asp?cd_funcionario=<%=strcod%>&cd_funcao=<%=cd_funcao%>', 'janela_vdlap', 'width=350, height=400'); return false;"><%=nm_cargo%>.</a>
																
								<%end if%></td>								
								<td style="background: #c7c7c7;">
								<%if strcod = "" Then%>
									<%strsql ="up_ADM_Status_lista"
									Set rs_cat = dbconn.execute(strsql)%>
								
									<select name="cd_status" class="inputs"> 
									<option value="">Selecione</option>
									<%Do While Not rs_cat.eof
									cd_cod_status = rs_cat("cd_codigo")
									%>
									<option value="<%=cd_cod_status%>"><%=cd_cod_status%>-<%=rs_cat("nm_status")%></option>
									<%rs_cat.movenext
									loop%>
									</select>
								<%else%>
								<%strsql ="SELECT * From View_funcionario_status where cd_funcionario = "&strcod&"  and cd_suspenso <> 1 order by dt_atualizacao"
								  	Set rs_status = dbconn.execute(strsql)
									
									do while not rs_status.EOF
									cd_status = rs_status("cd_status")
									nm_status = rs_status("nm_status")
									dt_atualizacao = rs_status("dt_atualizacao")
									
									rs_status.movenext
									loop%>
									<a href="javascript:;" onclick="window.open('empresa/janelas/funcionarios_status.asp?cd_funcionario=<%=strcod%>&cd_status=<%=cd_status%>', 'janela_vdlap', 'width=350, height=400'); return false;"><%=nm_status%>.</a>
								<%end if%>
									</td>
									<td style="background: #c7c7c7;">
									<%strsql ="TBL_unidades"
								  	Set rs_uni = dbconn.execute(strsql)%>
								<%if strcod = "" Then%>
									<select name="cd_unidade" class="inputs"> 
									<option value="">Selecione</option>
									<%Do While Not rs_uni.eof%>
									<option value="<%=rs_uni("cd_codigo")%>"><%=rs_uni("nm_Unidade")%></option>
									<%rs_uni.movenext
									loop
									'rs_uni.close
									'Set rs_uni = nothing%>
									</select> 	
								<%else%>
									<%strsql ="SELECT * From View_funcionario_unidade where cd_funcionario = "&strcod&" and cd_suspenso <> 1 order by dt_atualizacao asc"
								  	Set rs_unidade = dbconn.execute(strsql)
									
									do while not rs_unidade.EOF
									cd_unidade = rs_unidade("cd_unidade")
									nm_unidade = rs_unidade("nm_unidade")
									dt_atualizacao = rs_unidade("dt_atualizacao")
									
									
									rs_unidade.movenext
									loop%>
									 <a href="javascript:;" onclick="window.open('empresa/janelas/funcionarios_unidade.asp?cd_funcionario=<%=strcod%>&cd_unidade=<%=cd_unidade%>', 'janela_vdlap', 'width=350, height=400'); return false;"><%=nm_unidade%>.</a>
									<%rs_uni.close
									Set rs_uni = nothing%>
								<%end if%></td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td><b>Entrada</b></td>
								<td><b>Intervalo</b></td>
								<td><b>Saída</b></td>
								<td>&nbsp;</td>
							</tr>
							<%xsql = "Select * From TBL_funcionario_horario Where cd_funcionario = '"&strcod&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
							'xsql = "SELECT  * FROM         TBL_funcionario_horario WHERE     (cd_funcionario = 97) ORDER BY dt_atualizacao DESC"
							Set rs_hora = dbconn.execute(xsql)
								if not rs_hora.EOF then
									hr_entrada = rs_hora("hr_entrada")                                
									hr_saida = rs_hora("hr_saida")
									nm_intervalo = rs_hora("nm_intervalo")
								end if
							  %>														
							<tr>
								<td>
									<%if strcod = "" Then%>
										<input type="text" name="hr_entrada" size="6" maxlength="5" class="inputs">
									<%else%>
										<%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%>
									<%end if%>
								</td>
								<td>
									<%if strcod = "" Then%>
										<select name="nm_intervalo" size="1" class="inputs">
											<option value=""></option>
											<option value="1 Hora">1 Hora</option>
											<option value="15 minutos">15 minutos</option>
										</select>
									<%else%>
										<%=nm_intervalo%>
									<%end if%>
								</td>
								<td>
									<%if strcod = "" Then%>
										<input type="text" name="hr_saida" size="6" maxlength="5"  class="inputs">
									<%else%>
										<%=zero(hour(hr_saida))&":"&zero(minute(hr_saida))%>
										&nbsp;<a href="javascript:;" onclick="window.open('empresa/janelas/funcionarios_horario.asp?cd_funcionario=<%=strcod%>', 'janela_vdlap', 'width=350, height=430'); return false;"><img src="../../imagens/ic_novo.gif" alt="" width="10" height="12" border="0"></a></td>
									<%end if%>
								<td>&nbsp;</td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr><td colspan="100%">
								<%id_origem = str_funcionario
								cd_origem = 4
								pag_retorno = ".."&mem_posicao
								variaveis = "../../empresa.asp?tipo=cadastro$cod="&str_funcionario%>
								<!--#include file="../../ocorrencia/ocorrencia_formulario.asp"-->
							</td></tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td><img src="../imagens/px.gif" width="125" height="1"></td>
								<td><img src="../imagens/px.gif" width="130" height="1"></td>
								<td><img src="../imagens/px.gif" width="125" height="1"></td>
								<td><img src="../imagens/px.gif" width="115" height="1"></td>
							</tr>
							<tr id="no_print">
								<td colspan="4" valign="middle"><A href="empresa.asp?tipo=lista"><img src="imagens/bt_lista.gif" alt="Listar" border="0" width="119" height="22" border="1"></a>
								<input type="image" src="imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0">
								</td>
							</tr>
							<%if session_cd_usuario = 46 then%>
							<tr id="no_print"><td><br><br></td></tr>
							<tr id="no_print"><td><input onclick="javascript:JsDelete('<%=strcod%>')" type="button" value="Apagar" title="Apagar">
								</td>
							</tr>
							<%end if%>
																				
							</form>
							<%If strcod <> "" then%>
							
							<tr><td><form name="form_ex" action="empresa/acoes/empr_funcionarios_acao.asp" id="forms" method="post"  enctype="multipart/form-data">
							<input type="hidden" name="acao" value="excluir">
							<input type="hidden" name="cod" value="<%=cd_funcionario%>">
							</form>
							</td></tr>
							<%End if%>
							<tr><td>&nbsp;</td></tr>
							
						</table>
						</td></tr>
						</table>
						<br>
						<br>
							
					
									 								
							<%End if%>
			<SCRIPT LANGUAGE="javascript">
 {
        //MaskInput(document.forms.nm_cep, "99999-99");
		//MaskInput(document.forms.dt_nascimento, "99/99/9999");	    
 }
</SCRIPT>
			
			<SCRIPT LANGUAGE="javascript">
				function JsDelete(cod)
				   {
					if (confirm ("Tem certeza que deseja excluir?"))
				  {
				  document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir');
					}
					  }
				
				function JsDeleteDependente(cod_dependente,cod_funcionario)
				   {
					if (confirm ("Tem certeza que deseja o excluir o dependente?"))
				  {
				  document.location.href('empresa/acoes/funcionarios_acao.asp?cd_funcionario='+cod_funcionario+'&cod='+cod_dependente+'&acao=dependente_excluir');
					}
					  }  
				//function JsContatoDelete(cod){
					//if (confirm ("Tem certeza que deseja excluir o contato?")){
				  //document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir');
					//}
					  //}
			</SCRIPT>