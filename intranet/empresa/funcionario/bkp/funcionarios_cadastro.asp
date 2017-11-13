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
		strmatricula = rs("cd_matricula") 
		strnome = rs("nm_nome")
		strfoto = rs("nm_foto")
		strdt_nascimento = rs("dt_nascimento")
			strdia_nascimento = zero(day(strdt_nascimento))
			strmes_nascimento = zero(month(strdt_nascimento))
			strano_nascimento = year(strdt_nascimento)
		strnm_rg = rs("nm_rg")
		strdt_rg_emissao = rs("dt_rg_emissao")
		strnm_rg_expedidor = rs("nm_rg_expedidor")
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
		strbairro = rs("nm_bairro")
		str_nacionalidade = rs("nm_nacionalidade")
		str_pis = rs("cd_pis")
		str_ctps = rs("cd_ctps")
		str_ctps_serie = rs("cd_ctps_serie")
		str_estado_civil = rs("cd_estado_civil")
		
		str_assistencia_medica_matricula = rs("cd_assistencia_medica_matricula")
		str_assistencia_medica_coparticipacao = rs("nr_assistencia_medica_coparticipacao")
		str_assistencia_odonto_matricula = rs("cd_assistencia_odonto_matricula")
		str_assistencia_odonto_coparticipacao = rs("nr_assistencia_odonto_coparticipacao")
		str_vr = rs("cd_vr")
		str_vale_alimentacao = rs("cd_vale_alimentacao")
		str_cmt_bom = rs("cd_cmt_bom")
		str_sptrans = rs("cd_sptrans")
		
		str_banco = rs("nm_banco")
		str_banco_ag = rs("cd_banco_ag")
		str_banco_conta = rs("cd_banco_conta")
		
		cd_numreg = rs("cd_numreg")
		str_estado_civil = rs("cd_estado_civil")
		str_conjuge = rs("nm_conjuge")
		str_conjuge_nasc = rs("dt_conjuge_nasc")
		str_cd_plano_saude_conj = rs("cd_plano_saude_conj")
		str_naturalidade = rs("nm_naturalidade")
		strnr_tit_eleitor_zona = rs("nr_tit_eleitor_zona")
		strnr_tit_eleitor_seccao = rs("nr_tit_eleitor_seccao")
		strnm_reservista_cat = rs("nm_reservista_cat")
		strnr_hepatite_b = rs("nr_hepatite_b")
		strnr_dupla_adulto = rs("nr_dupla_adulto")
		strdt_dupla_adulto_validade = rs("dt_dupla_adulto_validade")
		strnr_scr = rs("nr_scr")
		strdt_exame_medico = rs("dt_exame_medico")
		strnr_caract_altura = rs("nr_caract_altura")
		strnr_caract_peso = rs("nr_caract_peso")
		strnr_caract_manequim = rs("nr_caract_manequim")
		strnr_caract_calcado = rs("nr_caract_calcado")
		strnm_parente = rs("nm_parente")
		strcd_parente_parentesco = rs("cd_parente_parentesco")
		strcd_residencia_tipo = rs("cd_residencia_tipo")
		strcd_residencia_financ = rs("cd_residencia_financ")
		strcd_residencia_reside = rs("cd_residencia_reside")
		strcd_veiculo_tipo = rs("cd_veiculo_tipo")
		strcd_veiculo_financ = rs("cd_veiculo_financ")
		
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
End if
%>


<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<script language="JavaScript">
function validar_cad(shipinfo){
	if (shipinfo.nm_nome.value.length==""){window.alert ("O nome deve ser preenchido.");shipinfo.nm_nome.focus();return (false);}
	if (shipinfo.cd_sexo.value.length==""){window.alert ("O sexo deve ser preenchido.");shipinfo.cd_sexo.focus();return (false);}
	if (shipinfo.dt_nascimento.value.length==""){window.alert ("A data de nascimento deve ser preenchida.");shipinfo.dt_nascimento.focus();return (false);}
	if (shipinfo.cd_regime_trabalho.value.length==""){window.alert ("O item Contrato deve ser preenchido.");shipinfo.cd_regime_trabalho.focus();return (false);}
	if (shipinfo.cd_dia.value.length==""){window.alert ("O dia de admissão deve ser preenchido.");shipinfo.cd_dia.focus();return (false);}
	if (shipinfo.cd_mes.value.length==""){window.alert ("O mês de admissão deve ser preenchido.");shipinfo.cd_mes.focus();return (false);}
	if (shipinfo.cd_ano.value.length==""){window.alert ("O ano de admissão deve ser preenchido.");shipinfo.cd_ano.focus();return (false);}
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

	
function adiciona_dependente(a,b,c,d,e,f,dependentes_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	f=f;
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
			while (f.indexOf(' ') != -1) {
	 		f = f.replace(' ','%20');}
			
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
			dependentes_total = dependentes_total + a + '#' + c + '#' + d + '#' + e + '#' + f + '#';
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
	document.form.dt_dependente_nascimento.value="";
	document.form.cd_dependente_sexo.value="";
	document.form.cd_dependente_obs.value="";
	document.form.nm_dependente_nome.focus();
	//chama Ajax	
	{
		var oHTTPRequest_dependentes = createXMLHTTP(); 
		oHTTPRequest_dependentes.open("post", "empresa/funcionario/dependentes/ajax/ajax_dependentes_lista.asp", true);
		oHTTPRequest_dependentes.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_dependentes.onreadystatechange=function(){
		if (oHTTPRequest_dependentes.readyState==4){
	    document.all.divDependente_lista.innerHTML = oHTTPRequest_dependentes.responseText;}}
	    oHTTPRequest_dependentes.send("dependentes_total=" + form.dependentes_total.value);
	} 
}

function adiciona_contato(a,b,c,d,e,f,g,contatos_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	f=f;
	g=g;
	contatos_total=contatos_total;
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
			while (f.indexOf(' ') != -1) {
	 		f = f.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		g = g.replace(' ','%20');}			
			
			while (contatos_total.indexOf('$$') != -1) {
	 		contatos_total = contatos_total.replace('$$','$');}
						
		if (contatos_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	contatos_total = contatos_total + separador
		if (a != ""){
			contatos_total = contatos_total + a + '#' + c + '#' + d + '#' + e + '#' + f + '#' + g + '#';
			document.form.contatos_result.value=a;} //adiciona				
						
		if (b != ""){
			contatos_total = contatos_total.replace(b,'');
			document.form.contatos_total.value=contatos_total;} //subtrai			

	document.form.contatos_total.value=(contatos_total.replace("$$", "$"));	
	document.form.contatos_total.value=(contatos_total); //c	
	document.form.cd_contato_categoria.value="";
	document.form.cd_contato_ddd.value="";
	document.form.nm_contato_numero.value="";
	document.form.nm_contato_obs.value="";
	document.form.cd_contato_categoria.focus();
	//chama Ajax	
	{
		var oHTTPRequest_contatos = createXMLHTTP(); 
		oHTTPRequest_contatos.open("post", "contatos/ajax/ajax_contatos_lista.asp", true);
		oHTTPRequest_contatos.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_contatos.onreadystatechange=function(){
		if (oHTTPRequest_contatos.readyState==4){
	    document.all.divContatos_lista.innerHTML = oHTTPRequest_contatos.responseText;}}
	    oHTTPRequest_contatos.send("contatos_total=" + form.contatos_total.value);
	} 
}


function adiciona_escolaridade(a,b,c,d,e,f,g,escolaridade_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	f=f;
	g=g;
	escolaridade_total=escolaridade_total;
			
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
			while (f.indexOf(' ') != -1) {
	 		f = f.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		g = g.replace(' ','%20');}			
			
			while (escolaridade_total.indexOf('$$') != -1) {
	 		escolaridade_total = escolaridade_total.replace('$$','$');}
						
		if (escolaridade_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	escolaridade_total = escolaridade_total + separador
		if (a != ""){
			escolaridade_total = escolaridade_total + a + '#' + c + '#' + d + '#' + e + '#' + f + '#' + g + '#';
			document.form.escolaridade_result.value=a;} //adiciona				
						
		if (b != ""){
			escolaridade_total = escolaridade_total.replace(b,'');
			document.form.escolaridade_total.value=escolaridade_total;} //subtrai			

	document.form.escolaridade_total.value=(escolaridade_total.replace("$$", "$"));	
	document.form.escolaridade_total.value=(escolaridade_total); //c	
	document.form.cd_ensino_grau.value="";
	document.form.nm_instituicao.value="";
	document.form.nm_curso.value="";
	document.form.cd_ensino_andamento.value="";
	//document.form.dt_termino.value="";
	document.form.nm_obs.value="";
	document.form.cd_ensino_grau.focus();
	//chama Ajax	
	{
		var oHTTPRequest_escolaridade = createXMLHTTP(); 
		oHTTPRequest_escolaridade.open("post", "empresa/funcionario/escolaridade/ajax/ajax_escolaridade_lista.asp", true);
		oHTTPRequest_escolaridade.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_escolaridade.onreadystatechange=function(){
		if (oHTTPRequest_escolaridade.readyState==4){
	    document.all.divEscolaridade_lista.innerHTML = oHTTPRequest_escolaridade.responseText;}}
	    oHTTPRequest_escolaridade.send("escolaridade_total=" + form.escolaridade_total.value);
	} 
}


function adiciona_emprego(a,b,c,d,e,f,emprego_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	f=f;
	emprego_total=emprego_total;
			
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
			while (f.indexOf(' ') != -1) {
	 		f = f.replace(' ','%20');}
			
			while (emprego_total.indexOf('$$') != -1) {
	 		emprego_total = emprego_total.replace('$$','$');}
						
		if (emprego_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	emprego_total = emprego_total + separador
		if (a != ""){
			emprego_total = emprego_total + a + '#' + c + '#' + d + '#' + e + '#' + f + '#';
			document.form.emprego_result.value=a;} //adiciona				
						
		if (b != ""){
			emprego_total = emprego_total.replace(b,'');
			document.form.emprego_total.value=emprego_total;} //subtrai			

	document.form.emprego_total.value=(emprego_total.replace("$$", "$"));	
	document.form.emprego_total.value=(emprego_total); //c	
	document.form.nm_emprego_empresa.value="";
	document.form.nm_emprego_cargo.value="";
	document.form.dt_emprego_entrada.value="";
	document.form.dt_emprego_saida.value="";
	document.form.dt_emprego_obs.value="";
	
	document.form.nm_emprego_empresa.focus();
	//chama Ajax	
	{
		var oHTTPRequest_emprego = createXMLHTTP(); 
		oHTTPRequest_emprego.open("post", "empresa/funcionario/emprego_anterior/ajax/ajax_emprego_lista.asp", true);
		oHTTPRequest_emprego.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_emprego.onreadystatechange=function(){
		if (oHTTPRequest_emprego.readyState==4){
	    document.all.divEmprego_lista.innerHTML = oHTTPRequest_emprego.responseText;}}
	    oHTTPRequest_emprego.send("emprego_total=" + form.emprego_total.value);
	}
 
}

function adiciona_valores(a,b,c,d,e,valores_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	valores_total=valores_total;
			
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
			
			while (valores_total.indexOf('$$') != -1) {
	 		valores_total = valores_total.replace('$$','$');}
						
		if (valores_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	valores_total = valores_total + separador
		if (a != ""){
			valores_total = valores_total + a + '#' + c + '#' + d + '#' + e + '#';
			document.form.valores_result.value=a;} //adiciona				
						
		if (b != ""){
			valores_total = valores_total.replace(b,'');
			document.form.valores_total.value=valores_total;} //subtrai			

	document.form.valores_total.value=(valores_total.replace("$$", "$"));	
	document.form.valores_total.value=(valores_total); //c
	document.form.cd_valor_tipo.value="";
	document.form.nr_valor.value="";
	document.form.dt_valor_atualizacao.value="";
	document.form.nm_valor_obs.value="";
	
	document.form.cd_valor_tipo.focus();
	//chama Ajax	
	{
		var oHTTPRequest_valores = createXMLHTTP(); 
		oHTTPRequest_valores.open("post", "empresa/funcionario/valores/ajax/ajax_valores_lista.asp", true);
		oHTTPRequest_valores.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_valores.onreadystatechange=function(){
		if (oHTTPRequest_valores.readyState==4){
	    document.all.divValores_lista.innerHTML = oHTTPRequest_valores.responseText;}}
	    oHTTPRequest_valores.send("valores_total=" + form.valores_total.value);
	}
 
}


function mostra_urgencia(){
	var objeto1='hidden_mostra1';
	var objeto2='hidden_mostra2';	
		
	if (document.form.cd_contato_categoria.value==5){
			var el1=document.getElementById(objeto1);
				el1.style.display='inline';		
			var el2=document.getElementById(objeto2);
				el2.style.display='inline';
		}
	if (document.form.cd_contato_categoria.value!=5){
			var el1=document.getElementById(objeto1);
				el1.style.display='none';
			var el2=document.getElementById(objeto2);
				el2.style.display='none';
		}
		
	}
//---------------------------------------------------------------------->
</script>

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
					<table align="center" border="0" rules="groups" width="100%" class="no_print">
					<tr class="no_print">
						<td>&nbsp;&nbsp;Listagem: <a href="empresa.asp?tipo=lista&func=ativo">Listagem</a>
						&nbsp;&nbsp;&nbsp; <a href="empresa.asp?tipo=novo">Novo</a></td>
					</tr></table>
			<table align="center" border="1" rules="groups">
				<tr><td>
					<table border="0" cellspacing="0" cellpadding="2" align="center" class="no_print"><!-- OK 2-->
						
						<form name="form" action="empresa/acoes/funcionarios_acao.asp"  method="POST"  onsubmit="return validar_cad(document.form);">
						<input type="hidden" name="cd_sessao" value="<%=session_cd_usuario%>">
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
							<tr class="no_print">
								<td rowspan="110"><img src="../imagens/px.gif" width="35" height="1"></td>
								<td rowspan="7" align="left" valign="top" style="color:#7d7d7d;">									
										<img src="foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"><br>
									<%if strcod <> "" then%>
										<a href="#" onClick="window.open('empresa/janelas/funcionarios_foto.asp?cd_codigo=<%=strcod%>','upload','width=500,height=550')" alt="Inserir/mudar foto">(Alterar foto)</a>
									<%end if%>
								</td>
								<%'if 
								'end if%>
								<td class="no_print"><b>Matrícula</b></td>
								<td class="no_print">&nbsp;</td>
								<td>&nbsp;</td>
								<td rowspan="100%"><img src="../imagens/px.gif" width="35" height="1"></td>
							</tr>
							<tr class="no_print">
								<td><input type="text" name="cd_matricula" value="<%=strmatricula%>" size="10" maxlength="4" class="inputs"></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr class="no_print"><td colspan="3"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr>
								<td rowspan="100%" class="ok_print"><img src="../imagens/px.gif" width="40" height="1"></td>
								<td rowspan="6" align="left" valign="top" class="ok_print"><img src="foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"><br></td>
								<td colspan="2" style="color:red;"><b>Nome</b></td>
								<td style="color:red;"><b>Sexo</b></td>
							</tr>
							<tr>
								<td colspan="2"><input type="text" name="nm_nome" value="<%=strnome %>" size="50" maxlength="70" class="inputs"></td>
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
								<td class="ok_print">&nbsp;</td>
								<td><b>CEP</b></td>
								<td colspan="2"><b>Cidade</b></td>
								<td><b>UF</b></td>
							</tr>
							<tr>
								<td class="ok_print">&nbsp;</td>
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
								<td><b>Estado Civil</b></td>
								<td colspan="2"><b>Nome do Cônjuge</b> <img src="../imagens/px.gif" width="165" height="1"> <b>Nasc.</b></td>
								<td><b>Opta plan. saúde?</b></td>
							</tr>
							<tr>								
								<td>
									<select name="cd_estado_civil" class="inputs">
									<option value=""></option>
									<%strsql = "Select * From TBL_estado_civil order by cd_codigo"
									Set rs_est_civil = dbconn.execute(strsql)%>
									<%Do While Not rs_est_civil.eof
											cd_estado_civil = rs_est_civil("cd_codigo")
											nm_estado_civil = rs_est_civil("nm_estado_civil")%>	
											<option value="<%=cd_estado_civil%>" <%if int(str_estado_civil) = cd_estado_civil then%>Selected<%end if%>><%=nm_estado_civil%></option>
											<%rs_est_civil.movenext
										loop
											rs_est_civil.close
											Set rs_est_civil = nothing%>											
									</select>
								</td>
								<td colspan="2"><input type="text" name="nm_conjuge" value="<%=str_conjuge%>" size="40" maxlength="100" class="inputs">
								<input type="text" name="dt_conjuge_nasc" value="<%=str_conjuge_nasc%>" size="12" maxlength="10" class="inputs"></td>
								<td><input type="checkbox" name="cd_plano_saude_conj" value="1" <%if str_cd_plano_saude_conj = 1 then%>checked<%end if%>  class="inputs">Sim</td>
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
							<tr><td><!--include file="../../contatos/default.asp"--></td></tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr><td colspan="4"><b>Contatos</b></td></tr>
							<tr>
								<!--td><b>Nome</b></td-->
								<td><b>Tipo</b></td>
								<td id="hidden_mostra1" style="display:none;"><b>Nome &nbsp; &nbsp; / &nbsp; &nbsp; Grau relac.</b></td>
								<td><b>Tel./Cel/E-mail/Nextel e outros</b></td>
								<td><b>Observações</b></td>
							</tr>
							<tr>
								<input type="hidden" name="cd_contato_origem" value="4">
								<input type="hidden" name="id_contato_origem" value="<%=str_funcionario%>">
								<td>								
								<select name="cd_contato_categoria" class="inputs" onchange="mostra_urgencia();">
									<option value=""></option>
									<%strsql = "Select * From TBL_contato_cat "'Where cd_codigo=1 or cd_codigo = 2"
									Set rs_contato = dbconn.execute(strsql)%>
									<%Do While Not rs_contato.eof
											categoria_contato = rs_contato("categoria_contato")
											cd_contato_cat = rs_contato("cd_codigo")%>	
											<option value="<%=cd_contato_cat%>" <%=ckcontr%>><%=categoria_contato%></option>
											<%rs_contato.movenext
										loop
											rs_contato.close
											Set rs_contato = nothing%>											
								</select></td>
								<td id="hidden_mostra2" style="display:none;">
									<input type="text" name="nm_contato_nome" value="<%=str_contato_nome%>" size="15" maxlength="50" class="inputs">&nbsp;
									<input type="text" name="nm_contato_cargo" value="<%=str_contato_cargo%>" size="15" maxlength="50" class="inputs">&nbsp;</td>
								<td><input type="text" name="cd_contato_ddd" value="<%=str_contato_ddd%>" size="2" maxlength="2" class="inputs">&nbsp;
									<input type="text" name="nm_contato_numero" value="<%=str_contato%>" size="20" maxlength="100" class="inputs"></td>
								<td><input type="text" name="nm_contato_obs" value="<%=str_contato_obs%>" size="25" maxlength="100" class="inputs">
								<input type="button" name="inclui_contato" value="+" onFocus="adiciona_contato(document.form.cd_contato_categoria.value,'',document.form.nm_contato_nome.value,document.form.nm_contato_cargo.value,document.form.cd_contato_ddd.value,document.form.nm_contato_numero.value,document.form.nm_contato_obs.value,document.form.contatos_total.value)"></td>
							</tr>
							<input type="hidden" name="contatos_total" value="" size="100"><br>	
							<input type="hidden" name="contatos_result" size="70">
							<tr><td colspan="4"><span id='divContatos_lista'> &nbsp;</span></td></tr>
							<tr>
								<td colspan="4">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_contato Where cd_origem=4 and id_origem='"&str_funcionario&"' order by cd_categoria"
											Set rs_contato = dbconn.execute(strsql)
											
											'if not rs_contato.EOF then
											while not rs_contato.EOF
												cd_contato = rs_contato("cd_codigo")
												nm_contato_nome = rs_contato("nm_nome")
												cd_contato_categoria = rs_contato("cd_categoria")
												nm_cargo = rs_contato("nm_cargo")
												cd_contato_ddd = rs_contato("cd_ddd")
												nm_contato_numero = rs_contato("nm_numero")
												nm_contato_obs = rs_contato("nm_obs")	
													if int(cd_contato_categoria) = 5 then
														cabeca_contato = 0
													end if%>
										
												<%if cabeca_contato = 0 Then%>
												<tr bgcolor="#808080" style="color:white;">													
													<td>&nbsp;</td>
													<%if cd_contato_categoria = 5 then%>
													<td>&nbsp;Nome</td>
													<td>&nbsp;Grau relac.</td>
													<%end if%>
													<td>&nbsp;<b>Tel/Cel/Email</b></td>
													<td>&nbsp;<b>Observações</b></td>
												</tr>
												<%end if%>
												
												<tr>
													<%strsql = "Select * From TBL_contato_cat Where cd_codigo='"&cd_contato_categoria&"'"'Where cd_codigo=1 or cd_codigo = 2"
													Set rs_contato_cat = dbconn.execute(strsql)%>
													<%Do While Not rs_contato_cat.eof
														categoria_contato = rs_contato_cat("categoria_contato")
														cd_contato_cat = rs_contato_cat("cd_codigo")
													rs_contato_cat.movenext
													loop
													rs_contato_cat.close
													Set rs_contato_cat = nothing%>													
													<td <%if cd_contato_categoria = 5 then%> style="color:red; font-weight:bold;"<%end if%>>&nbsp;<%=categoria_contato%></td>
													<%if cd_contato_categoria = 5 then%>
													<td>&nbsp;<%=nm_contato_nome%></td>
													<td>&nbsp;<%=nm_cargo%></td>
													<%end if%>
													<td>&nbsp;<%if cd_contato_ddd <> "" Then response.write(cd_contato_ddd &" - ")%> <%=nm_contato_numero%></td>
													<td>&nbsp;<%=nm_contato_obs%></td>
													<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsContatoDelete('<%=str_funcionario%>','<%=cd_contato%>')"></td>
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
								<td><b>Naturalidade (Cidade)</b></td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td><input type="text" name="dia_nascimento" value="<%=strdia_nascimento%>" size="2" maxlength="2" class="inputs">/
									<input type="text" name="mes_nascimento" value="<%=strmes_nascimento%>" size="2" maxlength="2" class="inputs">/
									<input type="text" name="ano_nascimento" value="<%=strano_nascimento%>" size="4" maxlength="4" class="inputs">
								<td><input type="text" name="nm_nacionalidade" value="<%=str_nacionalidade%>" size="12" maxlength="25" class="inputs"></td>
								<td><input type="text" name="nm_naturalidade" value="<%=str_naturalidade%>" size="20" maxlength="30" class="inputs"></td>
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
								<td><b>Data emissão &nbsp; - &nbsp; Órgão exped.</b></td>
								<td><b>CPF</b></td>
							</tr>
							<tr>
								<td><input type="text" name="nm_rg" value="<%=strnm_rg%>" size="20" maxlength="15" class="inputs"></td>
								<td><input type="text" name="dt_rg_emissao" value="<%=strdt_rg_emissao%>" size="10" maxlength="15" class="inputs"> &nbsp;
									<input type="text" name="nm_rg_expedidor" value="<%=strnm_rg_expedidor%>" size="8" maxlength="50" class="inputs"></td>
								<td><input type="text" name="nm_cpf" value="<%=strnm_cpf%>" size="20" maxlength="14" class="inputs"></td>
							</tr>
							<tr>
								<td><b>Tit. Eleitor</b></td>
								<td><b>Zona &nbsp; &nbsp; &nbsp; - &nbsp; Secção</b></td>
							</tr>
							<tr>								
								<td><input type="text" name="nm_tit_eleitor" value="<%=strnm_tit_eleitor%>" size="25" maxlength="14" class="inputs"></td>
								<td><input type="text" name="nr_tit_eleitor_zona" value="<%=strnr_tit_eleitor_zona%>" size="5" maxlength="14" class="inputs"> &nbsp;
									<input type="text" name="nr_tit_eleitor_seccao" value="<%=strnr_tit_eleitor_seccao%>" size="5" maxlength="14" class="inputs"></td>
							</tr>
							<tr>
								<td><b>Reservista</b></td>
								<td colspan="2"><b>Categoria</b></td>
							</tr>
							<tr>
								<td><input type="text" name="nm_reservista" value="<%=strnm_reservista%>" size="25" maxlength="14" class="inputs"></td>
								<td colspan="2"><input type="text" name="nm_reservista_cat" value="<%=strnm_reservista_cat%>" size="50" maxlength="200" class="inputs"></td>
							</tr>
							<!--*****************-->
							<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr>
								<td><b>Dependentes</b></td>
							</tr>
							<tr>
								<td>Nome</td>
								<td>Parentesco<img src="../../imagens/px.gif" alt="" width="30" height="1" border="0">Sexo</td>
								<td>Data nasc.</td>
								<td>Observações.</td>
							</tr>
							<!--form name="formdependente"-->
							<tr>
								<td><input type="text" name="nm_dependente_nome" size="20" maxlength="30" class="inputs"></td>
								<td><select name="cd_dependente_parentesco" class="inputs">
										<option value=""></option>
										<%strsql = "Select * From TBL_parentesco "'Where cd_codigo=1 or cd_codigo = 2"
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
												
									</select><img src="../../imagens/px.gif" alt="" width="12" height="1" border="0">
									<select name="cd_dependente_sexo" class="inputs">
										<option value=""></option>
										<option value="1">Masc.</option>
										<option value="2">Fem.</option>
									</select>
									</td>
								<td><input type="text" name="dt_dependente_nascimento" value="" size="10" maxlength="10" class="inputs"></td>
								<td><input type="text" name="cd_dependente_obs" value="" size="15" class="inputs">								
									<input type="button" name="inclui_dependente" value="+" onFocus="adiciona_dependente(document.form.nm_dependente_nome.value,'',document.form.cd_dependente_parentesco.value,document.form.dt_dependente_nascimento.value,document.form.cd_dependente_obs.value,document.form.cd_dependente_sexo.value,document.form.dependentes_total.value)"></td>
							</tr>
								<input type="text" name="dependentes_total" value="" size="100"><br>	
								<input type="hidden" name="dependentes_result" size="70">
							<tr><td colspan="4"><span id='divDependente_lista'> &nbsp;</span></td></tr>
							<!--*****************-->
							<tr>
								<td colspan="4">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_dependente Where cd_funcionario='"&str_funcionario&"' order by cd_parentesco,dt_nascimento,nm_nome"
											Set rs_dependente = dbconn.execute(strsql)
											
										if rs_dependente.EOF Then%>
											<td colspan="4" style="width:300px; font-weight:bold;" align="center">&nbsp;Não Possui dependentes</td>
										<%end if
										while not rs_dependente.EOF
											cd_dependente = rs_dependente("cd_codigo")
											nm_dependente_nome = rs_dependente("nm_nome")
											cd_dependente_parentesco = rs_dependente("cd_parentesco")
											cd_dependente_sexo = rs_dependente("cd_sexo")
											dt_dependente_nascimento = rs_dependente("dt_nascimento")
											nm_dependente_obs = rs_dependente("nm_obs")
										
										conta_dependentes = conta_dependentes + 1
										if cd_dependente_parentesco = 1 then
											conta_filhos = conta_filhos + 1
										end if%>
										<%if cabeca_dependente = "" Then%>
										<tr bgcolor="#808080" style="color:white;">
											<td>&nbsp;<b>&nbsp;</b>&nbsp;</td>
											<td>&nbsp;<b>Nome</b>&nbsp;</td>
											<td>&nbsp;<b>Parentesco</b>&nbsp;</td>
											<td>&nbsp;<b>Sexo</b>&nbsp;</td>
											<td>&nbsp;<b>Nascimento</b>&nbsp;</td>
											<td>&nbsp;<b>Idade</b>&nbsp;</td>
											<td>&nbsp;<b>Observações</b>&nbsp;</td>											
										</tr>
										<%end if%>
										<tr>
											<td><b><%=conta_dependentes%></b></td>
											<td>&nbsp;<%=nm_dependente_nome%></td>											
											<%strsql = "Select * From TBL_parentesco Where cd_codigo="&cd_dependente_parentesco&""
											Set rs_dependente_parent = dbconn.execute(strsql)
											
											While Not rs_dependente_parent.eof
												nm_dependente_parentesco = rs_dependente_parent("nm_parentesco")
											rs_dependente_parent.movenext
											wend
											'rs_dependente_parent.close
											'Set rs_dependente_parent = nothing%>
											
											<td>&nbsp;<%=nm_dependente_parentesco%></td>
											<td>&nbsp;<%if cd_dependente_sexo = 1 then
															response.write("Masc")
														elseif cd_dependente_sexo = 2 then
															response.write("Fem")
														end if%></td>
											<td>&nbsp;<%=dt_dependente_nascimento%></td>
										<%idade = datediff("m",dt_dependente_nascimento,now)
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
											<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsDependenteDelete('<%=str_funcionario%>','<%=cd_dependente%>')"></td>
										</tr>
										<%cabeca_dependente = cabeca_dependente + 1
										'conta_dependentes = conta_dependentes + 1
										rs_dependente.movenext
										wend
										if cd_dependente <> "" then%>
										<tr><td colspan="7" bgcolor="gray" style="color:white;"><b>Quantidade de filhos: <%=conta_filhos%></b></td></tr>
										<tr><td colspan="7" bgcolor="#c5c5c5" style="color:#f8f8f8;">Quantidade Total de dependentes: <b><%=conta_dependentes%></b></td></tr>
										<%end if%>
									</table>
								</td>
							</tr>
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
								<td><b>Carteira de vacinação</b></td>
							</tr>
							<tr>
								<td><b>Hepatite B</b></td>
								<td colspan="3">
									<%if int(strnr_hepatite_b) = 3 then
										ckhepatite_b3 = "checked"
										ckhepatite_b2 = "checked"
										ckhepatite_b1 = "checked"
									elseif int(strnr_hepatite_b) = 2 then
										ckhepatite_b2 = "checked"
										ckhepatite_b1 = "checked"
									elseif int(strnr_hepatite_b) = 1 then
										ckhepatite_b1 = "checked"
									end if%>
									<input type="checkbox" name="nr_hepatite_b1" value="1" style="border:0;" class="inputs" <%=ckhepatite_b1%>>1ª dose &nbsp; 
									<input type="checkbox" name="nr_hepatite_b2" value="2" style="border:0;" class="inputs" <%=ckhepatite_b2%>>2ª dose &nbsp;
									<input type="checkbox" name="nr_hepatite_b3" value="3" style="border:0;" class="inputs" <%=ckhepatite_b3%>>3ª dose</td>
							</tr>
							<tr>
								<td><b>Dupla Adulto</b></td>
								<td colspan="3">
									<%if int(strnr_dupla_adulto) = 3 then
										ckdp_adulto_b3 = "checked"
										ckdp_adulto_b2 = "checked"
										ckdp_adulto_b1 = "checked"
									elseif int(strnr_dupla_adulto) = 2 then
										ckdp_adulto_b2 = "checked"
										ckdp_adulto_b1 = "checked"
									elseif int(strnr_dupla_adulto) = 1 then
										ckdp_adulto_b1 = "checked"
									end if%>
									<input type="checkbox" name="nr_dupla_adulto1" value="1" style="border:0;" class="inputs" <%=ckdp_adulto_b1%>>1ª dose &nbsp; 
									<input type="checkbox" name="nr_dupla_adulto2" value="2" style="border:0;" class="inputs" <%=ckdp_adulto_b2%>>2ª dose &nbsp;
									<input type="checkbox" name="nr_dupla_adulto3" value="3" style="border:0;" class="inputs" <%=ckdp_adulto_b3%>>3ª dose &nbsp; &nbsp; 
									<b>Val.:</b> <input type="text" name="dt_dupla_adulto_validade" value="<%=strdt_dupla_adulto_validade%>" size="10" maxlength="10" class="inputs"></td>
							</tr>
							<tr>
								<td><b>SCR (Sarampo, Caxumba, Rubeóla)</b></td>
								<td><input type="radio" name="nr_scr" value="1" <%if int(strnr_scr) = 1 then%>checked<%end if%> style="border:0;" class="inputs"> SIM &nbsp; &nbsp; &nbsp; 
									<input type="radio" name="nr_scr" value="" <%if int(strnr_scr) = "" OR int(strnr_scr) = 0 then%>checked<%end if%> style="border:0;" class="inputs"> Não</td>								
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr>
								<td colspan="2"><b>Exame médico realizado em: </b><input type="text" name="dt_exame_medico" value="<%=strdt_exame_medico%>" size="10" maxlength="10" class="inputs"></td>
								<td colspan="2"><b>Validade: </b><%if strdt_exame_medico <> "" Then%><%=day(strdt_exame_medico)&"/"&month(strdt_exame_medico)&"/"&year(strdt_exame_medico)+1%><%end if%></td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr>
								<td><b>Cararacterísticas Físicas</b></td>
							</tr>
							<tr>
								<td colspan="4"><b>Altura: </b><input type="text" name="nr_caract_altura" value="<%=strnr_caract_altura%>" size="6" maxlength="6" class="inputs"> &nbsp; &nbsp;
								<b>Peso: </b><input type="text" name="nr_caract_peso" value="<%=strnr_caract_peso%>" size="6" maxlength="6" class="inputs"> &nbsp; &nbsp;
								<b>Manequim: </b><input type="text" name="nr_caract_manequim" value="<%=strnr_caract_manequim%>" size="6" maxlength="6" class="inputs"> &nbsp; &nbsp;
								<b>Calçado: </b><input type="text" name="nr_caract_calcado" value="<%=strnr_caract_calcado%>" size="6" maxlength="6" class="inputs"> &nbsp; &nbsp;
								</td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr>
								<td colspan="4"><b>Escolaridade</b><img src="../../imagens/px.gif" alt="" width="65" height="1" border="0">
								<b>Instituição</b><img src="../../imagens/px.gif" alt="" width="85" height="1" border="0">
								<b>Curso</b><img src="../../imagens/px.gif" alt="" width="105" height="1" border="0"> 
								<b>andamento</b><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"> 
								<b>Término</b><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0">
								<b>Observação</b></td>
							</tr>
							<tr>
								<td colspan="4"><input type="hidden" name="escolaridade_total" value="" size="100"><br>
									<input type="hidden" name="escolaridade_result" size="70"></td></tr>
							<tr>
								<td colspan="4">								
									<select name="cd_ensino_grau" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_escolaridade_grau order by cd_codigo"
										Set rs_grau = dbconn.execute(xsql)
										while not rs_grau.EOF
										cd_grau = rs_grau("cd_codigo")
										nm_grau = rs_grau("nm_grau")%>
											<option value="<%=cd_grau%>"><%=nm_grau%></option>
										<%rs_grau.movenext
										Wend%>
									</select> &nbsp; 
								<input type="text" name="nm_instituicao" class="inputs">
								<input type="text" name="nm_curso" class="inputs">
									<select name="cd_ensino_andamento" class="inputs">
										<option value="0"></option>
										<%xsql = "Select * From TBL_escolaridade_andamento order by cd_codigo DESC"
										Set rs_andamento = dbconn.execute(xsql)
										while not rs_andamento.EOF
										cd_andamento = rs_andamento("cd_codigo")
										nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
											<option value="<%=cd_andamento%>"><%=nm_ensino_andamento%></option>
										<%rs_andamento.movenext
										Wend%>
									</select>
									<input type="text" name="dt_termino" maxlength="10" size="10" class="inputs">
									<input type="text" name="nm_obs" size="20" maxlength="200" class="inputs">
									<input type="button" name="inclui_escolaridade" value="+" onFocus="adiciona_escolaridade(document.form.cd_ensino_grau.value,'',document.form.nm_instituicao.value, document.form.nm_curso.value, document.form.cd_ensino_andamento.value, document.form.dt_termino.value, document.form.nm_obs.value, document.form.escolaridade_total.value)">
								</td>
							</tr>
							<tr><td colspan="4"><span id='divEscolaridade_lista'> &nbsp;</span></td></tr>
							<tr>
								<td colspan="4">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_escolaridade Where cd_funcionario='"&str_funcionario&"' order by cd_ensino_grau"
											Set rs_escolaridade = dbconn.execute(strsql)
											
										'if rs_escolaridade.EOF Then%>
											<!--td colspan="100%" style="width:300px; font-weight:bold;" align="center">&nbsp;</td-->
										<%'end if
										while not rs_escolaridade.EOF
											cd_escolaridade = rs_escolaridade("cd_codigo")
											cd_escolaridade_grau = rs_escolaridade("cd_ensino_grau")
											cd_escolaridade_andamento = rs_escolaridade("cd_ensino_andamento")
											nm_escolaridade_instituicao = rs_escolaridade("nm_instituicao")
											nm_escolaridade_curso = rs_escolaridade("nm_curso")
											nm_escolaridade_obs = rs_escolaridade("nm_obs")
											dt_escolaridade_termino = rs_escolaridade("dt_termino")
										
										'conta_dependentes = conta_dependentes + 1
										'if cd_dependente_parentesco = 1 then
										'	conta_filhos = conta_filhos + 1
										'end if%>
										<%if cabeca_escolaridade = "" Then%>
										<tr bgcolor="#808080" style="color:white;">
											<td>&nbsp;<b>Grau</b>&nbsp;</td>
											<td>&nbsp;<b>Instituição</b>&nbsp;</td>
											<td>&nbsp;<b>Curso</b>&nbsp;</td>
											<td>&nbsp;<b>Andamento</b>&nbsp;</td>
											<td>&nbsp;<b>Término</b>&nbsp;</td>
											<td>&nbsp;<b>Observação</b></td>
											<td>&nbsp;</td>
										</tr>
										<%end if%>
										<tr>
											<td>&nbsp;<%xsql = "Select * From TBL_escolaridade_grau where cd_codigo='"&cd_escolaridade_grau&"'"
															Set rs_grau = dbconn.execute(xsql)
															if not rs_grau.EOF then
															cd_grau = rs_grau("cd_codigo")
															nm_grau = rs_grau("nm_grau")%>
																<%=nm_grau%>
															<%rs_grau.movenext
															end if%></td>											
											<td>&nbsp;<%=nm_escolaridade_instituicao%></td>
											<td>&nbsp;<%=nm_escolaridade_curso%></td>
											<td>&nbsp;<%xsql = "Select * From TBL_escolaridade_andamento where cd_codigo='"&cd_escolaridade_andamento&"'"
															Set rs_andamento = dbconn.execute(xsql)
															if not rs_andamento.EOF then
															cd_andamento = rs_andamento("cd_codigo")
															nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
																<%=nm_ensino_andamento%>
															<%rs_andamento.movenext
															end if%></td>
											<td>&nbsp;<%=dt_escolaridade_termino%></td>
											<td>&nbsp;<%=nm_escolaridade_obs%></td>					
											<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsEscolaridadeDelete('<%=str_funcionario%>','<%=cd_escolaridade%>')"></td>
										</tr>
										<%cabeca_escolaridade = cabeca_escolaridade + 1
										rs_escolaridade.movenext
										wend%>
									</table>
								</td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr>
								<td colspan="4">
									<b>Emprego Anterior</b><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0">
									<b>Função</b><img src="../../imagens/px.gif" alt="" width="105" height="1" border="0">
									<b>entrada</b><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0">
									<b>Saída</b><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0">						
							</tr>
							<!--tr><td colspan="4"--><input type="hidden" name="emprego_total" value="" size="100">
								<input type="hidden" name="emprego_result" size="70"><!--/td></tr-->
							<tr>
								<td colspan="4">
									<input type="text" name="nm_emprego_empresa" class="inputs">
									<input type="text" name="nm_emprego_cargo" class="inputs">
									<input type="text" name="dt_emprego_entrada" maxlength="10" size="10" class="inputs">
									<input type="text" name="dt_emprego_saida" maxlength="10" size="10" class="inputs">
									<input type="text" name="dt_emprego_obs" size="20" maxlength="200" class="inputs">
									<input type="button" name="inclui_emprego" value="+" onFocus="adiciona_emprego(document.form.nm_emprego_empresa.value,'',document.form.nm_emprego_cargo.value, document.form.dt_emprego_entrada.value, document.form.dt_emprego_saida.value, document.form.dt_emprego_obs.value, document.form.emprego_total.value)">
								</td>
							</tr>
							<tr><td colspan="4"><span id='divEmprego_lista'> &nbsp;</span></td></tr>
							<tr>
								<td colspan="4">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_emprego_anterior Where cd_funcionario='"&str_funcionario&"' order by dt_entrada"
											Set rs_emprego = dbconn.execute(strsql)
											
										'if rs_escolaridade.EOF Then%>
											<!--td colspan="100%" style="width:300px; font-weight:bold;" align="center">&nbsp;</td-->
										<%'end if
										while not rs_emprego.EOF
											
											cd_emprego = rs_emprego("cd_codigo")
											nm_emprego_empresa = rs_emprego("nm_empresa")
											nm_emprego_cargo = rs_emprego("nm_cargo")
											dt_emprego_entrada = rs_emprego("dt_entrada")
											dt_emprego_saida = rs_emprego("dt_saida")
											nm_emprego_obs = rs_emprego("nm_obs")
											
											conta_linha = conta_linha + 1%>
										<%if cabeca_emprego = "" Then%>
										<tr bgcolor="#808080" style="color:white;">
											<td>&nbsp;</td>
											<td>&nbsp;<b>Emprego Anterior</b>&nbsp;</td>
											<td>&nbsp;<b>Função</b>&nbsp;</td>
											<td>&nbsp;<b>Data Entrada</b>&nbsp;</td>
											<td>&nbsp;<b>Data saída</b>&nbsp;</td>
											<td>&nbsp;<b>Observação</b></td>
											<td>&nbsp;</td>
										</tr>
										<%end if%>
										<tr>
											<td>&nbsp;<%=conta_linha%></td>
											<td>&nbsp;<%=nm_emprego_empresa%></td>											
											<td>&nbsp;<%=nm_emprego_cargo%></td>
											<td>&nbsp;<%=dt_emprego_entrada%></td>
											<td>&nbsp;<%=dt_emprego_saida%></td>
											<td>&nbsp;<%=nm_emprego_obs%></td>
											<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsEmpregodeDelete('<%=str_funcionario%>','<%=cd_emprego%>')"></td>
										</tr>
										<%cabeca_emprego = cabeca_emprego + 1
										rs_emprego.movenext
										wend%>
									</table>
								</td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr>
								<td><b>Possui parente na Empresa? </b><%if strnm_parente <> "" Then%>SIM<%Else%>NÃO<%end if%></td>
								<td colspan="2"><b>Nome: </b><input type="text" name="nm_parente" value="<%=strnm_parente%>" size="50" maxlength="50" class="inputs"></td>
								<td><b>Parentesco: </b>
									<select name="cd_parente_parentesco" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_parentesco order by nm_parentesco DESC"
										Set rs_parentesco = dbconn.execute(xsql)
										while not rs_parentesco.EOF
										cd_parentesco = rs_parentesco("cd_codigo")
										nm_parentesco = rs_parentesco("nm_parentesco")%>
											<option value="<%=cd_parentesco%>" <%if int(strcd_parente_parentesco) = cd_parentesco then%> SELECTED<%end if%>><%=nm_parentesco%></option>
										<%rs_parentesco.movenext
										Wend%>
									</select>
								</td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr>
								<td><b>Residência: (tipo)</b></td>
								<td><b>Financiada?</b></td>
								<td><b>Reside</b></td>
							</tr>
							<tr>
								<td><select name="cd_residencia_tipo" class="inputs">
										<option value="0"></option>
										<option value="1" <%if int(strcd_residencia_tipo) = 1 Then%>selected<%end if%>><b>Própria<b></option>
										<option value="2" <%if int(strcd_residencia_tipo) = 2 Then%>selected<%end if%>><b>Alugada<b></option>
									</select>
								<td><select name="cd_residencia_financ" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_residencia_financ) = 1 Then%>selected<%end if%>>Sim</option>
										<option value="2" <%if int(strcd_residencia_financ) = 2 Then%>selected<%end if%>>Não</option>
									</select></td>
								<td><select name="cd_residencia_reside" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_residencia_reside) = 1 Then%>selected<%end if%>>Só</option>
										<option value="2" <%if int(strcd_residencia_reside) = 2 Then%>selected<%end if%>>Companheiro</option>
										<option value="3" <%if int(strcd_residencia_reside) = 3 Then%>selected<%end if%>>Pais</option>
										<option value="4" <%if int(strcd_residencia_reside) = 4 Then%>selected<%end if%>>Parentes</option>
										<option value="5" <%if int(strcd_residencia_reside) = 5 Then%>selected<%end if%>>Amigos</option>
									</select></td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr>
								<td><b>Veículo: (tipo)</b></td>
								<td><b>Financiado?</b></td>
							</tr>
							<tr>
								<td><select name="cd_veiculo_tipo" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_veiculo_tipo) = 1 Then%>selected<%end if%>>Carro</option>
										<option value="2" <%if int(strcd_veiculo_tipo) = 2 Then%>selected<%end if%>>Moto</option>
									</select>				
								</td>
								<td><select name="cd_veiculo_financ" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_veiculo_financ) = 1 Then%>selected<%end if%>>Sim</option>
										<option value="2" <%if int(strcd_veiculo_financ) = 2 Then%>selected<%end if%>>Não</option>
									</select></td>
							</tr>
							<tr><td colspan="4"><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<!--tr>
								<td colspan="4"--><input type="hidden" name="valores_total" value="" size="100">
									<input type="hidden" name="valores_result" size="70"><!--/td></tr-->
							<%if session_cd_usuario = 3 or session_cd_usuario = 52 or session_cd_usuario = 56 or session_cd_usuario = 46 then%>
							<tr>
								<td colspan="2"><b>Valores</b></td>
								<td><b>data início</b></td>
								<td><b>Observações</b></td>
							</tr>
							<tr>
								<td>								
									<select name="cd_valor_tipo" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_funcionario_valores_tipos order by nm_valor"
										Set rs_valor = dbconn.execute(xsql)
										while not rs_valor.EOF
										cd_valor = rs_valor("cd_codigo")
										nm_valor = rs_valor("nm_valor")%>
											<option value="<%=cd_valor%>"><%=nm_valor%></option>
										<%rs_valor.movenext
										Wend%>
									</select> &nbsp; 
								</td>
								<td>
								R$ <input type="text" name="nr_valor" class="inputs" value="">
								</td>
								<td>
									<input type="text" name="dt_valor_atualizacao" maxlength="10" size="10" class="inputs">
								</td>
								<td>	<input type="text" name="nm_valor_obs" size="20" maxlength="200" class="inputs">
									<input type="button" name="inclui_valores" value="+" onFocus="adiciona_valores(document.form.cd_valor_tipo.value,'',document.form.nr_valor.value, document.form.dt_valor_atualizacao.value, document.form.nm_valor_obs.value, document.form.valores_total.value)">
								</td>
							</tr>
							<tr><td colspan="4"><span id='divValores_lista'> &nbsp;</span></td></tr>
							<tr>
								<td colspan="4">
									<table border="1" style="border:1px solid gray; border-collapse:collapse;">
										
										<%strsql = "Select * From TBL_funcionario_valores Where cd_funcionario='"&str_funcionario&"' order by cd_tipo"
											Set rs_valor = dbconn.execute(strsql)
											
										'if rs_escolaridade.EOF Then%>
											<!--td colspan="100%" style="width:300px; font-weight:bold;" align="center">&nbsp;</td-->
										<%'end if
										while not rs_valor.EOF
											
											cd_valor = rs_valor("cd_codigo")
											cd_valor_tipo = rs_valor("cd_tipo")
											nr_valor = rs_valor("nr_valor")
											cd_funcionario = rs_valor("cd_funcionario")
											dt_valor_atualizacao = rs_valor("dt_atualizacao")
											nm_valor_obs = rs_valor("nm_obs")
											
											conta_linha = conta_linha + 1%>
										<%if cabeca_emprego = "" Then%>
										<tr bgcolor="#808080" style="color:white;">
											<td>&nbsp;</td>
											<td>&nbsp;<b>Tipo</b>&nbsp;</td>
											<td>&nbsp;<b>Valor</b>&nbsp;</td>
											<td>&nbsp;<b>Data</b>&nbsp;</td>
											<td>&nbsp;<b>Observação</b></td>
											<td>&nbsp;</td>
										</tr>
										<%end if%>
										<tr>
											<td>&nbsp;<%=conta_linha%></td>											
											<td>&nbsp;<%xsql = "Select * From TBL_funcionario_valores_tipos where cd_codigo='"&cd_valor_tipo&"'"
														Set rs_tipo = dbconn.execute(xsql)
														while not rs_tipo.EOF
														nm_valor = rs_tipo("nm_valor")%>
															<%=nm_valor%>
														<%rs_tipo.movenext
														Wend%>
											</td>
											<td>&nbsp;<%=nr_valor%></td>
											<td>&nbsp;<%=dt_valor_atualizacao%></td>
											<td>&nbsp;<%=nm_valor_obs%></td>
											<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsValorDelete('<%=str_funcionario%>','<%=cd_valor%>')"></td>
										</tr>
										<%cabeca_emprego = cabeca_emprego + 1
										rs_valor.movenext
										wend%>
									</table>
								</td>
							</tr>
							<%end if%>
							
							
							
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
								<td>&nbsp;<b>Co-participação</b></td>
								<td>&nbsp;<b>Assistência Odontológica</b></td>
								<td>&nbsp;<b>Co-participação</b></td>
							</tr>
							<tr>
								<td><input type="text" name="cd_assistencia_medica_matricula" value="<%=str_assistencia_medica_matricula%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="nr_assistencia_medica_coparticipacao" size="3" maxlength="3" value="<%=str_assistencia_medica_coparticipacao%>" class="inputs">%</td>
								<td><input type="text" name="cd_assistencia_odonto_matricula" value="<%=str_assistencia_odonto_matricula%>" size="20" maxlength="20" class="inputs"></td>
								<td><input type="text" name="nr_assistencia_odonto_coparticipacao" size="3" maxlength="3" value="<%=str_assistencia_odonto_coparticipacao%>" class="inputs">%</td>
							</tr>
							<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<%if strcod = "" Then%>
							<tr>
								<td style="color:red;"><b>Contrato</b> &nbsp;</td>
								<td style="color:red;"><b>Data de admissão</b></td>
									<%If strcod <> "" Then%>
										<td><b>Data de recisão</b></td>
									<%else%>
										<td>&nbsp;</td>
									<%End If%>
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
								<td><!--input type="text" name="dt_contratado" value="<%'=Strdt_contratado%>" size="20" maxlength="10" class="inputs"-->
									<input type="text" name="cd_dia" value="" size="3" maxlength="2" class="inputs">/
									<input type="text" name="cd_mes" value="" size="3" maxlength="2" class="inputs">/
									<input type="text" name="cd_ano" value="" size="6" maxlength="4" class="inputs"></td>								
								<td>&nbsp;</td>		
							</tr>
							<%else%>						
							<tr>
								<td><b>Contrato</b> &nbsp;</td>
								<td style="color:red;"><b>Data de admissão</b></td>
								<%If strcod <> "" Then%>
								<td><b>Data de recisão</b></td><%else%><td>&nbsp;</td><%End If%>
								<td>&nbsp;</td>
							</tr>
							
								<%strsql = "Select * From VIEW_funcionario_contrato_lista Where cd_funcionario = "&str_funcionario&" order by dt_admissao desc"
			  					Set rs_contrato = dbconn.execute(strsql)%>
								<%while not rs_contrato.EOF
									cd_funcionario = rs_contrato("cd_funcionario")
									cod_contrato = rs_contrato("cod_contrato")
									nm_regime_trab = rs_contrato("nm_regime_trab")
									dt_admissao = rs_contrato("dt_admissao")
									dt_demissao = rs_contrato("dt_demissao")%>
								<tr>
									<td>&nbsp;<a href="javascript:void();" onclick="window.open('empresa/janelas/funcionarios_contrato_altera.asp?cod_contrato=<%=cod_contrato%>&cd_funcionario=<%=str_funcionario%>', 'janela_vdlap', 'width=335, height=205'); return false;"><%=nm_regime_trab%></a></td>
									<td>&nbsp;<%=zero(day(dt_admissao))%>/<%=mesdoano(month(dt_admissao))%>/<%=year(dt_admissao)%></td>
									<td>&nbsp;<%=dt_demissao%></td>
									<td>&nbsp;&nbsp;
									<%if dt_demissao <> "" then
										if nr_contrato = "" then%>
											<a href="javascript:;" onclick="window.open('empresa/janelas/funcionarios_contrato.asp?cd_funcionario=<%=str_funcionario%>', 'janela_vdlap', 'width=350, height=400'); return false;"><b>*** Novo Contrato ***<!--img src="../../imagens/ic_aprovado.gif" alt="Novo Contrato." width="10" height="12" border="0" --></b></a>
										<%end if%>
									<%else%>
											<a href="javascript:;" onclick="window.open('empresa/janelas/funcionarios_recisao.asp?cod_contrato=<%=cod_contrato%>&cd_funcionario=<%=cd_funcionario%>','principal','width=370, height=400'); return false;">Rescindir contrato?</a>&nbsp; <!--day(dt_demissao) = "" then%-->
									<%end if%></td>
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
									xsql_cargo = "select * From View_funcionario_cargo Where cd_funcionario='"&strcod&"'"
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
									<%strsql ="TBL_unidades where cd_hospital > 0 and cd_status = 1"
								  	Set rs_uni = dbconn.execute(strsql)%>
								<%if strcod = "" Then%>
									<select name="cd_unidade" class="inputs"> 
										<option value="">Selecione</option>
										<%Do While Not rs_uni.eof%>
										<option value="<%=rs_uni("cd_codigo")%>"><%=rs_uni("nm_Unidade")%></option>
										<%rs_uni.movenext
										loop
										rs_uni.close
										Set rs_uni = nothing%>
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
										<input type="text" name="hr_saida" size="6" maxlength="5"  class="inputs"></td>
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
							<tr class="no_print">
								<td colspan="4" valign="middle"><A href="empresa.asp?tipo=lista"><img src="imagens/bt_lista.gif" alt="Listar" border="0" width="119" height="22" border="1"></a>
								<input type="image" src="imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0">
								<%=stracao%></td>
							</tr>
							<%if int(session_cd_usuario) = 46 then%>
							<tr class="no_print"><td><br><br></td></tr>
							<tr class="no_print"><td colspan="4"><input onclick="javascript:JsDelete('<%=strcod%>','1')" type="button" value="Apagar funcionario" title="Apagar 1">
								</td>
							</tr>
							<%end if%>
													
							</form>
							<%If strcod <> "" then%>
							
							<tr>
								<td>
								<form name="form_ex" action="empresa/acoes/empr_funcionarios_acao.asp" id="forms" method="post"  enctype="multipart/form-data">
									<input type="hidden" name="acao" value="excluir">
									<input type="hidden" name="cod" value="<%=cd_funcionario%>">
								</form></td>
							</tr>
							<%End if%>
							<tr><td>&nbsp;</td></tr>
							
						</table>
					</td></tr>
				</table>
						<br class="no_print">
						<br class="no_print">
							
					
									 								
							<%End if%>
			<SCRIPT LANGUAGE="javascript">
 {
        //MaskInput(document.forms.nm_cep, "99999-99");
		//MaskInput(document.forms.dt_nascimento, "99/99/9999");	    
 }
</SCRIPT>
			
			<SCRIPT LANGUAGE="javascript">
				function JsDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o funcionario?"))
				  {
				  document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir&protege_exclusao='+cod2+'');
					}
					  }
				
				function JsContatoDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o contato?"))
				  {
				  document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&cod_cont='+cod2+'&acao=apaga_contato');
					}
					  }  
					  
				function JsDependenteDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o dependente?"))
				  {
				  document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&cod_dep='+cod2+'&acao=apaga_dependente');
					}
					  }  
				
				function JsEscolaridadeDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse registro de escolaridade?"))
				  {
				  document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&cod_esc='+cod2+'&acao=apaga_escolaridade');
					}
					  }    
				
				function JsEmpregodeDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse registro de emprego anterior?"))
				  {
				  document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&cod_emp='+cod2+'&acao=apaga_emprego');
					}
					  } 
				
				function JsValorDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse valor?"))
				  {
				  document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&cod_valor='+cod2+'&acao=apaga_valor');
					}
					  } 
				
					  
				//function JsContatoDelete(cod){
					//if (confirm ("Tem certeza que deseja excluir o contato?")){
				  //document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir');
					//}
					  //}
			</SCRIPT>