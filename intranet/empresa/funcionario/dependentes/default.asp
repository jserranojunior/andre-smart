<!--#include file="../../../includes/inc_open_connection.asp"-->
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="js/ajax.js"-->
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
		oHTTPRequest_dependentes.open("post", "ajax/ajax_dependentes_lista.asp", true);
		oHTTPRequest_dependentes.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_dependentes.onreadystatechange=function(){
		if (oHTTPRequest_dependentes.readyState==4){
	    document.all.divDependente_lista.innerHTML = oHTTPRequest_dependentes.responseText;}}
	    oHTTPRequest_dependentes.send("dependentes_total=" + form.dependentes_total.value);
	}
 
}
//-------------------------------------------------------------------- -->
</script>

<table border="0">
	<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
	<tr>
		<td><b>Dependentes</b></td>
	</tr>
	<tr><td>Nome</td><td>Parentesco</td><td>Sexo</td><td>Data nasc.</td><td>Observações.</td><td>&nbsp;</td></tr>
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
		<td><input type="button" name="inclui_dependente" value="+" onFocus="adiciona_dependente(document.form.nm_dependente_nome.value,'',document.dependente.cd_dependente_parentesco.value,document.dependente.dt_dependente_nascimento.value,document.dependente.cd_dependente_obs.value,document.dependente.dependentes_total.value)"></td>
	</tr>
	<tr><td colspan="100%"><br>
	<input type="hidden" name="dependentes_total" value="" size="100"><br>	
	<input type="hidden" name="dependentes_result" size="70"></td></tr>	
	<!--/form-->
	<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
	<tr><td colspan="100%"><span id='divDependente_lista'> &nbsp;</span></td></tr>
</table>