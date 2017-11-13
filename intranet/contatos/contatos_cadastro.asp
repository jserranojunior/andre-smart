<%Response.Charset="ISO-8859-1" %>
<html>
<head>
</head>
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
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
function adiciona_contato(a,b,c,contatos_total){
	a=a; //nm_nome_contato
	b=b; //nm_posicao
	c=c;
	//d=d;
	//e=e;
	//f=f;
	//g=g;
	contatos_total=contatos_total;
	//nextfield ="cd_material_1";
		
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}	
			
			while (contatos_total.indexOf('$$') != -1) {
	 		contatos_total = contatos_total.replace('$$','$');}
			
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
			//contatos_total = contatos_total + a + '#' + c + d + e + f + g; //adiciona
			contatos_total = contatos_total + a + '#' + c; //adiciona				
				
			}
		if (b != ""){
			contatos_total = contatos_total.replace(b,''); //subtrai
			}

	document.contato.contatos_result.value=a; //a
	//document.contato.nm_nome_contato.value=b; //b
	//document.form.materiais_total.value=(materiais_total.replace("$$", "$"));	
	document.contato.contatos_total.value=(contatos_total); //c
	document.contato.nm_nome_contato.value="";
	document.contato.nm_posicao.value="";	
	//document.contato.cd_material_2.value="";
	//document.contato.qtd_material.value="";	
	//document.contato.cd_material_1.focus();
	
	//chama Ajax	
	{
		var oHTTPRequest_contatos = createXMLHTTP(); 
		oHTTPRequest_contatos.open("post", "ajax/ajax_contatos_lista.asp", true);
		oHTTPRequest_contatos.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_contatos.onreadystatechange=function(){
		if (oHTTPRequest_contatos.readyState==4){
	    document.all.divContatos_lista.innerHTML = oHTTPRequest_contatos.responseText;}}
	    oHTTPRequest_contatos.send("contatos_total=" + contato.contatos_total.value);
	}
 
}
//-------------------------------------------------------------------- -->
</script>
<script language="JavaScript" type="text/javascript" src="js/mascarainput.js"></script>
<!--#include file="js/ajax.js"-->



<table border="1">
	<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
	<tr>
		<td><b>Contatos</b></td>
	</tr>
	<tr><td>Nome</td><td>Cargo/Posição</td><td>tipo</td><td>Cód.</td><td>Contato</td><td>Obs.</td></tr>
	<form name="contato">
	<tr>
		<td><input type="text" name="nm_nome_contato" size="19" maxlength="30" class="inputs"><input type="text" name="nm_nome_excluir" size="19" maxlength="30" class="inputs">
		<!--input type="text" name="cd_material_1" size="19" maxlength="30" class="inputs"></td-->
		<td><input type="text" name="nm_posicao" size="15" maxlength="30" class="inputs">
		<!--input type="text" name="cd_material_2" size="15" maxlength="30" class="inputs"></td-->
		<td><select name="cat_contato" class="inputs">
				<option value=""></option>
				<%'	strsql = "Select * From TBL_contato_cat "'Where cd_codigo=1 or cd_codigo = 2"
				'Set rs_contato = dbconn.execute(strsql)%>
				<%'Do While Not rs_contato.eof
					'	categoria_contato = rs_contato("categoria_contato")
						'cd_contato_cat = rs_contato("cd_codigo")
						'if cd_contato_cat = cat_ then%><%'ckcontr = "selected"%><%'Else%><%'ckcontr=""%><%'end if%>	
						<!--option value="<%'=cd_contato_cat%>" <%'=ckcontr%>><%'=categoria_contato%></option-->
						<%'rs_contato.movenext
						'loop
						'rs_contato.close
						'Set rs_contato = nothing%>
						
				<option value="1">teste</option>			
			</select></td>
		<td><!--input type="text" name="nm_contato_cod" size="3" maxlength="30" class="inputs"-->
		<input type="text" name="qtd_material" size="3" maxlength="30" class="inputs"></td>
		<td><input type="text" name="nm_contato" size="20" maxlength="30" class="inputs"></td>
		<td><input type="text" name="nm_obs" size="20" maxlength="30" class="inputs"></td>
		<td><input type="button" name="inclui_contato" value="+" onFocus="adiciona_contato(document.contato.nm_nome_contato.value,document.contato.nm_nome_excluir.value,document.contato.nm_posicao.value,document.contato.contatos_total.value)"></td>
	</tr>
	<tr><td colspan="100%"><br>
	Total:<input type="text" name="contatos_total" value=""><br>	
	Resultado<input type="text" name="contatos_result" size="70"></td></tr>	
	</form>			
	<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
	<tr><td><span id='divContatos_lista'> &nbsp;</span></td></tr>
</table>