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
function soma_mat(a,b,c,materiais_total){
	a=a;
	b=b;
	c=c;	
	materiais_total=materiais_total;
	nextfield ="cd_material_1";
		
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}	
			
			while (materiais_total.indexOf('$$') != -1) {
	 		materiais_total = materiais_total.replace('$$','$');}
			
			while (materiais_total.indexOf('$$') != -1) {
	 		materiais_total = materiais_total.replace('$$','$');}
						
		if (materiais_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	materiais_total = materiais_total + separador
		if (a != ""){		
			materiais_total = materiais_total + a + '#' + c; //adiciona				
				
			}
		if (b != ""){
			materiais_total = materiais_total.replace(b,''); //subtrai
			}

	document.form.res_mat.value=a;
	//document.form.materiais_total.value=(materiais_total.replace("$$", "$"));
	document.form.materiais_total.value=(materiais_total);
	document.form.cd_material_1.value="";
	document.form.cd_material_2.value="";
	document.form.qtd_material.value="";	
	document.form.cd_material_1.focus();
	
	//chama Ajax	
	{
		var oHTTPRequest_mat = createXMLHTTP(); 
	    oHTTPRequest_mat.open("post", "ajax/ajax_materiais_lista.asp", true);
	    oHTTPRequest_mat.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	    oHTTPRequest_mat.onreadystatechange=function(){
	     if (oHTTPRequest_mat.readyState==4){
	        document.all.divMat_lista.innerHTML = oHTTPRequest_mat.responseText;}}
	       oHTTPRequest_mat.send("materiais_total=" + form.materiais_total.value);
	 }
 
}
//--------------------------------------------------------------------
function foco(){
	document.form.cd_material_1.focus();
}
				
// -->
</script>
<script language="JavaScript" type="text/javascript" src="js/mascarainput.js"></script>
<!--#include file="js/ajax.js"-->

<BODY onLoad="foco()">
<br>
	<table  border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno" align="center">
		<form method="post" action="<%=action_form%>" name="form" id="form">
		<input type="hidden" name="cd_material_2" value="">
		<input type="text" name="materiais_total" value="">
		<input type="text" name="res_mat" size="70">
		<tr>
			<td>&nbsp;Novo item</td>
			<td>&nbsp;Qtd.</td>
		</tr>
		<tr><td>Nome</td><td>Cargo/Posição</td><td>tipo</td><td>Cód.</td><td>Contato</td><td>Obs.</td></tr>	
		<tr>
			<td style="color:gray;">			
			<input type="text" name="cd_material_1" size="50" maxlength="100" class="inputs"></td>
			<td><input type="text" name="qtd_material" size="5" maxlength="3" class="inputs">
			<input type="button" name="somar_mat" value="+" onFocus="soma_mat(document.form.cd_material_1.value,document.form.cd_material_2.value,document.form.qtd_material.value,document.form.materiais_total.value)"></td>
		</tr>
				
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td colspan="2"><span id='divMat_lista'> &nbsp;</span></td>
		</tr>
		<tr>
			<td colspan="100%" bgcolor="#f5f5f5"> </td>
		</tr>	
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>		
		</form>	
</table>
<SCRIPT LANGUAGE="javascript">
	 {
	   	MaskInput(document.form.qtd_material, "999");				
	 }
</SCRIPT>
<br>
</BODY>
</HTML>


