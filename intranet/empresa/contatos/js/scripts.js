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
			while (b.indexOf(' ') != -1) {
	 		b = b.replace(' ','%20');}
			while (c.indexOf(' ') != -1) {
	 		c = c.replace(' ','%20');}
			//contatos_total = contatos_total.replace(' ','%20');}
			
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
	document.contato.contatos_total.value=(contatos_total.replace("$$", "$"));	
	document.contato.contatos_total.value=(contatos_total); //c
	document.contato.nm_nome_contato.value="";
	document.contato.nm_posicao.value="";
	
	//chama Ajax	
	{
		var oHTTPRequest_contatos = createXMLHTTP(); 
		oHTTPRequest_contatos.open("post", "../../ajax/ajax_contatos_lista.asp", true);
		oHTTPRequest_contatos.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_contatos.onreadystatechange=function(){
		if (oHTTPRequest_contatos.readyState==4){
	    document.all.divContatos_lista.innerHTML = oHTTPRequest_contatos.responseText;}}
	    oHTTPRequest_contatos.send("contatos_total=" + contato.contatos_total.value);
	}
 
}
//-------------------------------------------------------------------- -->
</script>