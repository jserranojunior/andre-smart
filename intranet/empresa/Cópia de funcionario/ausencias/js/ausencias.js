<script type="text/javascript" language="javascript">
function mostra_cid(cd_cid,cd_sexo) {
	if (Form.cd_cid.value.length >= 3 || Form.cd_cid.value.length == " "){
		//cd_cid = Form.cd_cid.value.replace(' ','%20')
		
	     var oHTTPRequest_cid = createXMLHTTP(); 
	     oHTTPRequest_cid.open("post", "empresa/funcionario/ausencias/ajax/ajax_cid.asp", true);
	     oHTTPRequest_cid.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	     oHTTPRequest_cid.onreadystatechange=function(){
	      if (oHTTPRequest_cid.readyState==4){
	         document.all.divCid.innerHTML = oHTTPRequest_cid.responseText;}}
	       //oHTTPRequest_cid.send("cd_cid=" + Form.cd_sexo.value + Form.cd_cid.value);
		   oHTTPRequest_cid.send("cd_cid=" + Form.cd_sexo.value + Form.cd_cid.value);
		 } 
	}
	

function calcula_T(qtd_dias,per_dia_i,per_mes_i,per_ano_i) {
	//if (Form.cd_cid.value.length >= 3 || Form.cd_cid.value.length == " "){
		//qtd_dias = Form.qtd_dias.value
		if (Form.per_dia_i.value == '' || Form.per_mes_i.value == '' || Form.per_ano_i.value == ''){
			f_data = new Date();
				dia_i = f_data.getDate();
				mes_i = f_data.getMonth()+1;
				ano_i = f_data.getFullYear();			
				data_inicial = dia_i + "/" + mes_i + "/" + ano_i
					document.Form.per_dia_i.value=dia_i;
					document.Form.per_mes_i.value=mes_i;
					document.Form.per_ano_i.value=ano_i;
		}
		else{
			data_inicial = Form.per_dia_i.value + "/" + Form.per_mes_i.value + "/" + Form.per_ano_i.value}
		
	     var oHTTPRequest_Tfinal = createXMLHTTP(); 
	     oHTTPRequest_Tfinal.open("post", "empresa/funcionario/ausencias/ajax/ajax_Tfinal.asp", true);
	     oHTTPRequest_Tfinal.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	     oHTTPRequest_Tfinal.onreadystatechange=function(){
	      if (oHTTPRequest_Tfinal.readyState==4){
	         document.all.divTfinal.innerHTML = oHTTPRequest_Tfinal.responseText;}}
		   oHTTPRequest_Tfinal.send("cd_tFinal=" + Form.qtd_dias.value + ";" + data_inicial);
		// } 
	}
</script>
