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
	
function data_final(dia,mes,ano,qtd_dias){
		//alert("teste");
		
		if (dia < 10){ 
			dia = "0" + dia}
		if (mes < 10){ 
			mes = "0" + mes}
		document.getElementById("per_dia_f").value = dia;
		document.getElementById("per_mes_f").value = mes;
		document.getElementById("per_ano_f").value = ano;
		document.getElementById("qtd_dias").value = qtd_dias;
		
	}
	
function data_final2(dias){
	document.getElementById("qtd_dias").value = dias;
	}
	
function calcula_data(qtd_dias,per_dia,per_mes,per_ano) {
	data_inicial = per_dia+"/"+per_mes+"/"+per_ano
	
	document.getElementById("per_dia_i").value = per_dia;
	document.getElementById("per_mes_i").value = per_mes;
	document.getElementById("per_ano_i").value = per_ano;
	document.getElementById("qtd_dias").value = qtd_dias;
	//chama Ajax	
		loadXMLDoc("ajax/ajax_Tfinal.asp?qtd_dias="+qtd_dias+"&dt_inicial="+data_inicial+"&dt_final=",function()
			{
		  		if (xmlhttp.readyState==4)//&& xmlhttp.status==200)
		    		{
		    			document.getElementById("divTfinal").innerHTML = xmlhttp.responseText;
		    		}
		  		}
			);
		
	}

function calcula_dias(per_dia,per_mes,per_ano,per_dia_f,per_mes_f,per_ano_f, qtd_dias) {
	data_inicial = per_dia+"/"+per_mes+"/"+per_ano
	data_final = per_dia_f+"/"+per_mes_f+"/"+per_ano_f
	
	document.getElementById("per_dia_i").value = per_dia;
	document.getElementById("per_mes_i").value = per_mes;
	document.getElementById("per_ano_i").value = per_ano;
	document.getElementById("per_dia_f").value = per_dia_f;
	document.getElementById("per_mes_f").value = per_mes_f;
	document.getElementById("per_ano_f").value = per_ano_f;
	//chama Ajax	
		loadXMLDoc("ajax/ajax_Tfinal.asp?dt_inicial="+data_inicial+"&dt_final="+data_final,function()
			{
		  		if (xmlhttp.readyState==4)//&& xmlhttp.status==200)
		    		{
		    			document.getElementById("divDias").innerHTML = xmlhttp.responseText;
		    		}
		  		}
			);
		
	}
	
	

</script>
