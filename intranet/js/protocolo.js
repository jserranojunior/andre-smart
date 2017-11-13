
<script type="text/javascript" language="javascript">
function mostra_convenio(cd_convenio) {
    busca = Form.cd_convenio.value;
	busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	  
	   loadXMLDoc("../ajax/protocolo/ajax_convenio.asp?cd_convenio="+busca,function(){
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200){
		    			document.getElementById("divAssist_m1").innerHTML = xmlhttp.responseText;
		    		}					
		  	});
}
 
function mostra_medico(cd_crm) {
     if (Form.cd_crm.value.length >= 3){
		busca = Form.cd_crm.value;
		busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
			/*if (isNaN(busca)!= true){
				document.getElementById("cd_crm_x").value = busca
		 	}*/
		 loadXMLDoc("../ajax/protocolo/ajax_medico.asp?cd_crm="+busca,function(){
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200){
		    			document.getElementById("divMed").innerHTML = xmlhttp.responseText;
		    		}					
		  	});
		}
}
 
 function mostra_rack(cd_rack) {
     unidade_rack = document.Form.cd_unidade.value;
	document.Form.cd_rack_hide.value=unidade_rack+';'+document.Form.cd_rack.value;	
	 
	 busca = Form.cd_rack_hide.value;
	 busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	 
		loadXMLDoc("../ajax/protocolo/ajax_rack.asp?cd_rack_hide="+busca,function(){
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200){
		    			document.getElementById("divRack").innerHTML = xmlhttp.responseText;
		    		}					
		  	});
 }
 
function mostra_espec(cd_especialidade) {
     busca = Form.cd_especialidade.value;
	 busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	 
	 	loadXMLDoc("../ajax/protocolo/ajax_espec.asp?cd_especialidade="+busca,function(){
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200){
		    			document.getElementById("divEspec").innerHTML = xmlhttp.responseText;
		    		}					
		  	});
 }
 
 function mostra_procedimento(cd_procedimento_busca) {
	     busca = Form.cd_procedimento_busca.value;
		 busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
		 
		 loadXMLDoc("../ajax/protocolo/ajax_procedimento_mostra.asp?cd_procedimento_busca="+busca,function(){
		 	if (xmlhttp.readyState==4 && xmlhttp.status==200){
		    			document.getElementById("divPro_mostra").innerHTML = xmlhttp.responseText;
		    		}					
		  	});
 }
 
  function mostra_material(cd_material_busca) {
     busca = Form.cd_material_busca.value;
	 busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	   
	 loadXMLDoc("../ajax/protocolo/ajax_material_mostra.asp?cd_material_busca="+busca,function(){
		  	if (xmlhttp.readyState==4 && xmlhttp.status==200){
		    			document.getElementById("divMat_mostra").innerHTML = xmlhttp.responseText;
		    		}					
		  	});
 } 
 
 function mostra_patrimonio(patrimonio_busca) {     		
	 busca = Form.cd_patrimonio_busca.value;
	 busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	   
	 loadXMLDoc("../ajax/protocolo/ajax_patrimonio_mostra.asp?cd_patrimonio_busca="+busca,function(){
		  		//alert(busca);
				if (xmlhttp.readyState==4 && xmlhttp.status==200){
		    			document.getElementById("divPat_mostra").innerHTML = xmlhttp.responseText;
		    		}					
		  	});
	
 }
</script>
