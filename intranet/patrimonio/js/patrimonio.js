<script type="text/javascript" language="javascript">
	function patrimonio_movimentacao()
		{
			document.form.dia_movimento_obj.value = "";
			document.form.mes_movimento_obj.value = "";
			document.form.ano_movimento_obj.value = "";
			document.form.und_destino_obj.value = "";			
			document.form.nm_observacoes_obj.value = "";
			document.form.rack_destino_obj.value = "";
			document.form.nm_solicitante_obj.value = "";
			
			loadXMLDoc("ajax/ajax_patrimonio_movimentacao.asp?mostra_movimentacao="+ form.cd_situacao.value,function()
		  		{
				//alert(xmlhttp.readyState+' - '+xmlhttp.status);
		  			if (xmlhttp.readyState==4 && xmlhttp.status==200)
		    			{
		    				document.getElementById("divPatmov").innerHTML = xmlhttp.responseText;
							//document.getElementById("divPatmov").innerHTML = xmlhttp.responseText;
		    			}
		  			}
				);
			
		}
 /*
 function patrimonio_empr_novo(movimentacao) {
     var oHTTPRequest_mov = createXMLHTTP(); 
     oHTTPRequest_mov.open("post", "ajax/ajax_patrimonio_movimentacao.asp", true);
     oHTTPRequest_mov.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_mov.onreadystatechange=function(){
      if (oHTTPRequest_mov.readyState==4){
         document.all.divPatmov_novo.innerHTML = oHTTPRequest_mov.responseText;}}
       oHTTPRequest_mov.send("mostra_movimentacao=" + form.cd_nova_situacao.value);   
 }
*/
function patrimonio_empr_novo()
		{
			loadXMLDoc("ajax/ajax_patrimonio_movimentacao.asp?mostra_movimentacao="+ form.cd_nova_situacao.value,function()
		  		{
		  			if (xmlhttp.readyState==4 && xmlhttp.status==200)
		    			{
		    				document.getElementById("divPatmov_novo").innerHTML = xmlhttp.responseText;
		    			}
		  			}
				);
		}
</script>

