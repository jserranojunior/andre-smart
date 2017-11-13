<script type="text/javascript" language="javascript">
	function mostra_calendario(dt_calendario) {
		var oHTTPRequest_calendario = createXMLHTTP(); 
		oHTTPRequest_calendario.open("post", "../ajax/calendario/ajax_calendario.asp", true);
		oHTTPRequest_calendario.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_calendario.onreadystatechange=function(){
			if (oHTTPRequest_calendario.readyState==4){
				document.all.divCalendario.innerHTML = oHTTPRequest_calendario.responseText;}}
				oHTTPRequest_calendario.send("dt_calendario=" + calendario.dt_calendario.value);
}
 
	function testa_calendario() {
		alert("asd");	
	}

</script>
