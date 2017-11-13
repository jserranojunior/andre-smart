<script type="text/javascript" language="javascript">

 function mostra_patrimonio_e()
 {
	 var oHTTPRequest_pat = createXMLHTTP(); 
     oHTTPRequest_pat.open("Post", "../manutencao_2/ajax/ajax_patrimonio_mostra.asp", true);
     oHTTPRequest_pat.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_pat.onreadystatechange=function(){
      if (oHTTPRequest_pat.readyState==4){
         document.all.divPat_mostra.innerHTML = oHTTPRequest_pat.responseText;}}
       oHTTPRequest_pat.send("no_patrimonio=" + "e:"+ form.no_patrimonio.value);
 }
 
 function mostra_patrimonio_o()
 {
 //setTimeout(
 	  var oHTTPRequest_pat = createXMLHTTP(); 
     oHTTPRequest_pat.open("post", "../manutencao_2/ajax/ajax_patrimonio_mostra.asp", true);
     oHTTPRequest_pat.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_pat.onreadystatechange=function(){
      if (oHTTPRequest_pat.readyState==4){
         document.all.divPat_mostra.innerHTML = oHTTPRequest_pat.responseText;}}
       oHTTPRequest_pat.send("no_patrimonio=" + "o:"+ form.no_patrimonio.value);
//	  ,3000);
 }
</script>
