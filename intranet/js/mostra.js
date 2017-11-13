<script type="text/javascript" language="javascript">
  
   function mostra_espec(cd_especialidade) {
     var oHTTPRequest_especialidade = createXMLHTTP(); 
     oHTTPRequest_especialidade.open("post", "../ajax/mostra/ajax_especialidade.asp", true);
     oHTTPRequest_especialidade.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_especialidade.onreadystatechange=function(){
      if (oHTTPRequest_especialidade.readyState==4){
         document.all.divEspecialidade.innerHTML = oHTTPRequest_especialidade.responseText;}}
       oHTTPRequest_especialidade.send("cd_especialidade=" + Form.cd_especialidade.value);
 } 
 
 function mostra_equipamento(cd_equipamento) {
     var oHTTPRequest_equipamento = createXMLHTTP(); 
     oHTTPRequest_equipamento.open("post", "../ajax/mostra/ajax_equipamento.asp", true);
     oHTTPRequest_equipamento.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_equipamento.onreadystatechange=function(){
      if (oHTTPRequest_equipamento.readyState==4){
         document.all.divEquip.innerHTML = oHTTPRequest_equipamento.responseText;}}
       oHTTPRequest_equipamento.send("cd_equipamento=" + Form.cd_equipamento.value);
 } 
</script>
