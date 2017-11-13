<script type="text/javascript" language="javascript">
 function fill_equip()
 {
     var oHTTPRequest_equip = createXMLHTTP(); 
     oHTTPRequest_equip.open("post", "../ajax/manutencao/ajax_equipamentos.asp", true);
     oHTTPRequest_equip.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_equip.onreadystatechange=function(){
      if (oHTTPRequest_equip.readyState==4){
         document.all.divEquip.innerHTML = oHTTPRequest_equip.responseText;}}
       oHTTPRequest_equip.send("cd_equipamento=" + form.cd_equipamento.value);
 }

 function fill_marca()
 {
     var oHTTPRequest_marca = createXMLHTTP(); 
     oHTTPRequest_marca.open("post", "../ajax/manutencao/ajax_marcas.asp", true);
     oHTTPRequest_marca.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_marca.onreadystatechange=function(){
      if (oHTTPRequest_marca.readyState==4){
         document.all.divMarca.innerHTML = oHTTPRequest_marca.responseText;}}
       oHTTPRequest_marca.send("cd_marca=" + form.cd_marca.value);
 }
 
  function fill_unid()
 {
     var oHTTPRequest_unid = createXMLHTTP(); 
     oHTTPRequest_unid.open("post", "../ajax/manutencao/ajax_unidades.asp", true);
     oHTTPRequest_unid.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_unid.onreadystatechange=function(){
      if (oHTTPRequest_unid.readyState==4){
         document.all.divUnid.innerHTML = oHTTPRequest_unid.responseText;}}
       oHTTPRequest_unid.send("nm_unidade=" + form.nm_unidade.value);
 }
 
 function fill_forn()
 {
     var oHTTPRequest_forn = createXMLHTTP(); 
     oHTTPRequest_forn.open("post", "../ajax/manutencao/ajax_fornecedores.asp", true);
     oHTTPRequest_forn.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_forn.onreadystatechange=function(){
      if (oHTTPRequest_forn.readyState==4){
         document.all.divForn.innerHTML = oHTTPRequest_forn.responseText;}}
       oHTTPRequest_forn.send("cd_fornecedor=" + form.cd_fornecedor.value);
 }
 
function fill_sit()
 {
     var oHTTPRequest_sit = createXMLHTTP(); 
     oHTTPRequest_sit.open("post", "../ajax/manutencao/ajax_situacao.asp", true);
     oHTTPRequest_sit.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_sit.onreadystatechange=function(){
      if (oHTTPRequest_sit.readyState==4){
         document.all.divSit.innerHTML = oHTTPRequest_sit.responseText;}}
       oHTTPRequest_sit.send("cd_situacao=" + form.cd_situacao.value);
 }
 
</script>
