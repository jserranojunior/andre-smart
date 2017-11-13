<script type="text/javascript" language="javascript">
function mostra_equipamento() {
	busca = Form.busca_equipamento.value;
	busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	var oHTTPRequest_equip = createXMLHTTP(); 
     oHTTPRequest_equip.open("post", "ferramentas/ajax/ajax_equipamento.asp", true);
     oHTTPRequest_equip.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_equip.onreadystatechange=function(){
      if (oHTTPRequest_equip.readyState==4){
         document.all.divEquip.innerHTML = oHTTPRequest_equip.responseText;}}
       oHTTPRequest_equip.send("busca_equipamento=" + busca);
 }
 
 function verifica_nome() {
	busca = Form.nm_equipamento.value;
	busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	var oHTTPRequest_equip_novo = createXMLHTTP();
     oHTTPRequest_equip_novo.open("post", "ajax/ajax_equipamento_verifica.asp", true);
     oHTTPRequest_equip_novo.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_equip_novo.onreadystatechange=function(){
      if (oHTTPRequest_equip_novo.readyState==4){
         document.all.divEquip_novo.innerHTML = oHTTPRequest_equip_novo.responseText;}}
       oHTTPRequest_equip_novo.send("verifica_nome=" + busca);
 }
 
 function mostra_empresa() {
	busca = Form.busca_empresa.value;
	busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	var oHTTPRequest_empresa = createXMLHTTP(); 
     oHTTPRequest_empresa.open("post", "ferramentas/ajax/ajax_empresa.asp", true);
     oHTTPRequest_empresa.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_empresa.onreadystatechange=function(){
      if (oHTTPRequest_empresa.readyState==4){
         document.all.divEmpresa.innerHTML = oHTTPRequest_empresa.responseText;}}
       oHTTPRequest_empresa.send("busca_empresa=" + busca);
 }
 
 function verifica_empresa() {
	busca_empresa = Form.nm_empresa.value;
	busca_empresa = busca_empresa.replace(/\ /gi, "+"); //substitui o espaço por "+".
	var oHTTPRequest_empresa = createXMLHTTP();
     oHTTPRequest_empresa.open("post", "ajax/ajax_empresa_verifica.asp", true);
     oHTTPRequest_empresa.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_empresa.onreadystatechange=function(){
      if (oHTTPRequest_empresa.readyState==4){
         document.all.divEmpresa.innerHTML = oHTTPRequest_empresa.responseText;}}
       oHTTPRequest_empresa.send("verifica_empresa=" + busca_empresa);
 }
  
 function mostra_medico(busca_medico) {
     var oHTTPRequest_med = createXMLHTTP(); 
     oHTTPRequest_med.open("post", "../ajax/ferramentas/ajax_medico.asp", true);
     oHTTPRequest_med.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_med.onreadystatechange=function(){
      if (oHTTPRequest_med.readyState==4){
         document.all.divMed.innerHTML = oHTTPRequest_med.responseText;}}
       oHTTPRequest_med.send("busca_medico=" + Form.busca_medico.value);   
 }
 
 function mostra_fornecedor() {
	busca = Form.busca_fornecedor.value;
	busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	var oHTTPRequest_forn = createXMLHTTP(); 
     oHTTPRequest_forn.open("post", "ferramentas/ajax/ajax_fornecedor.asp", true);
     oHTTPRequest_forn.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_forn.onreadystatechange=function(){
      if (oHTTPRequest_forn.readyState==4){
         document.all.divForn.innerHTML = oHTTPRequest_forn.responseText;}}
       oHTTPRequest_forn.send("busca_fornecedor=" + busca);
 }
</script>
