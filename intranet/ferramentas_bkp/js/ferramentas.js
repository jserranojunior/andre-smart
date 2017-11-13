<script type="text/javascript" language="javascript">
function mostra_equipamento(status) {
	status = status;
	busca = Form.busca_equipamento.value;
	busca = busca.replace(/\ /gi, "+"); //substitui o espaço por "+".
	var oHTTPRequest_equip = createXMLHTTP(); 
     oHTTPRequest_equip.open("post", "ferramentas/ajax/ajax_equipamento.asp", true);
     oHTTPRequest_equip.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_equip.onreadystatechange=function(){
      if (oHTTPRequest_equip.readyState==4){
         document.all.divEquip.innerHTML = oHTTPRequest_equip.responseText;}}
       oHTTPRequest_equip.send("busca_equipamento=" + busca + "|" + status);	   
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
 
// -----------------------------------------------------------------------------------------------------
 
 function mostra_unidade(status) {
 	status = status;
	busca = Form.busca_unidade.value;
	busca = busca.replace(/\ /gi, "_"); //substitui o espaço por "+".
	var oHTTPRequest_unid = createXMLHTTP(); 
     oHTTPRequest_unid.open("post", "ferramentas/ajax/ajax_unidade.asp", true);
     oHTTPRequest_unid.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_unid.onreadystatechange=function(){
      if (oHTTPRequest_unid.readyState==4){
         document.all.divUnid.innerHTML = oHTTPRequest_unid.responseText;}}
       oHTTPRequest_unid.send("busca_unidade=" + busca + "|" + status);
 }
 
// -----------------------------------------------------------------------------------------------------

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
</script>
