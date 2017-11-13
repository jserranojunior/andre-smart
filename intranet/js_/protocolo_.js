<script type="text/javascript" language="javascript">
 function fill_prot()
 {
     var oHTTPRequest_prot = createXMLHTTP(); 
     oHTTPRequest_prot.open("post", "../ajax/procedimento/ajax_procedimentos.asp", true);
     oHTTPRequest_prot.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_prot.onreadystatechange=function(){
      if (oHTTPRequest_prot.readyState==4){
         document.all.divprot.innerHTML = oHTTPRequest_prot.responseText;}}
       oHTTPRequest_prot.send("procedimentos=" + form.procedimentos.value);
 }
 
  function campo_materiais()
 {
     var oHTTPRequest_cmat = createXMLHTTP(); 
     oHTTPRequest_cmat.open("post", "../ajax/procedimento/ajax_campo_materiais.asp", true);
     oHTTPRequest_cmat.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_cmat.onreadystatechange=function(){
      if (oHTTPRequest_cmat.readyState==4){
         document.all.divCMat.innerHTML = oHTTPRequest_cmat.responseText;}}
       oHTTPRequest_cmat.send("nm_cirug_realizada=" + form.nm_cirug_realizada.value);
 }
 
 function mostra_convenio(cd_convenio)
 {
     var oHTTPRequest_conv = createXMLHTTP(); 
     oHTTPRequest_conv.open("post", "../ajax/protocolo/ajax_convenio.asp", true);
     oHTTPRequest_conv.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_conv.onreadystatechange=function(){
      if (oHTTPRequest_conv.readyState==4){
         document.all.divAssist_m1.innerHTML = oHTTPRequest_conv.responseText;}}
       oHTTPRequest_conv.send("cd_convenio=" + Form.cd_convenio.value);
 }
 
  function mostra_medico(cd_crm)
 {
     var oHTTPRequest_med = createXMLHTTP(); 
     oHTTPRequest_med.open("post", "../ajax/protocolo/ajax_medico.asp", true);
     oHTTPRequest_med.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_med.onreadystatechange=function(){
      if (oHTTPRequest_med.readyState==4){
         document.all.divMed.innerHTML = oHTTPRequest_med.responseText;}}
       oHTTPRequest_med.send("cd_crm=" + Form.cd_crm.value);
 }
 
  function mostra_rack(cd_rack)
 {
     var oHTTPRequest_rack = createXMLHTTP(); 
     oHTTPRequest_rack.open("post", "../ajax/protocolo/ajax_rack.asp", true);
     oHTTPRequest_rack.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_rack.onreadystatechange=function(){
      if (oHTTPRequest_rack.readyState==4){
         document.all.divRack.innerHTML = oHTTPRequest_rack.responseText;}}
       oHTTPRequest_rack.send("cd_rack=" + Form.cd_rack.value);
 }
 
   function mostra_espec(cd_especialidade)
 {
     var oHTTPRequest_espec = createXMLHTTP(); 
     oHTTPRequest_espec.open("post", "../ajax/protocolo/ajax_espec.asp", true);
     oHTTPRequest_espec.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_espec.onreadystatechange=function(){
      if (oHTTPRequest_espec.readyState==4){
         document.all.divEspec.innerHTML = oHTTPRequest_espec.responseText;}}
       oHTTPRequest_espec.send("cd_especialidade=" + Form.cd_especialidade.value);
 }
 
    function mostra_procedimento(cd_procedimento)
 {
     var oHTTPRequest_proc = createXMLHTTP(); 
     oHTTPRequest_proc.open("post", "../ajax/protocolo/ajax_procedimento.asp", true);
     oHTTPRequest_proc.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_proc.onreadystatechange=function(){
      if (oHTTPRequest_proc.readyState==4){
         document.all.divProc.innerHTML = oHTTPRequest_proc.responseText;}}
       oHTTPRequest_proc.send("cd_procedimento=" + Form.cd_procedimento.value);
 }
 
</script>
