
<script type="text/javascript" language="javascript">
     function mostra_material(cd_material_1) {
     var oHTTPRequest_mat = createXMLHTTP(); 
     oHTTPRequest_mat.open("post", "../ajax/ajax_materiais_mostra.asp", true);
     oHTTPRequest_mat.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_mat.onreadystatechange=function(){
      if (oHTTPRequest_mat.readyState==4){
         document.all.divMat_mostra.innerHTML = oHTTPRequest_mat.responseText;}}
       oHTTPRequest_mat.send("cd_material_1=" + Form.cd_material_1.value);
 } 
 
</script>
