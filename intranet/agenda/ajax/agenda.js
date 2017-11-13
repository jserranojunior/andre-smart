<script type="text/javascript" language="javascript">
 function mostra_convenio(cd_convenio) {
     var oHTTPRequest_conv = createXMLHTTP(); 
     oHTTPRequest_conv.open("post", "../ajax/protocolo/ajax_convenio.asp", true);
     oHTTPRequest_conv.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_conv.onreadystatechange=function(){
      if (oHTTPRequest_conv.readyState==4){
         document.all.divAssist_m1.innerHTML = oHTTPRequest_conv.responseText;}}
       oHTTPRequest_conv.send("cd_convenio=" + Form.cd_convenio.value);
 }

</script>
