<script type="text/javascript" language="javascript">
 function fill_equip()
 {
     var oHTTPRequest = createXMLHTTP(); 
     oHTTPRequest.open("post", "ajax_equipamentos.asp", true);
     oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest.onreadystatechange=function(){
      if (oHTTPRequest.readyState==4){
         document.all.divCliente.innerHTML = oHTTPRequest.responseText;}}
       oHTTPRequest.send("txtBusca=" + frm1.txtBusca.value);
 }
</script>