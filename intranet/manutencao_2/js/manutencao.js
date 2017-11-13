<script type="text/javascript" language="javascript">
 function fill_equip()
 {
     var oHTTPRequest_equip = createXMLHTTP(); 
     oHTTPRequest_equip.open("post", "../manutencao_2/ajax/ajax_equipamentos.asp", true);
     oHTTPRequest_equip.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_equip.onreadystatechange=function(){
      if (oHTTPRequest_equip.readyState==4){
         document.all.divEquip.innerHTML = oHTTPRequest_equip.responseText;}}
       oHTTPRequest_equip.send("cd_equipamento=" + form.cd_equipamento.value);
 }

 function fill_marca()
 {
     var oHTTPRequest_marca = createXMLHTTP(); 
     oHTTPRequest_marca.open("post", "../manutencao_2/ajax/ajax_marcas.asp", true);
     oHTTPRequest_marca.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_marca.onreadystatechange=function(){
      if (oHTTPRequest_marca.readyState==4){
         document.all.divMarca.innerHTML = oHTTPRequest_marca.responseText;}}
       oHTTPRequest_marca.send("cd_marca=" + form.cd_marca.value);
 }
  
  function fill_unid()
 {
     var oHTTPRequest_unid = createXMLHTTP(); 
     oHTTPRequest_unid.open("post", "../manutencao_2/ajax/ajax_unidades.asp", true);
     oHTTPRequest_unid.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_unid.onreadystatechange=function(){
      if (oHTTPRequest_unid.readyState==4){
         document.all.divUnid.innerHTML = oHTTPRequest_unid.responseText;}}
       oHTTPRequest_unid.send("nm_unidade=" + form.nm_unidade.value);
 }
 
 function fill_forn()
 {
     var oHTTPRequest_forn = createXMLHTTP(); 
     oHTTPRequest_forn.open("post", "../manutencao_2/ajax/ajax_fornecedores.asp", true);
     oHTTPRequest_forn.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_forn.onreadystatechange=function(){
      if (oHTTPRequest_forn.readyState==4){
         document.all.divForn.innerHTML = oHTTPRequest_forn.responseText;}}
       oHTTPRequest_forn.send("cd_fornecedor=" + form.cd_fornecedor.value);
 }
 
function fill_sit()
 {
     var oHTTPRequest_sit = createXMLHTTP(); 
     oHTTPRequest_sit.open("post", "../manutencao_2/ajax/ajax_situacao.asp", true);
     oHTTPRequest_sit.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_sit.onreadystatechange=function(){
      if (oHTTPRequest_sit.readyState==4){
         document.all.divSit.innerHTML = oHTTPRequest_sit.responseText;}}
       oHTTPRequest_sit.send("cd_situacao=" + form.cd_situacao.value);
 }
 
 /* ///////////////////////////////////////////////////////////////////////////////////// */
 
 function mostra_os_natureza(item)
 {
 	 item = item;
	 cod = form.cd_codigo.value;
	 var oHTTPRequest_itens = createXMLHTTP(); 
     oHTTPRequest_itens.open("post", "../manutencao_2/ajax/ajax_os_natureza.asp", true);
     oHTTPRequest_itens.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_itens.onreadystatechange=function(){
      if (oHTTPRequest_itens.readyState==4){
         document.all.divFase0.innerHTML = oHTTPRequest_itens.responseText;}}
       //oHTTPRequest_itens.send("tipo_item=" + form.tipo_item.value);
	   oHTTPRequest_itens.send("natureza_os=" + item + ':' + cod);
 }
 
 
 function mostra_itens_m(item)
 {
 	 item = item;
	 var oHTTPRequest_itens = createXMLHTTP(); 
     oHTTPRequest_itens.open("post", "../manutencao_2/ajax/ajax_patrimonio_busca.asp", true);
     oHTTPRequest_itens.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_itens.onreadystatechange=function(){
      if (oHTTPRequest_itens.readyState==4){
         document.all.divFase1.innerHTML = oHTTPRequest_itens.responseText;}}
       //oHTTPRequest_itens.send("tipo_item=" + form.tipo_item.value);
	   oHTTPRequest_itens.send("tipo_item=" + item);
 }
 
 
 function mostra_itens_c(item)
 {
 	 item = item;
	 cod = form.cd_codigo.value;
	 var oHTTPRequest_itens = createXMLHTTP(); 
     oHTTPRequest_itens.open("post", "../manutencao_2/ajax/ajax_patrimonio_compra.asp", true);
     oHTTPRequest_itens.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_itens.onreadystatechange=function(){
      if (oHTTPRequest_itens.readyState==4){
         document.all.divFase1.innerHTML = oHTTPRequest_itens.responseText;}}
       //oHTTPRequest_itens.send("tipo_item=" + form.tipo_item.value);
	   oHTTPRequest_itens.send("tipo_item=" + item + ':' + cod);
 }
 
 
 function mostra_patrimonio(item)
 {
 	 item = item;
	 num = form.num_patrimonio.value;
     var oHTTPRequest_pat = createXMLHTTP(); 
     oHTTPRequest_pat.open("post", "../manutencao_2/ajax/ajax_patrimonio_mostra.asp", true);
     oHTTPRequest_pat.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_pat.onreadystatechange=function(){
      if (oHTTPRequest_pat.readyState==4){
         document.all.divPat_mostra.innerHTML = oHTTPRequest_pat.responseText;}}
       oHTTPRequest_pat.send("no_patrimonio=" + item +":"+ num);
 }
 
 function sel_patrimonio(cod_pat)
 {
 	 cod_pat = cod_pat;
	 var oHTTPRequest_selpat = createXMLHTTP(); 
     oHTTPRequest_selpat.open("post", "../manutencao_2/ajax/ajax_patrimonio_seleciona.asp", true);
     oHTTPRequest_selpat.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_selpat.onreadystatechange=function(){
      if (oHTTPRequest_selpat.readyState==4){
         document.all.divPat_reg.innerHTML = oHTTPRequest_selpat.responseText;}}
       oHTTPRequest_selpat.send("no_patrimonio=" + cod_pat);
	   
	   //alert('teste');
 }
/* 
function mostra_orcamento()
 {
 	 num = form.num_os_assist.value;
     var oHTTPRequest_orc = createXMLHTTP(); 
    oHTTPRequest_orc.open("post", "../manutencao_2/ajax/ajax_orcamento_mostra.asp", true);
     oHTTPRequest_orc.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_orc.onreadystatechange=function(){
      if (oHTTPRequest_orc.readyState==4){
         document.all.divorcamento.innerHTML = oHTTPRequest_orc.responseText;}}
       oHTTPRequest_orc.send("num_os_assist=" + num);
 }
 */
</script>
