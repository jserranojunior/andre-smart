<script type="text/javascript" language="javascript">
 //função em ajax que vai buscar a página que preenche a lista
 function preenchelista()
 {
  /*----------------------------------------------------------------------------------------------*/
     // criacao do objeto XMLHTTP do arquivo ajax.js
     var oHTTPRequest = createXMLHTTP(); 
     oHTTPRequest.open("post", "../../ajax/busca/objcliente.asp", true); //enviamos para a página que faz o select do que foi digitado e traz a lista preenchida.
   // para solicitacoes utilizando o metodo post deve ser acrescentado 
   // este cabecalho HTTP
     oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
   // a funcao abaixo e executada sempre que o estado do objeto muda (onreadystatechange)
     oHTTPRequest.onreadystatechange=function(){
     // o valor 4 significa que o objeto ja completou a solicitacao
      if (oHTTPRequest.readyState==4){// abaixo o texto gerado no arquivo executa.asp e colocado no div
         document.all.divCliente.innerHTML = oHTTPRequest.responseText;}}
       oHTTPRequest.send("txtBusca=" + frm1.txtBusca.value);
  /*---------------------------------------------------------------*/
 }
</script>