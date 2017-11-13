<!-- ***************************************-->
<!-- **** Avança campos no formulario ******-->
<!-- ***************************************-->
<script language="javascript">
<!--
/* ******************************
*** Pula para o próximo campo ***
****************************** */
VerifiqueTAB=true;
function Mostra(quem, tammax) {
if ( (quem.value.length == tammax) && (VerifiqueTAB) ) { 
var i=0,j=0, indice=-1;
for (i=0; i<document.forms.length; i++) { 
for (j=0; j<document.forms[i].elements.length; j++) { 
if (document.forms[i].elements[j].name == quem.name) { 
indice=i;
break;
} 
} 
if (indice != -1) break; 
} 
for (i=0; i<=document.forms[indice].elements.length; i++) { 
if (document.forms[indice].elements[i].name == quem.name) { 
while ( (document.forms[indice].elements[(i+1)].type == "hidden") &&
(i < document.forms[indice].elements.length) ) { 
i++;
} 
document.forms[indice].elements[(i+1)].focus();
VerifiqueTAB=false;
break;
} 
} 
} 
} 
function PararTAB(quem) { VerifiqueTAB=false; } 
function ChecarTAB() { VerifiqueTAB=true; } 

/* *************************************
*** muda a cor da linha sob o cursor *** 
************************************* */
  function mOvr(src,clrOver) {
    if (!src.contains(event.fromElement)) {
	  src.style.cursor = 'hand';
	  src.bgColor = clrOver;
	}
  }
  function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	  src.style.cursor = 'default';
	  src.bgColor = clrIn;
	}
  }
/* *************************************************
*** Seleciona todos os checkbox de um formulário ***
************************************************* */
 function selecionar_tudo(){ 
   for (i=0;i<document.form.elements.length;i++)
      if(document.form.elements[i].type == "checkbox"){
        document.form.elements[i].checked=1}
	}
 function deselecionar_tudo(){ 
  for (i=0;i<document.form.elements.length;i++) 
     if(document.form.elements[i].type == "checkbox"){
         document.form.elements[i].checked=0}
	}

function foco(){
document.form.cd_patrimonio.focus(); }
// -->	
</script>