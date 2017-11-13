
  <!--#include file="includes/util.asp"-->
  <!--#include file="includes/javascripts.asp"-->
<!--#include file="includes/inc_open_connection.asp"-->
   <!--#include file="includes/inc_area_restrita.asp"-->

  <html>
<head>
    <title>ADM - VDLAP</title>
	  <LINK href="css/estilo.css " type=text/css rel=stylesheet>
		<script type="text/javascript" src="js\menu.js"></script>
		
<style type="text/css" media="print">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.cabecalho {color: #000000;font-family: verdana;font-size:14px;}
.txt_ {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
#no_print{ 
	visibility:hidden; 
	display: none;}
#ok_print{
	visibility:visible;
	display:block;}
table{
	background-color: #ffffff; 
    border: 0px solid #cccccc;
	width: 200px;
	font-family: verdana;
	font-size:9px;
	text-decoration:none;}
.divrolagem {
	overflow: auto; 
    height: auto;
    width: auto;}
.inputs { background-color: #cdcdcd; font: 10px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; }  
-->
</style>
<style type="text/css" media="screen">
<!--
#ok_print{ visibility:hidden; display: none;}
#no_print{ visibility:visible; display:block;}
table{
	background-color: #ffffff; 
    border: 1px solid #cccccc;
	font-family: verdana;
	font-size: 9px;
	text-decoration:none;}

.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.cabecalho {color: #000000;font-family: verdana;font-size:14px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; }  
-->
</style>

<script language="JavaScript" src="js/scripts.js"></script>  
<script language="JavaScript">
function mueveReloj(){
  momentoActual = new Date()
  hora = momentoActual.getHours()
  minuto = momentoActual.getMinutes()
  segundo = momentoActual.getSeconds()
  
  str_segundo = new String (segundo)
  if (str_segundo.length == 1) 
    segundo = "0" + segundo
    
  str_minuto = new String (minuto)
  if (str_minuto.length == 1) 
    minuto = "0" + minuto

  str_hora = new String (hora)
  if (str_hora.length == 1) 
    hora = "0" + hora
    
  horaImprimible = hora + ":" + minuto + ":" + segundo
  
  document.form_reloj.reloj.value = horaImprimible
  
  setTimeout("mueveReloj()",1000)
}


Hoje = new Date() 
  
Data = Hoje.getDate() 
  
Dia = Hoje.getDay() 
  
Mes = Hoje.getMonth() 
  
Ano = Hoje.getFullYear() 
  
// 
  
if (Data<10) { 
  
Data = "0" + Data} 
  
if (Ano < 2000) { 
  
Ano = "19" + Ano} 
  
// 
  
NomeDia = new Array(7) 
  
NomeDia[0] = "Domingo" 
  
NomeDia[1] = "Segunda-feira" 
  
NomeDia[2] = "Terça-feira" 
  
NomeDia[3] = "Quarta-feira" 
  
NomeDia[4] = "Quinta-feira" 
  
NomeDia[5] = "Sexta-feira" 
  
NomeDia[6] = "Sábado" 
  
// 
  
NomeMes = new Array(12) 
  
NomeMes[0] = "Janeiro" 
  
NomeMes[1] = "Fevereiro" 
  
NomeMes[2] = "Março" 
  
NomeMes[3] = "Abril" 
  
NomeMes[4] = "Maio" 
  
NomeMes[5] = "Junho" 
  
NomeMes[6] = "Julho" 
  
NomeMes[7] = "Agosto" 
  
NomeMes[8] = "Setembro" 
  
NomeMes[9] = "Outubro" 
  
NomeMes[10] = "Novembro" 
  
NomeMes[11] = "Dezembro" 
  
// 
  
  
function MostrarData() {
  
document.write ( NomeDia[Dia] + ", " + Data + " de " + NomeMes[Mes] + " de " + Ano + " -</font>")
 }

{

</script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<script language="javascript">

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

	
</script>

<%
	strsession = session("cd_grupo")
	session_cd_usuario = session("cd_codigo")
	seccao = request("seccao")

	'modulos = session("modulos")
	cd_codigo = session_cd_usuario%>
	<!--#include file="modulos.asp"-->

</head>
<body bgcolor="#FFFFFF"  leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table width="100%" height="100%" cellspacing="0" cellpadding="0" valign="top">
    
    <form name="form_reloj">
    <tr height="100" id="no_print">
        <td width="1%" height="105" bgcolor="#BCBCBC"></td>
        <td width="15%" >&nbsp;<a href="home.asp"><img  border=0 src="imagens/Vdlap2p.gif"></a></td>
        <td width="85%"VALIGN=TOP  align="right" class="txt_header"><font size=1><SCRIPT>MostrarData(); </SCRIPT><input type="text" name="reloj" size="8" style="background-color :White; color : Black; font-family : Verdana, Arial, Helvetica; font-size : 8pt; text-align : center; border:0px solid #000000; " onfocus="window.document.form_reloj.reloj.blur()"></td>
    </tr>
	</form>
    <tr height="20" id="no_print">
        <td width="1%" bgcolor="#BCBCBC"></td>
        <td width="15%" class="txt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Bemvindo <%=session("nm_nome")%>, <%'=modulos%></td>
        <td width="85%" align="right" class="txt"><B>Esté o seu acesso nº <%=session("nr_qtd_acasso")%> , último acesso em <%=session("dt_ultimo_acesso")%>&nbsp;&nbsp;&nbsp;</b></font></td>
    </tr>
	<tr height="5" id="no_print"><td colspan= "100%" bgcolor="#c0c0c0" id="no_print"><img src="../imagens/px.gif"></td></tr>
