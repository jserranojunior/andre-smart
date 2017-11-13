
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
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>

<script language="javascript">
  
  function foco_protocolo(){  
	h_data = new Date();
			h_dia = document.calendario.dt_dia.value;
			h_mes = document.calendario.dt_mes.value;
			h_ano = document.calendario.dt_ano.value;
				
				h_data = h_dia + "/" + h_mes + "/" + h_ano;
				document.calendario.variavel.value=h_mes + "/" + h_ano;
	
	if (document.Form.cd_unidade.value == ''){			
			document.form.cd_unidade.focus();
			//alert("unidade");
		}
	else{
		document.Form.nm_nome.focus();
		}		
	}	
</script>
<%	id = request("cod")
	if id = "" Then
		id = request("id")
	end if
	
	session_cd_usuario = session("cd_codigo")
	seccao = request("seccao")
	cd_codigo = session_cd_usuario
	
	usuario_restrito = 57%>
	
	<!--include file="modulos.asp"-->
	
	<%mem_posicao = Request.ServerVariables("script_name")%>

</head>
<body bgcolor="#FFFFFF"  leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onLoad="foco_protocolo();">
<table width="100%" height="100%" cellspacing="0" cellpadding="0" border="0" valign="top">
    
  
    
	
    <tr height="1" id="no_print">
        <td width="1%" bgcolor="#BCBCBC">&nbsp;<img src="../imagens/px.gif" height="5"></td>
        <td width="10%" class="txt" align="center">&nbsp;Bemvindo <%=session("nm_nome")%> <%'=session_cd_usuario%> <%'=" - "&mem_posicao%></td>
        <td width="89%" align="center" class="txt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <B>Esté o seu acesso nº <%=session("nr_qtd_acasso")%> , último acesso em <%=session("dt_ultimo_acesso")%>&nbsp;&nbsp;&nbsp;</b></font></td>
    </tr>
	<tr height="3" id="no_print"><td colspan= "100%" bgcolor="#c0c0c0" id="no_print"><img src="../imagens/px.gif"></td></tr>
