	
	<!--#include file="includes/util.asp"-->
	<!--#include file="includes/javascripts.asp"-->
	<!--#include file="includes/inc_open_connection.asp"-->
	<!--#include file="includes/inc_area_restrita.asp"-->
   <html>
<head>
	<link rel="stylesheet" type="text/css" src="print.css" media="print" />
	<!--meta http-equiv="Content-Type" content="text/html;charset=utf-8" -->
	
    <title>ADM - VDLAP</title>

<!--script type="text/javascript" src="js\drop_down.js"></script-->
<!--script type="text/javascript" src="js\startlist.js"></script-->
<script type="text/javascript" src="js\menu.js"></script>
<!--script type="text/javascript" src="js\hoverpseudo.js"></script-->

<script language="javascript">
  function foco_protocolo(){  
	h_data = new Date();
			h_dia = document.calendario.dt_dia.value;
			h_mes = document.calendario.dt_mes.value;
			h_ano = document.calendario.dt_ano.value;
				
				h_data = h_dia + "/" + h_mes + "/" + h_ano;
				document.calendario.variavel.value=h_mes + "/" + h_ano;
	
	/*if (document.Form.cd_unidade.value == ''){			
			document.form.cd_unidade.focus();
			//alert("unidade");
		}
	else{
		document.Form.nm_nome.focus();
		}	*/	
	}	
</script>		


<!--#include file="css/geral.htm"-->

<style type="text/css">
	@import "css/menu2.css";
</style>
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
<style>
.right{
text-align:left;
color: red;
float:right;
margin-right:45px;

a.no_print{
    color: #0066ff;
    text-shadow: 4px 7px 8px black;
    text-decoration: none;

}
</style>
<span class="right no_print"><a class="no_print" href="http://smartnew/vdlap/">Teste a nova versão</a></span>
<body bgcolor="#FFFFFF"  leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onLoad="foco_protocolo();">
	<span style="border:0px solid red; position:absolute; width:5px; height:98%; left:0px; top:0px; margin:0px; padding:0px; background-color:gray;"" class="no_print">&nbsp;<img src="imagens/ic_print.gif" alt="imprimir" width="16" height="16" border="0" onclick="visualizarImpressao();" style="position:absolute; left:800px; top:0;" class="no_print"></span>
	<span style="border:0px solid blue; position:absolute; width:189px; left:20px; top:0px; margin:0px; padding:0px;" class="no_print">&nbsp;Bemvindo <%=session("nm_nome")%> - ção</span>
	<span style="border:0px solid orange; position:absolute; width:498px; left:210px; top:0px; margin:0px; padding:0px;" class="no_print">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <B>Esté o seu acesso nº <%=session("nr_qtd_acasso")%> , último acesso em <%=session("dt_ultimo_acesso")%>&nbsp;&nbsp;&nbsp;</b></span>	
	<span style="border:0px solid orange; position:absolute; width:99%; height:3px; left:5px; top:15px; margin:0px; padding:0px; background-color:gray;" class="no_print"><!--img src="imagens/pxc.gif" alt="" width="50" height="1" border="0" class="no_print"--></span>	
	