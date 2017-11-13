<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
	<title>Untitled</title>
<style type="text/css">
<!--
.relatorio_titulo{color: #000000;font-family: verdana;font-size:10px;font-weight:bold}
.relatorio_campos{color: #000000;font-family: verdana;font-size:7px;}
-->
</style>	
	
</head>

<body>
<%
protocolo = 1
posi_pg = 0
qtd_inicial = 10
quantidade = 0

cor = 993300

do while not quantidade = qtd_inicial%>

<%if int(protocolo) = int(qtd_inicial) then page_break = "" else page_break = " page-break-after:always;" end if%>
<div>
		<div style="border:1px solid black; width:50px; height:10px; position:relative; top:0px; left:400px;"> teste</div>
		<div style="border:1px solid black; width:50px; height:10px; position:relative; top:20px; left:400px; <%=page_break%>" align="center">(<%=protocolo%>:<%=qtd_inicial%>) </div>
</div>

<%


protocolo = protocolo + 1
'cor = cor + 33
'posi_pg = posi_pg + 822
quantidade = quantidade + 1
loop
%>
</body>
</html>
