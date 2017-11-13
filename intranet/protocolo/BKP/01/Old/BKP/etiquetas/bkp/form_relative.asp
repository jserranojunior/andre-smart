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

'cor = 993300

do while not quantidade = qtd_inicial%>

<!--div style="page-break-before:always; border:0px solid #<%=cor%>; width:770px; height:1024px; position:relative; top:<%=posi_pg%>pt; left:0px; background:white;"-->
	<%dif_linha = 40
	mult = 2%>
	<div style="border:1px solid black; width:75px; height:43px; position:relative; top:-0px; left:40px;" class="relatorio_campos"><img src="../../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></div>
	<div style="border:1px solid black; width:400px; height:40px; position:relative; top:-<%=dif_linha%>px; left:150;" align="center" class="relatorio_titulo"><br>RELATÓRIO TÉCNICO DE CIRÚRGIA</div>
	<div style="border:1px solid lightgray; width:150px; height:40px; position:relative; top:-<%=dif_linha*mult%>px; left:590px;" class="relatorio_campos" align="center">&nbsp;<br>&nbsp;<%=protocolo%></div>
	<%dif_linha = dif_linha + 30%>
	<div style="border:1px solid red; width:350px; height:82px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="center"><br><br><br><br>ETIQUETA</div>
	<%dif_linha = dif_linha+82%>
			<div style="border:1px solid red; width:179px; height:20px; position:relative; top:-<%=dif_linha%>px; left:561px;" class="relatorio_campos">&nbsp;Data</div>
			<div style="border:1px solid red; width:179px; height:20px; position:relative; top:-<%=dif_linha+1%>px; left:561px;" class="relatorio_campos">&nbsp;Hora agendada</div>
			<div style="border:1px solid red; width:179px; height:20px; position:relative; top:-<%=dif_linha+2%>px; left:561px;" class="relatorio_campos">&nbsp;Hora início</div>
			<div style="border:1px solid red; width:179px; height:20px; position:relative; top:-<%=dif_linha+3%>px; left:561px;" class="relatorio_campos">&nbsp;Hora final<%=dif_linha+3%></div>
	<%dif_linha = dif_linha-20%>
	<div style="border:1px solid red; width:532px; height:20px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;Cirurgião</div>
			<div style="border:1px solid red; width:179px; height:20px; position:relative; top:-<%=dif_linha+20%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;CRM</div>
	<%dif_linha = dif_linha+21%>
	<div style="border:1px solid red; width:710px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" >&nbsp;Tipo de rack: ( &nbsp; ) Geral&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; ) Gineco&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Histero&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Otorrino&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Uro&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Torax&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Ortopedia</div>
	<%dif_linha = dif_linha-14%>
	<div style="border:1px solid red; width:532px; height:15px; position:relative; top:-<%=dif_linha-1%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;Eletiva ( )&nbsp;&nbsp;&nbsp;&nbsp;Pós-Agendada ( )&nbsp;&nbsp;&nbsp;&nbsp;Urgência()</div>
	<div style="border:1px solid red; width:532px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;Cirurgia(s) realizada(s)</div>
	<div style="border:1px solid red; width:532px; height:18px; position:relative; top:-<%=dif_linha+1%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;1</div>
	<div style="border:1px solid red; width:532px; height:18px; position:relative; top:-<%=dif_linha+2%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;2<%=dif_linha%></div>
	<div style="border:1px solid red; width:532px; height:18px; position:relative; top:-<%=dif_linha+3%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;3<%=dif_linha%></div>
	<div style="border:1px solid red; width:532px; height:18px; position:relative; top:-<%=dif_linha+4%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;4<%=dif_linha%></div>
	<%dif_linha = dif_linha+101%>
			<div style="border:1px solid red; width:179px; height:15px; position:relative; top:-<%=dif_linha%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;Cód. AMB</div>
			<div style="border:1px solid red; width:179px; height:15px; position:relative; top:-<%=dif_linha+1%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
			<div style="border:1px solid red; width:179px; height:18px; position:relative; top:-<%=dif_linha+2%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
			<div style="border:1px solid red; width:179px; height:18px; position:relative; top:-<%=dif_linha+3%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
			<div style="border:1px solid red; width:179px; height:18px; position:relative; top:-<%=dif_linha+4%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
			<div style="border:1px solid red; width:179px; height:18px; position:relative; top:-<%=dif_linha+5%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha-9%>
	<div style="border:1px solid red; width:178px; height:32px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;Sterrad<br><br>&nbsp;Data</div>
	<div style="border:1px solid red; width:178px; height:32px; position:relative; top:-<%=dif_linha+32%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;autoclave<br><br>&nbsp;Data</div>
	<div style="border:1px solid red; width:178px; height:32px; position:relative; top:-<%=dif_linha+64%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;Statin<br><br>&nbsp;Data</div>
	<div style="border:1px solid red; width:179px; height:70px; position:relative; top:-<%=dif_linha+96%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;Anti-sepsia da pele do paciente<br><br><br> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; SIM ( )&nbsp;&nbsp;Não ( )<br><br><br>&nbsp;PVPI&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;( )<br>&nbsp;Clorexidina &nbsp; &nbsp;( )</div>
	<%dif_linha = dif_linha+135%>
	<div style="border:1px solid red; width:178px; height:20px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;Validade</div>
	<div style="border:1px solid red; width:178px; height:20px; position:relative; top:-<%=dif_linha+20%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;Validade</div>
	<div style="border:1px solid red; width:178px; height:20px; position:relative; top:-<%=dif_linha+40%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;Validade</div>
	<%dif_linha = dif_linha+41%>
	<div style="border:1px solid red; width:178px; height:20px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:178px; height:20px; position:relative; top:-<%=dif_linha+20%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:178px; height:20px; position:relative; top:-<%=dif_linha+40%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
		
	<%dif_linha = dif_linha +25
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha +106
	mult = 2
	difer = 15%>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+difer%>px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*2)%>px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*3)%>px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*4)%>px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:relative; top:-<%=dif_linha+(difer*5)%>px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:relative; top:-<%=dif_linha+(difer*6)%>px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:relative; top:-<%=dif_linha+(difer*7)%>px; left:660px;" class="relatorio_campos" align="left">&nbsp;<%=dif_linha%></div>
	<%dif_linha = dif_linha+95%>
	<div style="border:1px solid red; width:710px; height:20px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;Equipamentos e instrumental Externos</div>
	<div style="border:1px solid red; width:710px; height:20px; position:relative; top:-<%=dif_linha+1%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:710px; height:20px; position:relative; top:-<%=dif_linha+2%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<%dif_linha = dif_linha%>
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos">
	<%dif_linha = dif_linha%>
	<div style="border:1px solid red; width:276px; height:50px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;Circulante da Sala</div>
	<div style="border:1px solid red; width:435px; height:50px; position:relative; top:-<%=dif_linha+50%>px; left:305px;" class="relatorio_campos" align="left">&nbsp;Gravação DVD &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; SIM ( )  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Não ( )<br><br>&nbsp;Entregue para</div>
	<%dif_linha = dif_linha+51%>
	<div style="border:1px solid red; width:276px; height:40px; position:relative; top:-<%=dif_linha%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;Enfermagem técnica CMI</div>
	<div style="border:1px solid red; width:435px; height:79px; position:relative; top:-<%=dif_linha+40%>px; left:305px;" class="relatorio_campos" align="left">&nbsp;Observações <br><br><br><br><br><br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Médico / Enfermeira</div>
	<div style="border:1px solid red; width:276px; height:40px; position:relative; top:-<%=dif_linha+80%>px; left:30px;" class="relatorio_campos" align="left">&nbsp;Enfermagem técnica CMI</div-->	
	<div style="page-break-before:always;" align="center"> </div>
<%
protocolo = protocolo + 1
quantidade = quantidade + 1
loop
%>
</body>
</html>
