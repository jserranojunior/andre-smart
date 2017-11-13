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
<%'if posi_pg > 0 then posi_pg = posi_pg + 1 end if%>
<!--div-->
<div style="page-break-before:always; border:0px solid #<%=cor%>; width:770px; height:1024px; position:absolute; top:<%=posi_pg%>pt; left:0px; background:white;">
	<div style="border:0px solid black; width:75px; height:43px; position:absolute; top:0px; left:40px;" class="relatorio_campos"><img src="../../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></div>
	<div style="border:0px solid black; width:400px; height:40px; position:absolute; top:0px; left:150;" align="center" class="relatorio_titulo"><br>RELATÓRIO TÉCNICO DE CIRÚRGIA</div>
	<div style="border:1px solid lightgray; width:150px; height:40px; position:absolute; top:0px; left:590px;" class="relatorio_campos" align="center">&nbsp;<br>&nbsp;<%=protocolo%></div>
	
	<div style="border:1px solid red; width:350px; height:82px; position:absolute; top:55px; left:30px;" class="relatorio_campos" align="center"><br><br><br><br>ETIQUETA</div>
	
	<div style="border:1px solid red; width:179px; height:20px; position:absolute; top:55px; left:561px;" class="relatorio_campos">&nbsp;Data</div>
	<div style="border:1px solid red; width:179px; height:20px; position:absolute; top:74px; left:561px;" class="relatorio_campos">&nbsp;Hora agendada</div>
	<div style="border:1px solid red; width:179px; height:20px; position:absolute; top:93px; left:561px;" class="relatorio_campos">&nbsp;Hora início</div>
	<div style="border:1px solid red; width:179px; height:20px; position:absolute; top:112px; left:561px;" class="relatorio_campos">&nbsp;Hora final</div>
	
	<div style="border:1px solid red; width:532px; height:20px; position:absolute; top:145px; left:30px;" class="relatorio_campos" align="left">&nbsp;Cirurgião</div>
	<div style="border:1px solid blue; width:179px; height:20px; position:absolute; top:145px; left:561px;" class="relatorio_campos" align="left">&nbsp;CRM</div>
	
	<div style="border:1px solid red; width:710px; height:15px; position:absolute; top:164px; left:30px; class="relatorio_campos">&nbsp;Tipo de rack: ( &nbsp; ) Geral&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; ) Gineco&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Histero&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Otorrino&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Uro&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Torax&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )Ortopedia</div>
	
	<div style="border:1px solid red; width:532px; height:15px; position:absolute; top:190px; left:30px;" class="relatorio_campos" align="left">&nbsp;Eletiva ( )&nbsp;&nbsp;&nbsp;&nbsp;Pós-Agendada ( )&nbsp;&nbsp;&nbsp;&nbsp;Urgência()</div>
	<div style="border:1px solid red; width:532px; height:15px; position:absolute; top:204px; left:30px;" class="relatorio_campos" align="left">&nbsp;Cirurgia(s) realizada(s)</div>
	<div style="border:1px solid red; width:532px; height:18px; position:absolute; top:218px; left:30px;" class="relatorio_campos" align="left">&nbsp;1</div>
	<div style="border:1px solid red; width:532px; height:18px; position:absolute; top:235px; left:30px;" class="relatorio_campos" align="left">&nbsp;2</div>
	<div style="border:1px solid red; width:532px; height:18px; position:absolute; top:252px; left:30px;" class="relatorio_campos" align="left">&nbsp;3</div>
	<div style="border:1px solid red; width:532px; height:18px; position:absolute; top:269px; left:30px;" class="relatorio_campos" align="left">&nbsp;4</div>
	
	<div style="border:1px solid blue; width:179px; height:15px; position:absolute; top:190px; left:561px;" class="relatorio_campos" align="left">&nbsp;Cód. AMB</div>
	<div style="border:1px solid blue; width:179px; height:15px; position:absolute; top:204px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid blue; width:179px; height:18px; position:absolute; top:218px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid blue; width:179px; height:18px; position:absolute; top:235px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid blue; width:179px; height:18px; position:absolute; top:252px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid blue; width:179px; height:18px; position:absolute; top:269px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
		
	<div style="border:1px solid red; width:178px; height:32px; position:absolute; top:300px; left:30px;" class="relatorio_campos" align="left">&nbsp;Sterrad<br><br>&nbsp;Data</div>
	<div style="border:1px solid red; width:178px; height:32px; position:absolute; top:300px; left:207px;" class="relatorio_campos" align="left">&nbsp;autoclave<br><br>&nbsp;Data</div>
	<div style="border:1px solid red; width:178px; height:32px; position:absolute; top:300px; left:384px;" class="relatorio_campos" align="left">&nbsp;Statin<br><br>&nbsp;Data</div>
	<div style="border:1px solid red; width:179px; height:70px; position:absolute; top:300px; left:561px;" class="relatorio_campos" align="left">&nbsp;Anti-sepsia da pele do paciente<br><br><br> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; SIM ( )&nbsp;&nbsp;Não ( )<br><br><br>&nbsp;PVPI&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;( )<br>&nbsp;Clorexidina &nbsp; &nbsp;( )</div>
	
	<div style="border:1px solid red; width:178px; height:20px; position:absolute; top:331px; left:30px;" class="relatorio_campos" align="left">&nbsp;Validade</div>
	<div style="border:1px solid red; width:178px; height:20px; position:absolute; top:331px; left:207px;" class="relatorio_campos" align="left">&nbsp;Validade</div>
	<div style="border:1px solid red; width:178px; height:20px; position:absolute; top:331px; left:384px;" class="relatorio_campos" align="left">&nbsp;Validade</div>
	
	<div style="border:1px solid red; width:178px; height:20px; position:absolute; top:350px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:178px; height:20px; position:absolute; top:350px; left:207px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:178px; height:20px; position:absolute; top:350px; left:384px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:382px; left:30px;" class="relatorio_campos" align="left">&nbsp;EQUIPAMENTOS</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:382px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:382px; left:207px;" class="relatorio_campos" align="left">&nbsp;INSTRUMENTAL</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:382px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:382px; left:384px;" class="relatorio_campos" align="left">&nbsp;Faca sachs</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:382px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:382px; left:561px;" class="relatorio_campos" align="left">&nbsp;Tesoura</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:382px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:396px; left:30px;" class="relatorio_campos" align="left">&nbsp;Rack Num.</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:396px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:396px; left:207px;" class="relatorio_campos" align="left">&nbsp;Afast. de fígado</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:396px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:396px; left:384px;" class="relatorio_campos" align="left">&nbsp;Fio bipolar</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:396px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:396px; left:561px;" class="relatorio_campos" align="left">&nbsp;Trocater 05 mm</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:396px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:410px; left:30px;" class="relatorio_campos" align="left">&nbsp;Equipamentos</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:410px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:410px; left:207px;" class="relatorio_campos" align="left">&nbsp;Agulha de verrez</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:410px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:410px; left:384px;" class="relatorio_campos" align="left">&nbsp;Fio monopolar</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:410px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:410px; left:561px;" class="relatorio_campos" align="left">&nbsp;Torcater 10 mm</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:410px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:424px; left:30px;" class="relatorio_campos" align="left">&nbsp;Óptica</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:424px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:424px; left:207px;" class="relatorio_campos" align="left">&nbsp;Apalpador</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:424px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:424px; left:384px;" class="relatorio_campos" align="left">&nbsp;Gancho</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:424px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:424px; left:561px;" class="relatorio_campos" align="left">&nbsp;Ved 5 mm - ext</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:424px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:438px; left:30px;" class="relatorio_campos" align="left">&nbsp;Instrumental</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:438px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:438px; left:207px;" class="relatorio_campos" align="left">&nbsp;Cabo de alta f.</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:438px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:438px; left:384px;" class="relatorio_campos" align="left">&nbsp;Insert</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:438px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:438px; left:561px;" class="relatorio_campos" align="left">&nbsp;Ved. 5 mm int</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:438px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:452px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:452px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:452px; left:207px;" class="relatorio_campos" align="left">&nbsp;Camisa 20</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:452px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:452px; left:384px;" class="relatorio_campos" align="left">&nbsp;Kit de irrigação</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:452px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:452px; left:561px;" class="relatorio_campos" align="left">&nbsp;Ved 10 mm ext</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:452px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:466px; left:30px;" class="relatorio_campos" align="left">&nbsp;DVD</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:466px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:466px; left:207px;" class="relatorio_campos" align="left">&nbsp;Camisa 21</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:466px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:466px; left:384px;" class="relatorio_campos" align="left">&nbsp;Mandril 10mm</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:466px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:466px; left:561px;" class="relatorio_campos" align="left">&nbsp;Ved 10 mm int</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:466px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:480px; left:30px;" class="relatorio_campos" align="left">&nbsp;Fibra Ótica</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:480px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:480px; left:207px;" class="relatorio_campos" align="left">&nbsp;Camisa 22</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:480px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:480px; left:384px;" class="relatorio_campos" align="left">&nbsp;Mandril 5mm</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:480px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:480px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:480px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:494px; left:30px;" class="relatorio_campos" align="left">&nbsp;Fonte de luz</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:494px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:494px; left:207px;" class="relatorio_campos" align="left">&nbsp;Camisa 24</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:494px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:494px; left:384px;" class="relatorio_campos" align="left">&nbsp;Manupulador ult</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:494px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:494px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:494px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:508px; left:30px;" class="relatorio_campos" align="left">&nbsp;Histeroflator</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:508px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:508px; left:207px;" class="relatorio_campos" align="left">&nbsp;Camisa 26</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:508px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:508px; left:384px;" class="relatorio_campos" align="left">&nbsp;Pinça apreenção</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:508px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:508px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:508px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:522px; left:30px;" class="relatorio_campos" align="left">&nbsp;Insuflador</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:522px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:522px; left:207px;" class="relatorio_campos" align="left">&nbsp;Camisa Diag.</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:522px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:522px; left:384px;" class="relatorio_campos" align="left">&nbsp;Pinça atrumática</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:522px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:522px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:522px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:536px; left:30px;" class="relatorio_campos" align="left">&nbsp;Mang. Silico. Ins.</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:536px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:536px; left:207px;" class="relatorio_campos" align="left">&nbsp;Cânula aspiração</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:536px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:536px; left:384px;" class="relatorio_campos" align="left">&nbsp;Pinça Babcock</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:536px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:536px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:536px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:550px; left:30px;" class="relatorio_campos" align="left">&nbsp;Mang. Silico. His.</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:550px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:550px; left:207px;" class="relatorio_campos" align="left">&nbsp;Cânula punção</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:550px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:550px; left:384px;" class="relatorio_campos" align="left">&nbsp;Pinça Biópsia</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:550px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:550px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:550px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:564px; left:30px;" class="relatorio_campos" align="left">&nbsp;Microcâmera</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:564px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:564px; left:207px;" class="relatorio_campos" align="left">&nbsp;Clipador</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:564px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:564px; left:384px;" class="relatorio_campos" align="left">&nbsp;Pinça flex corpo</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:564px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:564px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:564px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:578px; left:30px;" class="relatorio_campos" align="left">&nbsp;TV</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:578px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:578px; left:207px;" class="relatorio_campos" align="left">&nbsp;Cauterizador</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:578px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:578px; left:384px;" class="relatorio_campos" align="left">&nbsp;Pinça gras. dente</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:578px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:578px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:578px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:592px; left:30px;" class="relatorio_campos" align="left">&nbsp;Vídeo Cassete</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:592px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:592px; left:207px;" class="relatorio_campos" align="left">&nbsp;Conector Mono</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:592px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:592px; left:384px;" class="relatorio_campos" align="left">&nbsp;Pinça grasper reta</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:592px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:592px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:592px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:606px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:606px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:606px; left:207px;" class="relatorio_campos" align="left">&nbsp;Conec. Aspira</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:606px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:606px; left:384px;" class="relatorio_campos" align="left">&nbsp;Pinça Pozzi</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:606px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:606px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:606px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:620px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:620px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:620px; left:207px;" class="relatorio_campos" align="left">&nbsp;Conec. Irriga.</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:620px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:620px; left:384px;" class="relatorio_campos" align="left">&nbsp;Pinça saca mioma</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:620px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:620px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:620px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:634px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:634px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:634px; left:207px;" class="relatorio_campos" align="left">&nbsp;Elem. de trab. 1</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:634px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:634px; left:384px;" class="relatorio_campos" align="left">&nbsp;Ponte 1 via</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:634px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:634px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:634px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:648px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:648px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:648px; left:207px;" class="relatorio_campos" align="left">&nbsp;Elem. de trab. 2</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:648px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:648px; left:384px;" class="relatorio_campos" align="left">&nbsp;Ponte 2 vias</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:648px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:648px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:648px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:662px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:662px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:662px; left:207px;" class="relatorio_campos" align="left">&nbsp;Empunhadura</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:662px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:662px; left:384px;" class="relatorio_campos" align="left">&nbsp;Porta agulha</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:662px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:662px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:662px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:676px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:676px; left:129px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:676px; left:207px;" class="relatorio_campos" align="left">&nbsp;Evacuador Ellick</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:676px; left:306px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:676px; left:384px;" class="relatorio_campos" align="left">&nbsp;Redutor metálico</div>
	<div style="border:1px solid red; width:79px; height:15px; position:absolute; top:676px; left:483px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:100px; height:15px; position:absolute; top:676px; left:561px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:80px; height:15px; position:absolute; top:676px; left:660px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<div style="border:1px solid red; width:710px; height:20px; position:absolute; top:705px; left:30px;" class="relatorio_campos" align="left">&nbsp;Equipamentos e instrumental Externos</div>
	<div style="border:1px solid red; width:710px; height:20px; position:absolute; top:724px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	<div style="border:1px solid red; width:710px; height:20px; position:absolute; top:743px; left:30px;" class="relatorio_campos" align="left">&nbsp;</div>
	
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:absolute; top:775px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:absolute; top:795px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:absolute; top:815px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:absolute; top:835px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:absolute; top:855px; left:30px;" class="relatorio_campos">
	<img src="../../imagens/blackdot.gif" alt="" width="710" height="1" border="0" style="position:absolute; top:875px; left:30px;" class="relatorio_campos">
	
	<div style="border:1px solid red; width:276px; height:50px; position:absolute; top:895px; left:30px;" class="relatorio_campos" align="left">&nbsp;Circulante da Sala</div>
	<div style="border:1px solid red; width:435px; height:50px; position:absolute; top:895px; left:305px;" class="relatorio_campos" align="left">&nbsp;Gravação DVD &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; SIM ( )  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Não ( )<br><br>&nbsp;Entregue para</div>
	
	<div style="border:1px solid red; width:276px; height:40px; position:absolute; top:944px; left:30px;" class="relatorio_campos" align="left">&nbsp;Enfermagem técnica CMI</div>
	<div style="border:1px solid red; width:435px; height:79px; position:absolute; top:944px; left:305px;" class="relatorio_campos" align="left">&nbsp;Observações <br><br><br><br><br><br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Médico / Enfermeira</div>

	<div style="border:1px solid blue; width:276px; height:40px; position:absolute; top:983px; left:30px;" class="relatorio_campos" align="left">&nbsp;Enfermagem técnica CMI</div>	
	<!--div style="page-break-before:always;" align="center"> <%=posi_pg%> </div-->
</div>
<%
protocolo = protocolo + 1
cor = cor + 33
posi_pg = posi_pg + 822
quantidade = quantidade + 1
loop
%>
</body>
</html>
