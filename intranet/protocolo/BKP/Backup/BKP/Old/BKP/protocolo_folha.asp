<head>
	<!--#include file="../includes/inc_open_connection.asp"-->
	<!--#include file="../includes/util.asp"-->
    
	<%tipo = request("tipo")
	
	
	cd_unidade = request("cd_unidade")
	cd_protocolo = request("cd_protocolo")
	cd_digito = request("cd_digito")
	
	'cd_unidade = "15"
	'cd_protocolo = "003103"
	'cd_digito = "8"
	%>

<title>Relatório Técnico de Cirurgia - VDLAP</title>
	  <LINK href="css/estilo.css " type=text/css rel=stylesheet>
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
<body>
<table style="border:1px solid silver; border-collapse:collapse; width:400px;">
	<%xsql ="SELECT * FROM TBL_protocolo WHERE A002_numpro = '"&cd_protocolo&"' AND A053_codung = '"&cd_unidade&"'"
			Set rs = dbconn.execute(xsql)
			do while Not rs.EOF
			
				cd_codigo = rs("a002_numseq")
				protocolo = rs("A002_numpro")
				cd_prot = rs("A002_numpro")
				cd_unidade = rs("A053_codung")
				nm_nome = rs("A002_pacnom")
				cd_idade = rs("A002_pacida")
				cd_registro = rs("A002_pacreg")
				cd_convenio = rs("A054_codcon")
				nm_sexo = rs("A002_pacsex")
				nm_cirug_realizada = rs("A002_limarm")
				dt_procedimento = rs("A002_datpro")
				nm_agenda = rs("A002_tipage")
				dt_hr_agenda = rs("A002_horage")
				dt_inicio = rs("A002_horini")
				dt_fim = rs("A002_horfin")
					dt_duracao = total_mes(datediff("n",dt_inicio,dt_fim))'&" ["&dt_inicio&" - "&dt_fim&"]"
				cd_med = rs("A055_codmed")
				cd_rack = rs("a056_codrac")
				cd_especialidade = rs("A057_codesp")
				'cd_procedimento = rs("cd_procedimento")
			
			rs.movenext
			loop%>
	<tr bgcolor="#808080">
		<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;RELATÓRIO TÉCNICO DE CIRURGIA</b></font></td>
	</tr>				
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr bgcolor="#f5f5f5">
		<td>&nbsp;Cód.: <b><%=cdun(cd_unidade)%>.<%=proto(cd_protocolo)%> &nbsp;<%=mensagem%></b></td>
		<%strsql ="Select * from TBL_unidades Where cd_codigo='"&cd_unidade&"'"
			Set rs_und = dbconn.execute(strsql)
			if not rs_und.EOF Then
				nm_unidade = rs_und("nm_unidade")
			end if%>			
		<td colspan="3"><b><%=nm_unidade%></b></td>
	</tr>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr bgcolor="#808080">
		<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;Identificação do Paciente</b></font></td>
	</tr>
	<tr bgcolor="#e0e0e0">
		<td align="left" colspan="3">&nbsp;Nome</td>
		<td align="left" colspan="2">&nbsp;Idade</td>
	</tr>
	<tr>
		<td align="left" colspan="3">&nbsp;<b><%=nm_nome%></b></td>
		<td align="left">&nbsp;<b><%=cd_idade%></b></td>
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
	<tr bgcolor="#e0e0e0">
		<td align="left">&nbsp;Registro</td>
		<td align="left" colspan="2">&nbsp;Convênio: </td>
		<td align="left">&nbsp;Sexo: </td>
	</tr>
	<tr>
		<td align="left">&nbsp;<b><%=cd_registro%></b>&nbsp;&nbsp;
		<td align="left" valign="top" colspan="2">&nbsp;<b><%'=cd_convenio%>
					<%strsql ="Select * from TBL_convenios Where cd_codigo='"&cd_convenio&"'"
					  		Set rs_conv = dbconn.execute(strsql)%>
							<%=rs_conv("nm_convenio")%>
							</b></td>
		<td align="left">&nbsp;<b><%if nm_sexo="M" Then%><%="Masculino"%>
						 <%elseif nm_sexo="F" Then%><%="Feminino"%>
						 <%end if%>
						 </b></td>			
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
	<tr bgcolor="#808080">
		<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;Dados Gerais</b></font></td>
	</tr>
	<tr bgcolor="#e0e0e0">
		<td>&nbsp;Cirurgia Realizada?</td>
		<td colspan="2">&nbsp;Data</td>
		<td>&nbsp;Agenda</td>
	</tr>
	<tr>
		<td>&nbsp;<b><%if nm_cirug_realizada="S" Then%><%="Sim"%>
			<%Elseif nm_cirug_realizada="N" Then%><%="Não"%>
			<%end if%></b></td>
		<td colspan="2"><b><%=dt_procedimento%></b></td>
		<td>&nbsp;<b><%if nm_agenda="A" Then%><%="A seguir"%>
			<%Elseif nm_agenda="E" Then%><%="Empréstimo"%>
			<%Elseif nm_agenda="N" Then%><%="Normal"%>
			<%Elseif nm_agenda="P" Then%><%="Pós Agendada"%>
			<%Elseif nm_agenda="U" Then%><%="Urgência"%>
			<%end if%>
			</b></td>
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td>&nbsp;Hora Agendada</td>
			<td colspan="2">&nbsp;Início</td>
			<td>&nbsp;Término</td>
		</tr>
		<%
		if dt_hr_agenda <> "" Then
			dt_min_agenda = minute(dt_hr_agenda)
			dt_hr_agenda = hour(dt_hr_agenda)
		end if
		if dt_inicio <> "" Then
			dt_min_inicio = minute(dt_inicio)
			dt_hr_inicio = hour(dt_inicio)
		end if
		if dt_fim <> "" Then
			dt_min_fim = minute(dt_fim)
			dt_hr_fim = hour(dt_fim)
		end if
		%>
		<tr>
			<td>&nbsp;<b><%=zero(dt_hr_agenda)%> :	<%=zero(dt_min_agenda)%></b></td>
			<td colspan="2">&nbsp;<b><%=dt_hr_inicio%> : <%=dt_min_inicio%></b></td>
			<td>&nbsp;<b><%=zero(dt_hr_fim)%> : <%=zero(dt_min_fim)%></b></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;Procedimento Cirúrgico</b></font></td>
		</tr>
		<tr bgcolor="#e0e0e0">
			<td colspan="100%">&nbsp;CRM &nbsp; / Médico</td>
		</tr>
		<tr>
			<td colspan="100%">&nbsp;<b><%'=cd_med%> 
							<%strsql ="Select * from TBL_medicos where a055_codmed = '"&cd_med&"'"
						  		Set rs_med = dbconn.execute(strsql)
								Do while Not rs_med.EOF%>
								<%=rs_med("a055_crmmed")%> - 
								<%=rs_med("a055_desmed")%>
								<%rs_med.movenext
						 		Loop%></b></td>							
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td>&nbsp;Rack</td>
			<td colspan="3">&nbsp;Especialidade</td>
		</tr>
		<tr>
			<td>&nbsp;<b><%'=cd_rack%>  
				<%strsql ="Select * from TBL_Rack where a056_codrac='"&cd_rack&"'"
				Set rs_rack = dbconn.execute(strsql)%>
				<%Do while Not rs_rack.EOF%><%=rs_rack("a056_desrac")%>
				<%rs_rack.movenext
				Loop%></b></td>							
			<td colspan="100%">&nbsp;<b><%'=cd_especialidade%>  
				<%strsql ="Select * from TBL_espec where cd_codigo = '"&cd_especialidade&"'"
				Set rs_espec = dbconn.execute(strsql)%>
				<%Do while Not rs_espec.EOF
				response.write(rs_espec("nm_especialidade"))%>
				<%rs_espec.movenext
				Loop%></b></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td colspan="100%">&nbsp;Procedimento</td>
		</tr>
		<tr>
			<td colspan="4">
			<b>
			<%'strsql ="Select * from TBL_protocolo_procedimento where cd_protocolo="&cd_protocolo&""
			strsql ="Select * from TBL_protocolo_procedimento where a002_numseq="&cd_protocolo&""
			Set rs_prot = dbconn.execute(strsql)
			Do while not rs_prot.EOF
			cd_procedimento = rs_prot("a058_codpro")
				'***** Substitui a vírgula pela sentença de busca *****
				'cd_procedimento = "10022, 20010, 30031"
				'cd_procedimento = replace(cd_procedimento,","," OR cd_codigo=")
				'cd_procedimento = replace(cd_procedimento,","," OR a058_codpro=")%>
				<table style="border:0px solid white; border-collapse:collapse;">
				<%strsql ="Select * from TBL_procedimento where cd_codigo="&cd_procedimento&""
				Set rs = dbconn.execute(strsql)
					Do while not rs.EOF
					cd_cod_procedimento = rs("cd_codigo")
					procedimento = rs("nm_procedimento")%>
					<tr>
						<td align="right" valign="top"><b><%=mid(cd_cod_procedimento,1,2)&"."&mid(cd_cod_procedimento,3,2)&"."&mid(cd_cod_procedimento,4,3)&"."&mid(cd_cod_procedimento,8,1)%></b>&nbsp;</td>
						<td align="left"><%=rs("nm_procedimento")&"<br>"%></td>				
					</tr>
					<%rs.movenext
					loop
			
			rs_prot.movenext
			loop%>
			</td>
		</tr>
			<tr><td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
			</tr>
			</table>
			</b></td>
		</tr>	
	</table>                 
</body>
</html>