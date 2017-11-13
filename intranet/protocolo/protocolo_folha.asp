<head>
	<!--#include file="../includes/inc_open_connection.asp"-->
	<!--#include file="../includes/util.asp"-->
    
	<%tipo = request("tipo")
	
	
	cd_unidade = request("cd_unidade")
	cd_protocolo = request("cd_protocolo")
	cd_digito = request("cd_digito")
	
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
<!--table style="border:1px solid silver; border-collapse:collapse; width:400px;"-->
<table>
	<%xsql ="SELECT * FROM TBL_protocolo WHERE A002_numpro = '"&cd_protocolo&"' AND A053_codung = '"&cd_unidade&"'"
			Set rs = dbconn.execute(xsql)
			do while Not rs.EOF
			
				'cd_codigo = rs("a002_numseq")
				codigo = rs("cd_cod_protocolo")
				cd_cod_protocolo = rs("cd_cod_protocolo")
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
				cd_co = rs("cd_co")
				
			rs.movenext
			loop
			
			rs.close
			set rs = nothing%>
	<tr bgcolor="#808080">
		<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;RELATÓRIO TÉCNICO DE CIRURGIA</b></font></td>
	</tr>				
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr bgcolor="#f5f5f5">
		<td>&nbsp;Cód.: <b><a href="../protocolo.asp?tipo=digitacao&modalidade=alterar&cd_unidade=<%=zero(cd_unidade)%>&no_protocolo=<%=proto(cd_protocolo)%>&no_digito=<%=right(digitov(zero(cd_unidade)&"."&proto(cd_protocolo)),1)%>&action_form=protocolo/acoes/protocolo_acao.asp" target="_blank" title="."><%=digitov(zero(cd_unidade)&"."&proto(cd_protocolo))%></a> &nbsp;<%=mensagem%></b></td>
		<%strsql ="Select * from TBL_unidades Where cd_codigo='"&cd_unidade&"'"
			Set rs_und = dbconn.execute(strsql)
			if not rs_und.EOF Then
				nm_unidade = rs_und("nm_unidade")
			end if
			
			rs_und.close
			set rs_und = nothing%>			
		<td><b><%=nm_unidade%></b> <%=codigo%></td>
	</tr>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr bgcolor="#808080">
		<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;Identificação do Paciente</b></font></td>
	</tr>
	<tr bgcolor="#e0e0e0">
		<td align="left" colspan="2">&nbsp;Nome</td>
		<td align="left">&nbsp;Idade</td>
	</tr>
	<tr>
		<td align="left" colspan="2">&nbsp;<b><%=nm_nome%></b></td>
		<td align="left">&nbsp;<b><%=cd_idade%></b></td>
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
	<tr bgcolor="#e0e0e0">
		<td align="left">&nbsp;Registro</td>
		<td align="left" colspan="1">&nbsp;Convênio: </td>
		<td align="left">&nbsp;Sexo: </td>
	</tr>
	<tr>
		<td align="left">&nbsp;<b><%=cd_registro%></b>&nbsp;&nbsp;
		<td align="left" valign="top" colspan="1">&nbsp;<b>
					<%strsql ="Select * from TBL_convenios Where cd_codigo='"&cd_convenio&"'"
					Set rs_conv = dbconn.execute(strsql)%>
						<%=rs_conv("nm_convenio")%>
							
						<%rs_conv.close
						set rs_conv = nothing%></b></td>
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
		<td>&nbsp;Data</td>
		<td>&nbsp;Agenda</td>
	</tr>
	<tr>
		<td>&nbsp;<b><%if nm_cirug_realizada="S" Then%><%="Sim"%>
			<%Elseif nm_cirug_realizada="N" Then%><%="Não"%>
			<%end if%></b></td>
		<td><b><%=dt_procedimento%></b></td>
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
			<td>&nbsp;Início</td>
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
			<td>&nbsp;<b><%=dt_hr_inicio%> : <%=dt_min_inicio%></b></td>
			<td>&nbsp;<b><%=zero(dt_hr_fim)%> : <%=zero(dt_min_fim)%></b></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;Procedimento Cirúrgico</b></font></td>
		</tr>
		<tr bgcolor="#e0e0e0">
			<td colspan="3">&nbsp;CRM &nbsp; / Médico</td>
		</tr>
		<tr>
			<td colspan="3">&nbsp;<b><%'=cd_med%> 
							<%strsql ="Select * from TBL_medicos where a055_codmed = '"&cd_med&"'"
						  		Set rs_med = dbconn.execute(strsql)
								Do while Not rs_med.EOF%>
								<%=rs_med("a055_crmmed")%> - 
								<%=rs_med("a055_desmed")%>
								<%rs_med.movenext
						 		Loop
								
								rs_med.close
								set rs_med = nothing%></b></td>							
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td>&nbsp;Rack</td>
			<td>&nbsp;Especialidade</td>
			<td>&nbsp;Cirurgia realizada no CO?</td>
		</tr>
		<tr>
			<td>&nbsp;<b><%'=cd_rack%>  
				<%strsql ="Select * from TBL_Rack where a056_codrac='"&cd_rack&"'"
				Set rs_rack = dbconn.execute(strsql)%>
				<%Do while Not rs_rack.EOF%><%=rs_rack("a056_desrac")%>
				<%rs_rack.movenext
				Loop
				
				rs_rack.close
				set rs_rack = nothing%></b></td>							
			<td>&nbsp;<b><%'=cd_especialidade%>  
				<%strsql ="Select * from TBL_espec where cd_codigo = '"&cd_especialidade&"'"
				Set rs_espec = dbconn.execute(strsql)%>
				<%Do while Not rs_espec.EOF
				response.write(rs_espec("nm_especialidade"))%>
				<%rs_espec.movenext
				Loop
				
				rs_espec.close
				set rs_espec = nothing%></b></td>
			<td>&nbsp;<%if cd_co = 1 then
					response.write("<b>Sim</b>")
				Else 
					response.write("<b>Não</b>")
				end if%></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>		
		<tr bgcolor="#e0e0e0">
			<td colspan="3">&nbsp;Procedimento</td>			
		</tr>
		<tr>
			<td colspan="100%">
				<b>
				<%'strsql ="Select * from View_protocolo_procedimento where a002_numpro="&cd_protocolo&" AND a053_codung="&cd_unidade&" order by a003_propri"
				strsql ="Select * from View_protocolo_procedimento where cd_protocolo="&cd_cod_protocolo&" order by a003_propri"
				Set rs_prot = dbconn.execute(strsql)
				Do while not rs_prot.EOF
				cd_cod_procedimento = rs_prot("a058_codpro")
				nm_procedimento = rs_prot("nm_procedimento")%>
					<table style="border:0px solid white; border-collapse:collapse;">
						<%if head = 0 then%>
						<tr>
							<td align="right" valign="top" style="color:gray;">Código AMB.&nbsp;</td>
							<td align="left" valign="top" style="color:gray;">&nbsp;Procedimento</td>
						</tr>
						<%end if%>
						<tr>
							<td align="right" valign="top">&nbsp;<b><%=mid(cd_cod_procedimento,1,2)&"."&mid(cd_cod_procedimento,3,2)&"."&mid(cd_cod_procedimento,4,3)&"."&mid(cd_cod_procedimento,8,1)%></b>&nbsp;&nbsp;</td>
							<td align="left">&nbsp;<%=nm_procedimento&"<br>"%></td>				
						</tr>
						<tr>
							<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
							<td><img src="../imagens/px.gif" alt="" width="315" height="1" border="0"></td>
						</tr>
					</table>
				<%head = head + 1
				rs_prot.movenext
				loop
				
				rs_prot.close
				set rs_prot = nothing%>
			</td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>			
		<tr bgcolor="#e0e0e0">
			<td colspan="3">&nbsp;Óticas</td>
		</tr>
		<tr>
			<td colspan="100%">
			<b>
				<table style="border:0px solid white; border-collapse:collapse;">
					<%head = 0
					'strsql ="Select * from View_protocolo_patrimonio_lista where cd_protocolo="&cd_protocolo&" AND cd_unidade="&cd_unidade&" AND nm_tipo='O' order by nm_equipamento"
					strsql ="Select * from View_protocolo_patrimonio_lista where cd_protocolo="&cd_cod_protocolo&" order by nm_equipamento"
					Set rs_pat = dbconn.execute(strsql)
					Do while not rs_pat.EOF%>
				
						<%if head = 0 then%>
						<tr>
							<td align="right" valign="top" style="color:gray;">Nº Pat.&nbsp;</td>
							<td align="left" valign="top" style="color:gray;">&nbsp;Item</td>
						</tr>
						<%end if%>
						<tr>
							<td align="right" valign="top">&nbsp;<b><%=rs_pat("nm_tipo")%><%=rs_pat("no_patrimonio")%></b>&nbsp;&nbsp;</td>
							<td align="left">&nbsp;<%=rs_pat("nm_equipamento")%></td>				
						</tr>
						<tr>
							<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
							<td><img src="../imagens/px.gif" alt="" width="315" height="1" border="0"></td>
						</tr>				
					<%head = head + 1
					rs_pat.movenext
					loop
					
					rs_pat.close
					set rs_pat = nothing%>
				</table>
			</td>
		</tr>
		<%'if nm_agenda = "E" Then%>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;Empréstimos</b></font></td>
		</tr>		
		<tr bgcolor="#e0e0e0">
			<td colspan="3">&nbsp;Emprestimos</td>
		</tr>
		<tr>
			<td colspan="100%">
			<b>
				<table style="border:0px solid white; border-collapse:collapse;">
					<%head = 0
					'strsql ="Select * from View_protocolo_material where a002_numseq="&cd_protocolo&" AND a053_codung="&cd_unidade&" order by a061_nommat"
					strsql ="Select * from View_protocolo_material where cd_protocolo="&cd_cod_protocolo&" order by a061_nommat"
					Set rs_mat = dbconn.execute(strsql)
					Do while not rs_mat.EOF%>
				
						<%if head = 0 then%>
						<tr>
							<td align="right" valign="top" style="color:gray;">Qtd.&nbsp;</td>
							<td align="left" valign="top" style="color:gray;">&nbsp;Item</td>
						</tr>
						<%end if%>
						<tr>
							<td align="right" valign="top">&nbsp;<b><%=rs_mat("a010_quantidade")%></b>&nbsp;&nbsp;</td>
							<td align="left">&nbsp;<%=rs_mat("a061_nommat")%></td>				
						</tr>
						<tr>
							<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
							<td><img src="../imagens/px.gif" alt="" width="315" height="1" border="0"></td>
						</tr>
				
					<%head = head + 1
					rs_mat.movenext
					loop
					
					rs_mat.close
					set rs_mat = nothing%>
				</table>
			</td>
		</tr>
		<%'end if%>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr>
			<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="165" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="140" height="1" border="0"></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>	
	</table>                 
</body>
</html>