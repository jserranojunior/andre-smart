		<tr bgcolor="#f5f5f5">
			<td colspan="100%">&nbsp;Protocolo: <b><%=cdun(cd_unidade)%>.<%=proto(cd_protocolo)%> &nbsp;<%=mensagem%></b></td>
		</tr>
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
			<td align="left" valign="top" colspan="2">&nbsp;<b><%=cd_convenio%> -
								<%strsql ="Select * from TBL_convenios Where cd_codigo='"&cd_convenio&"'"
						  		Set rs_conv = dbconn.execute(strsql)%>
								<%=rs_conv("nm_convenio")%>
								</b>
								<%rs_conv.close
								set rs_conv = nothing%></td>
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
						 		Loop
								
								rs_med.close
								set rs_med = nothing%></b></td>							
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td>&nbsp;Rack</td>
			<td colspan="3">&nbsp;Especialidade</td>
		</tr>
		<tr>
			<td>&nbsp;<b><%=cd_rack%> - 
				<%strsql ="Select * from TBL_Rack where a056_codrac='"&cd_rack&"'"
				Set rs_rack = dbconn.execute(strsql)%>
				<%Do while Not rs_rack.EOF%><%=rs_rack("a056_desrac")%>
				<%rs_rack.movenext
				Loop
				
				rs_rack.close
				set rs_rack = nothing%></b></td>							
			<td colspan="100%">&nbsp;<b><%=cd_especialidade%> - 
				<%strsql ="Select * from TBL_espec where cd_codigo = '"&cd_especialidade&"'"
				Set rs_espec = dbconn.execute(strsql)%>
				<%Do while Not rs_espec.EOF
				response.write(rs_espec("nm_especialidade"))%>
				<%rs_espec.movenext
				Loop
				
				rs_espec.close
				set rs_espec = nothing%></b></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td colspan="100%">&nbsp;Procedimento</td>
		</tr>
		</tr>
			<td colspan="100%" align="left">
				<table border="0" width="200" align="left">
				<b>
				<%'strsql ="Select * from TBL_protocolo_procedimento where cd_protocolo="&cd_protocolo&""
				strsql ="Select * from TBL_protocolo_procedimento where a002_numpro="&cd_protocolo&" AND a053_codung="&cd_unidade&""
				Set rs_prot = dbconn.execute(strsql)
				Do while not rs_prot.EOF
				cd_procedimento = rs_prot("a058_codpro")%>
					
					<%strsql ="Select * from TBL_procedimento where cd_codigo="&cd_procedimento&""
					Set rs = dbconn.execute(strsql)
						Do while not rs.EOF
						cd_cod_procedimento = rs("cd_codigo")
						procedimento = rs("nm_procedimento")%>
						<tr>
							<td align="right" valign="top"><b><%=mid(cd_cod_procedimento,1,2)&"."&mid(cd_cod_procedimento,3,2)&"."&mid(cd_cod_procedimento,4,3)&"."&mid(cd_cod_procedimento,8,1)&"<br>"%></b></td>
							<td align="left">- <%=rs("nm_procedimento")&"<br>"%></td>				
						</tr>
						<%rs.movenext
						loop
						
						rs.close
						set rs = nothing%>
						<tr><td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
							<td><img src="../imagens/px.gif" alt="" width="400" height="1" border="0"></td>
						</tr>			
					
					<%rs_prot.movenext
					loop
					
					rs_prot.close
					set rs_prot = nothing%>
				</table>
					<!--tr><td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
						<td><img src="../imagens/px.gif" alt="" width="400" height="1" border="0"></td>
					</tr>			
			</table-->
			</b></td>
		</tr>
	<%if nm_agenda = "E" Then%>
		<tr bgcolor="#e0e0e0">
			<td colspan="100%">&nbsp;Empréstimos</td>
		</tr>		
		</tr>
			<td colspan="100%">
			<b>
				<table border="0" width="200">
				<%cab_mat = 0
				strsql ="Select * from TBL_protocolo_material where a002_numseq="&cd_protocolo&" and a053_codung = "&cd_unidade&""
				Set rs_mat = dbconn.execute(strsql)
				Do while not rs_mat.EOF
				cd_material = rs_mat("a061_codmat")
				quantidade = rs_mat("a010_quantidade")%>					
					<%if cab_mat = "0" then%>
					<tr>
						<td align="right" bgcolor="#c0c0c0">Qtd.</td>
						<td align="left" bgcolor="#c0c0c0">Item</td>
					</tr><%end if%>				
					<%strsql ="Select * from TBL_material where a061_codmat="&cd_material&""
					Set rs = dbconn.execute(strsql)
						Do while not rs.EOF
						cd_cod_material = rs("a061_codmat")
						material = rs("a061_nommat")
						%>
						<tr>
							<td align="right" valign="top"><b><%=quantidade%></b></td>
							<td align="left">- <%=material&"<br>"%></td>				
						</tr>
						<%cab_mat = 1
						rs.movenext
						loop
						
						cab_mat = ""
						rs.close
						set rs = nothing%>
					<tr><td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
						<td><img src="../imagens/px.gif" alt="" width="400" height="1" border="0"></td>
					</tr>	

				<%rs_mat.movenext
				loop
			
				rs_mat.close
				set rs_mat = nothing%>
			
		<%end if%>
		
		
		<tr bgcolor="#e0e0e0">
			<td colspan="100%">&nbsp;Óticas</td>
		</tr>		
		</tr>
			<td colspan="100%">
			<b>
			<%head = 0
			strsql ="Select * from TBL_protocolo_patrimonio where cd_protocolo="&cd_protocolo&" and cd_unidade = "&cd_unidade&""
			Set rs_pat = dbconn.execute(strsql)
			Do while not rs_pat.EOF
			cd_codigo = rs_pat("cd_codigo")	%>
				<table border="0" width="200">
				<%if head = 0 then%>
				<tr>
					<td align="right" bgcolor="#c0c0c0">Nº. Patrimônio</td>
					<td align="left" bgcolor="#c0c0c0">Ótica</td>
				</tr><%end if%>				
				<%strsql ="Select * from View_protocolo_patrimonio_lista where cd_codigo="&cd_codigo&""
				Set rs = dbconn.execute(strsql)
					while not rs.EOF
					cd_patrimonio = rs("cd_patrimonio")
					nm_equipamento = rs("nm_equipamento")
					nm_tipo = rs("nm_tipo")%>
					<tr>
						<td align="right" valign="top"><b><%=nm_tipo&cd_patrimonio%></b></td>
						<td align="left"><%=nm_equipamento&"<br>"%></td>				
					</tr>
					<%head =  1
					rs.movenext
					wend
					head = head + 1
					
					rs.close
					set rs = nothing%>
		<tr><td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="400" height="1" border="0"></td>
		</tr>	

			<%rs_pat.movenext
			loop
			
			rs_pat.close
			set rs_pat = nothing%>
							
			</table>
			</b></td>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr>
			<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
			<td colspan="2"><img src="../imagens/px.gif" alt="" width="270" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>							
		</tr>		
		
		
		
		
		
		
		
		
		
		
		
		