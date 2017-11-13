		<tr bgcolor="#f5f5f5">
			<td colspan="100%">&nbsp;Protocolo: <b><%=cd_unidade%>.<%=cd_protocolo%> &nbsp;<%=mensagem%></b></td>
		</tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;Identificação do Paciente</b></font></td>
		</tr>
		<tr bgcolor="#e0e0e0">
			<td rowspan="4" align="center" valign="center"><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
			<td align="left" valign="top" colspan="1" width="100"><b>Paciente</b></td>
			<td align="left" valign="top" colspan="1" width="80"><%=cd_registro%></td>
			<td align="left" valign="top" colspan="1" width="300"><%=nm_nome%></td>
		</tr>
			<%strsql ="Select * from TBL_medicos where cd_codigo = '"&cd_med&"'"
			Set rs_med = dbconn.execute(strsql)
			Do while Not rs_med.EOF%>
		<tr bgcolor="#e0e0e0">
			<td align="left" valign="top" colspan="1"><b>Médico</b></td>
			<td align="left" valign="top" colspan="1"><%'=cd_med%><%=rs_med("cd_crm")%></td>
			<td align="left" valign="top" colspan="1"><%=rs_med("nm_medico")%></td>
		</tr>
			<%rs_med.movenext
			Loop%>
		<tr bgcolor="#e0e0e0">
			<td align="left" valign="top" colspan="1"><b>Data e Duração</b></td>
			<td align="left" valign="top" colspan="1"><%=dt_procedimento%></td>
			<td align="left" valign="top" colspan="1"><%=hour(dt_duracao)%>hs <%=minute(dt_duracao)%>min</td>
		</tr>
			<%'***** Substitui a vírgula pela sentença de busca *****
			'cd_procedimento = "10022, 20010, 30031"
			
			'cd_procedimento = replace(cd_procedimento,","," OR cd_codigo=")%>
			<%'strsql ="Select * from TBL_procedimento where cd_codigo="&cd_procedimento&""
			'Set rs = dbconn.execute(strsql)
			%>
		<tr bgcolor="#e0e0e0" bgcolor="#e0e0e0">
				<td align="left" valign="top" valign="top"><b>Procedimentos</b></td>
				<td align="left" valign="top" colspan="2">
					<table border="0">
					<%'*** Procedimentos ***
					xsql_proc ="SELECT * FROM TBL_protocolo_procedimento WHERE A002_numseq = '"&cd_codigo&"' order by a003_propri"
					Set rs_proc = dbconn.execute(xsql_proc)
					while Not rs_proc.EOF
					cd_procedimento = rs_proc("A058_codpro")%>
							
						<%cd_procedimento = replace(cd_procedimento,","," OR cd_codigo=")%>
						<%strsql ="Select * from TBL_procedimento where cd_codigo="&cd_procedimento&""
						Set rs = dbconn.execute(strsql)%>		
						<%Do while not rs.EOF%>	
						<tr>
							<td align="left" valign="top" bgcolor="#e0e0e0" width="78"><%=cd_procedimento%></td>
							<td align="left" valign="top" bgcolor="#e0e0e0" width="300"><%=rs("nm_procedimento")%></td>
						</tr>
						<%rs.movenext
						loop%>
					<%rs_proc.movenext
					wend%>
						
					</table>
				</td>
		</tr>
	</table>
</td>
</tr>