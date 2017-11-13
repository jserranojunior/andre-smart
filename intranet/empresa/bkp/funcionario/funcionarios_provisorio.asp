<%
strpag = CInt(request("pag"))
strpag1 = 50 
'strbusca = request("busca")
strtipo = request("tipo")
ordem = request("ordem")
	if ordem <> "" Then
	ordem = ordem&","
	Else
	ordem = ""
	end if

'If strbusca = "" And strtipo = "" Then
'strtipo = "nm_status"
'strbusca = ""
'End If 

	

'If strpag = 0 Then
'	strpag = 1 
'End If 
hoje = month(now)&"/"&day(now)&"/"&year(now)

xsql = "Select * From tbl_tipo_contrato "'Where "&func&" order by "&ordem&" cd_regime_trabalho,nm_nome,nm_sobrenome"
	Set rs_contrato = dbconn.execute(xsql)

while not rs_contrato.EOF
	cd_regime_trab = rs_contrato("cd_codigo")
	nm_regime_trab = rs_contrato("nm_regime_trab")%>

<%'=nm_regime_trab%><br>


<%str_func = request("func")
if str_func = "ativo" or str_func = "" then
	func = "(cd_regime_trabalho = '"&cd_regime_trab&"') AND dt_desliga IS NULL OR (cd_regime_trabalho <> '"&cd_regime_trab&"') AND (DATEDIFF(day, '"&hoje&"', dt_desliga) > 0)"
	'func = "dt_desliga IS NULL OR (DATEDIFF(day, '"&hoje&"', dt_desliga) > 0)"	
else
	func = "(cd_regime_trabalho = '"&cd_regime_trab&"') AND (DATEDIFF(day, '"&hoje&"', dt_desliga) <= 0)"
	'func = "(DATEDIFF(day, '"&hoje&"', dt_desliga) <= 0)"
	
end if

	'xsql = "Select * From View_funcionario Where "&func&" order by "&ordem&" nm_nome,nm_sobrenome"<br>
	xsql = "Select * From View_funcionario Where "&func&" order by "&ordem&" cd_regime_trabalho,nm_nome,nm_sobrenome"
	Set rs = dbconn.execute(xsql)%>


<table width="500" border="0" cellspacing="0" cellpadding="0">
	<!--tr>
		<td height="30" colspan="5"><img src="px.gif" alt="" width="1" height="30" border="0"></td>
	</tr>   
	
		<td class="txt_cinza" align=right></td>
	</tr> 
	<tr>
		<td height="30"><img src="px.gif" alt="" width="1" height="30" border="0"></td>
	</tr-->
	
	<tr>
		<td valign="top" align=center >

<%' If rs.eof then%>
		<!--table>
		  <tr>
		   <td class="textoPadrao">
			<b>Nenhum Registro Encontrado</b>
			</td>
		  </tr>
		 </table--> 
	<%'else
	if not rs.eof then%>		
			<table width="98%" border="0" cellspacing="0" cellpadding="0" class="textoPadrao">
				<tr>
					<td class="txt_cinza" colspan="100%"><b>Lista &raquo; <font color="red">Colaborador <%'=Ucase(str_func)&"S"%> &nbsp; <%=nm_regime_trab%></font></b> &nbsp;<br><br>/ <a href="empresa.asp?tipo=novo">Novo</a>/</td>
				</tr>
				<tr id="no_print">
					<td colspan="100%">&nbsp;&nbsp;<!--(Listagem: <a href="empresa.asp?tipo=lista&func=ativo">Ativos</a> / <a href="empresa.asp?tipo=lista&func=inativo">Inativos</a>)
					&nbsp;&nbsp;&nbsp; (<a href="empresa.asp?tipo=novo">Novo</a>)</td-->
				</tr>
				<tr bgcolor=#c0c0c0>
					<td>&nbsp;</td>
					<td><b>Mat.</b></td>
					<td>&nbsp;<a href="empresa.asp?tipo=lista&ordem="><b>Nome</b></a></td>
					<td>&nbsp;<b>Reg. Profissional</b></td>
					<td>&nbsp;<b>Cargo</b></td>
					<td>&nbsp;<b>Horario</b></td>
					<td>&nbsp;<b>Unid.</b></td>
					<td>&nbsp;<b>Situação</b></td>					
					<td>&nbsp;<b>Admissão</b class="textopadrao"></td>
					<td>&nbsp;<b>Demissão</b></td>	
				 </tr>
				 <tr><td>&nbsp;</td></tr>
				 <%num = 1
				 
				 row_color = 1
				 
				 	
				 
				 
			Do While Not rs.eof 
			
				if row_color = 1 then
						cor_linha = "#FFFFFF"
					else
						cor_linha = "#ccffff"
					end if
				
				strcod = rs("cd_codigo")
				cd_regime_trabalho = rs("cd_regime_trabalho")
				cd_matricula = rs("cd_matricula")
				nm_nome = rs("nm_nome")
				nm_sobrenome = rs("nm_sobrenome")
				contrato = rs("dt_contratado")
					if contrato <> "" Then
						contrato = zero(day(contrato))&"/"&zero(month(contrato))&"/"&right(year(contrato),2)
					end if
				desliga = rs("dt_desliga")
					if desliga <> "" Then
						desliga = zero(day(desliga))&"/"&zero(month(desliga))&"/"&right(year(desliga),2)
					end if
				
					xsql_cargo = "Select * From View_funcionario_cargo Where cd_funcionario = "&strcod&" ORDER BY dt_atualizacao desc"
					Set rs_cargo = dbconn.execute(xsql_cargo)
						if not rs_cargo.EOF then
							nm_cargo = rs_cargo("nm_cargo")
						end if
				
					xsql = "Select * From View_funcionario_horario Where cd_funcionario = "&strcod&" order by dt_atualizacao DESC"
					Set rs_hora = dbconn.execute(xsql)
						if not rs_hora.EOF Then
							hr_entrada = zero(hour(rs_hora("hr_entrada")))&":"&zero(minute(rs_hora("hr_entrada")))
							hr_saida = zero(hour(rs_hora("hr_saida")))&":"&zero(minute(rs_hora("hr_saida")))
							nm_intervalo = rs_hora("nm_intervalo")
						end if
						
					xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&strcod&" order by dt_atualizacao DESC"
					Set rs_unid = dbconn.execute(xsql)
						if not rs_unid.EOF Then
							nm_unidade = rs_unid("nm_unidade")
							nm_sigla = rs_unid("nm_sigla")
						end if
						
					xsql = "Select * From View_funcionario_status Where cd_funcionario = "&strcod&" order by dt_atualizacao DESC"
					Set rs_status = dbconn.execute(xsql)
						if not rs_status.EOF then
							nm_status = rs_status("nm_status")
						end if
						
					cd_numreg = rs("cd_numreg")
					%>
				 <tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs("cd_codigo")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px;" bgcolor="<%=cor_linha%>">
					<td style="color:silver;""><%=zero(num)%></td>
					<td valign="middle" align=right><%=strcod%><%'=cd_matricula%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td valign="middle">&nbsp;<%=nm_nome%>&nbsp;<%=nm_sobrenome%></td>
					<td valign="middle">&nbsp;<%=cd_numreg%></td>
					<td valign="middle">&nbsp;<%=nm_cargo%></td>
					<td valign="middle">&nbsp;<%=hr_entrada%> - <%=hr_saida%></td>
					<td valign="middle">&nbsp;<%=nm_sigla%><%'=nm_unidade%></td>
					<td valign="middle">&nbsp;<%=nm_status%></td>
					<td valign="middle" align="center"><%=contrato%></td>
					<td valign="middle" align="center"><%=desliga%></td>					
				 </tr>				 
				 <%
				  nm_unidade = ""
				  nm_sigla = ""
				  nm_status = ""
				  cd_status = ""
				  strcor = ""
				  
				  nm_cargo = ""
				  hr_entrada = ""
				  hr_saida = ""
				  
				  num = num + 1
				  
				  row_color = row_color + 1
				  	if row_color > 2 then
						row_color = 1
					end if
				  
					rs.movenext
					Loop
					rs.close
					Set rs = Nothing
					%>
				<tr><td>&nbsp;</td></tr>
				<tr>
				 <td  colspan=5 width="98%" bgcolor=#c0c0c0><td>
				</tr>
				<tr>
				 	<td><img src="imagens/px.gif" alt="" width="20" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="40" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="150" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="160" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="60" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="65" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="65" height="1" border="0"></td>
				</tr>
		</table> 
  <%End if%>
					</td>
				</tr>				
				<!--tr>
					<td height="26" align=center>
					  <table width=98%	border=0 class="textoPadrao" ><tr><td align=right>
					  <%For i = 1 To strtotal
						strlink ="<a href=funcionarios_lista.asp?tipo="&strtipo&"&busca="&strbusca&"&pag="&i&">"

						If i = strpag  Then
							strlink = strlink & "<b>"&i&"</b></a>"
						else
							strlink = strlink & " "&i&" </a>"
						End If 

						response.write strlink
					  next%></td></tr></table>
					  </td>
				</tr--> 
				<tr>
					<td height="26"><img src="px.gif" alt="" width="1" height="26" border="0"></td>
<%rs_contrato.movenext
wend%>				