<%
strpag = CInt(request("pag"))
strpag1 = 50 
strbusca = request("busca")
strtipo = request("tipo")
ordem = request("ordem")
	if ordem <> "" Then
	ordem = ordem&","
	Else
	ordem = ""
	end if

If strbusca = "" And strtipo = "" Then
strtipo = "nm_status"
strbusca = ""
End If 


If strpag = 0 Then
	strpag = 1 
End If 


	xsql = "Select * From View_funcionario Where cd_regime_trabalho=1 order by "&ordem&" nm_nome,nm_sobrenome"
	Set rs = dbconn.execute(xsql)



%>

<script language="javascript">
<!--
  function mOvr(src,clrOver) {
    if (!src.contains(event.fromElement)) {
	  src.style.cursor = 'hand';
	  src.bgColor = clrOver;
	}
  }
  function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	  src.style.cursor = 'default';
	  src.bgColor = clrIn;
	}
  }

// -->	
</script>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="30" colspan="5"><img src="px.gif" alt="" width="1" height="30" border="0"></td>
	</tr>   
	<tr>
		<td class="txt_cinza"><b>Lista &raquo; <font color="red">Funcionarios</font></b></td>
	</tr>
	<form name="busca" method="">
	<tr>
		<td class="txt_cinza" align=right></td>
	</tr> 
	</form>
	<tr>
		<td height="30"><img src="px.gif" alt="" width="1" height="30" border="0"></td>
	</tr>
	<tr>
		<td valign="top" align=center >

<% If rs.eof then%>
		<table>
		  <tr>
		   <td class="textoPadrao">
			<b>Nenhum Registro Encontrado</b>
			</td>
		  </tr>
		 </table> 
	<%else%>		
			<table width="98%" border="0" cellspacing="0" cellpadding="0" class="textoPadrao">
				<tr bgcolor=#c0c0c0>
					<td valign="top" align=center width="10%" class="textopadrao">					
						<b>Matricula</b>
					</td>
					<td valign="top" width="30%" class="textopadrao">					
						<b>Nome</b>
					</td>
					<td valign="top" width="20%" class="textopadrao">					
						<b>Sobrenome</b>
					</td>
					<td valign="top" width="20%">					
						<a href="empresa_funcionarios_lista.asp?ordem=nm_sigla"><b>Unidade</b class="textopadrao"></a>
					</td>
					<td valign="top" width="10%" class="textopadrao">					
						<b>Status</b>
					</td>	
				 </tr>
				 <tr><td>&nbsp;</td></tr>
				 <% Do While Not rs.eof 
				 
'				 IF rs("nm_status") = "Em Serviço" OR rs("nm_status") = "Em Treinamento" THEN
'				   strcor="#000000"
'				 ELSE strcor = "#D3D3D3"
'				 END IF
				
				'status = rs("cd_status")
				
				If status = "5" Then
				status = "Treinamento"
				Else
				status = ""
				End If	
				
				strcod = rs("cd_codigo")
				strunidade = rs("cd_unidades")
				nm_sigla = rs("nm_sigla")
				%>
				 <tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa_funcionarios_cad.asp?pag=<%=strpag%>&cod=<%=rs("cd_codigo")%>&tipo=<%=strtipo%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'');">
					<td valign="top" align=right>					
						<font color="<%=strcor%>"><%=rs("cd_matricula")%>&nbsp;&nbsp;&nbsp;&nbsp;</font>	
					</td>
					<td valign="top" >					
						<font color="<%=strcor%>"><%=rs("nm_nome")%></font>	
					</td>
				<td valign="top" >					
						<font color="<%=strcor%>"><%=rs("nm_sobrenome")%></font>	
					</td>
						<td valign="top">					
						<font color="<%=strcor%>"><%'=cd_unidade%><%'=nm_unidade%><%=nm_sigla%></font>
					</td>
					<%strsql ="SELECT * From View_funcionario_status where cd_funcionario = "&strcod&" order by dt_atualizacao desc"
				  	Set rs_status = dbconn.execute(strsql)
					
					do while not rs_status.EOF
						'cd_status = rs_status("cd_status")
						nm_status = rs_status("nm_status")
						'dt_atualizacao = rs_status("dt_atualizacao")
					rs_status.movenext
					loop%>
					<td valign="top" >					
						<font color="<%=strcor%>"><%=nm_status%></font>
					</td>	
				 </tr>
				 <%
				  nm_unidade = ""
				  nm_sigla = ""
				  nm_status = ""
				  
					rs.movenext
					Loop
					rs.close
					Set rs = Nothing
					%>
				<tr><td>&nbsp;</td></tr>
				<tr>
				 <td  colspan=5 width="98%" bgcolor=#c0c0c0><td>
				</tr>
				
		</table> 
  <%End if%>
					</td>
				</tr>				
				<tr>
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
				</tr> 
				<tr>
					<td height="26"><img src="px.gif" alt="" width="1" height="26" border="0"></td>
				