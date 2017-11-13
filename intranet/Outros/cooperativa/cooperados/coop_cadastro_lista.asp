<%
strpag = CInt(request("pag"))
strpag1 = 50 
strbusca = request("busca")
strtipo = request("tipo")
strcontr = request("contr")

If strbusca = "" And strtipo = "" Then
strtipo = "nm_status"
strbusca = ""
End If 


If strpag = 0 Then
	strpag = 1 
End If 

'   xsql ="up_ADM_Categoria_total"
   

If strtipo <> ""  And strbusca <> "" Then

'	xsql = "EXEC SELECT_WITH_PAGING ' cd_codigo, nm_nome, nm_sobrenome, dt_contratado, dt_nascimento, nm_email, nm_rg, nm_cpf, nm_ddd, nm_fone, nm_endereco, nr_numero,                      nm_complemento, nm_cidade, nm_cep, nm_status', 'cd_codigo','VIEW_Funcionarios_listar_geral', "&strpag&", "&strpag1&", 1, '"&strtipo&" like ''%"&strbusca&"%''','nm_nome'"  
	
'	xsql1 = "EXEC SELECT_WITH_PAGING 'count(*) as qtd_pagina', 'cd_codigo','VIEW_Funcionarios_listar_geral', '', '', 1, '"&strtipo&" Like ''%"&strbusca&"%''', ''"
'	Set rs  = dbconn.execute(xsql1)

'  If Not rs.eof Then
'		strtotal = rs("qtd_pagina") / strpag1  
'  End If
' rs.close
' Set rs = nothing
	
'Else

'    xsql = "EXEC SELECT_WITH_PAGING ' cd_codigo, nm_nome, nm_sobrenome, dt_contratado, dt_nascimento, nm_email, nm_rg, nm_cpf, nm_ddd, nm_fone, nm_endereco, nr_numero,                       nm_complemento, nm_cidade, nm_cep, nm_status', 'cd_codigo','VIEW_Funcionarios_listar_geral', "&strpag&", "&strpag1&", 1, '', 'nm_nome'"  	
'	 xsql1 = "EXEC SELECT_WITH_PAGING 'count(*) as  qtd_pagina', 'cd_codigo','VIEW_Funcionarios_listar_geral', '', '', 1, '', ''"
' 	 Set rs  = dbconn.execute(xsql1)

'   If Not rs.eof Then
'		strtotal = rs("qtd_pagina") / strpag1   
'  End If
'  rs.close
'  Set rs = Nothing
  


End If
	xsql = "Up_Beneficios_lista @cd_regime_trabalho=2"
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
		<td class="txt_cinza"><b>Lista &raquo; <font color="red">Cooperados</font></b></td>
	</tr>
	<form name="busca" method="">
	<tr>
		<td class="txt_cinza" align=right>
<table>
		  <tr>
		   <Td>
			<b>Busca: 
			<select name="tipo">
				  <option <%If strtipo = "cd_codigo" then%>Selected<%End if%> value="cd_codigo">Código</option>	
 				  <option <%If strtipo = "nm_nome" then%>Selected<%End if%> value="nm_nome">Nome</option>				
	  			  <option <%If strtipo = "nm_sobrenome" then%>Selected<%End if%> value="nm_sobrenome">Sobrenome</option>	
				   <option <%If strtipo = "dt_nascimento" then%>Selected<%End if%> value="dt_nascimento">Data Nascimento</option>
				  <option <%If strtipo = "nm_status" then%>Selected<%End if%>  value="nm_status">Status</option>				 
			</select>
		</td>
		<td>
		  <input type="text" name=busca value="<%=strbusca%>"></b> <input type="submit" value="ok"> 
		</td>
		</tr>
		<tr>
		 <td colspan=2 align=center><input type="button" onclick="document.location.href('funcionarios_lista.asp');" value="Lista Geral"></td>
		</tr>
	 </table>
	 </td>
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
						<b>Unidade</b class="textopadrao">
					</td>
					<td valign="top" width="10%" class="textopadrao">					
						<b>Status</b>
					</td>	
				 </tr>
				 <% Do While Not rs.eof 
				 
'				 IF rs("nm_status") = "Em Serviço" OR rs("nm_status") = "Em Treinamento" THEN
'				   strcor="#000000"
'				 ELSE strcor = "#D3D3D3"
'				 END IF
				
				status = rs("cd_status")
				
				If status = "5" Then
				status = "Treinamento"
				Else
				status = ""
				End If	
				
				nm_unidade = rs("nm_unidade")
				 
				 if nm_unidade <> uni Then
				 %>
				 <tr><td>&nbsp;</td></tr>
				 <%End If%>
				
				 <tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('coop_cooperados_cad.asp?pag=<%=strpag%>&cod=<%=rs("cd_codigo")%>&tipo=<%=strtipo%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'');">
					<td valign="top" align=center>					
						<font color="<%=strcor%>"><%=rs("cd_matricula")%></font>	
					</td>
					<td valign="top" >					
						<font color="<%=strcor%>"><%=rs("nm_nome")%></font>	
					</td>
				<td valign="top" >					
						<font color="<%=strcor%>"><%=rs("nm_sobrenome")%></font>	
					</td>
					<td valign="top">					
						<font color="<%=strcor%>"><%=nm_unidade%></font>
					</td>
					<td valign="top" >					
						<font color="<%=strcor%>"><%=status%></font>
					</td>	
				 </tr>
				 <%
				  uni = nm_unidade
					rs.movenext
					Loop
					rs.close
					Set rs = Nothing
					%>
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
				