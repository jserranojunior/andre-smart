<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

busca_procedimento = request("busca_procedimento")

num_lista = 1%>

<%If  busca_procedimento = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CRM</b></td>
					<td>&nbsp;<b>PROCEDIMENTO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>						
				</tr>
	<%end if%>
</table>
<%Elseif busca_procedimento = "todos" Then '*** Mostra campo normal ***%>
	<%strsql ="SELECT top 100 * FROM TBL_procedimento where cd_status = 1 ORDER BY nm_procedimento"
		Set rs_procedimento = dbconn.execute(strsql)
		num_lista = 1%>
	<table border="0" width="500">	
		<%Do while Not rs_procedimento.EOF
			
			cd_procedimento = rs_procedimento("cd_codigo")
			cod_amb = rs_procedimento("cd_codigo")
			nm_procedimento = rs_procedimento("nm_procedimento")
			cd_status = rs_procedimento("cd_status")%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CRM</b></td>
			<td>&nbsp;<b>PROCEDIEMENTO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/blackdot.gif" alt="" width="110" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>						
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">			
		<td valign="top" align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<!--td align="center" valign="top"><b>&nbsp;<%'=zero(num_lista)%></b><%=cd_status%></td-->
		<td valign="top"><a href="ferramentas.asp?tipo=procedimento&cd_procedimento=<%=cd_procedimento%>&cod_amb=<%=cod_amb%>&nm_procedimento=<%=nm_procedimento%>&ordem=<%=ordem%>&status=1&acao=editar"><%=cod_amb%></a>&nbsp;&nbsp;</td>
		<td valign="top"><a href="ferramentas.asp?tipo=procedimento&cd_procedimento=<%=cd_procedimento%>&cd_crm=<%=num_crm%>&nm_procedimento=<%=nm_procedimento%>&ordem=<%=ordem%>&status=1&acao=editar"><%=nm_procedimento%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_procedimento%>')" type="button" value="A" title="Desativar"></td>		
			<%num_lista = num_lista + 1
			rs_procedimento.movenext
			loop%>
	</tr>	
</table>
<%Elseif IsNumeric (busca_procedimento) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_procedimento  where cd_codigo = '"&busca_procedimento&"' and cd_status = 1 ORDER BY cd_codigo "
	Set rs_procedimento = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_procedimento.EOF%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>CRM</b></td>
			<td>&nbsp;<b>PROCEDIMENTO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>						
		</tr>
	<%end if%>
	<tr  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td valign="top" align="right"><%=num_lista%>&nbsp;&nbsp;</td>		
		<%cd_procedimento = rs_procedimento("cd_codigo")
		cod_amb = rs_procedimento("cd_codigo")
		nm_procedimento = rs_procedimento("nm_procedimento")%>
		<td><a href="ferramentas.asp?tipo=procedimento&cd_procedimento=<%=cd_procedimento%>&cod_amb=<%=cod_amb%>&nm_procedimento=<%=nm_procedimento%>&status=1&acao=editar"><%=cod_amb%></a>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="ferramentas.asp?tipo=procedimento&cd_procedimento=<%=cd_procedimento%>&cod_amb=<%=cod_amb%>&nm_procedimento=<%=nm_procedimento%>&status=1&acao=editar"><%=nm_procedimento%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_procedimento%>')" type="button" value="A" title="Desativar"></td>
		<%num_lista = num_lista + 1
		rs_procedimento.MoveNext
		med_sel = ""
		  wend
		
		  rs_procedimento.close
		  set rs_procedimento = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (busca_procedimento) then%>
	<%xsql = "SELECT * FROM TBL_procedimento where nm_procedimento like '"&busca_procedimento& "%' and cd_status = 1 ORDER BY nm_procedimento"
	Set rs_procedimento = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_procedimento.EOF Then%>
		<%while not rs_procedimento.EOF%>
		<%cd_procedimento = rs_procedimento("cd_codigo")
		cod_amb = rs_procedimento("cd_codigo")
		nm_procedimento = rs_procedimento("nm_procedimento")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CRM</b></td>
			<td>&nbsp;<b>PROCEDIEMENTO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>						
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td valign="top" align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td valign="top" align="right"><a href="ferramentas.asp?tipo=procedimento&cd_procedimento=<%=cd_procedimento%>&cod_amb=<%=cod_amb%>&nm_procedimento=<%=nm_procedimento%>&status=1&acao=editar"><%=cod_amb%>&nbsp;&nbsp;</a>
		<td valign="top">&nbsp;<a href="ferramentas.asp?tipo=procedimento&cd_procedimento=<%=cd_procedimento%>&cod_amb=<%=cod_amb%>&nm_procedimento=<%=nm_procedimento%>&status=1&acao=editar"><%=nm_procedimento%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_procedimento%>')" type="button" value="A" title="Desativar"></td>
	</tr>
		<%num_lista = num_lista + 1
		rs_procedimento.MoveNext
		wend
	else%>
		
	<%End if
	  rs_procedimento.close
	  set rs_procedimento = nothing%>	
</table>
<%end if%>
