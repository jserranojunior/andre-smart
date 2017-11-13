<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

busca_especialidade = request("busca_especialidade")

num_lista = 1	%>

<%If  busca_especialidade = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD</b></td>
					<td>&nbsp;<b>ESPECIALIDADE</b></td>
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
<%Elseif busca_especialidade = "todos" Then '*** Mostra campo normal ***%>
	<%strsql ="SELECT * FROM TBL_especialidades where cd_status = 1 ORDER BY nm_especialidade"
		Set rs_especialidade = dbconn.execute(strsql)
		num_lista = 1%>
	<table border="0" width="500">		
		<%Do while Not rs_especialidade.EOF
			
			cd_especialidade = rs_especialidade("cd_codigo")
			nm_especialidade = rs_especialidade("nm_especialidade")
			cd_status = rs_especialidade("cd_status")%>
	
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>ESPECIALIDADE</b></td>
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
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<!--td align="center" valign="top"><b>&nbsp;<%'=zero(num_lista)%></b><%=cd_status%></td-->
		<td valign="top"><a href="ferramentas.asp?tipo=especialidade&cd_especialidade=<%=cd_especialidade%>&nm_especialidade=<%=nm_especialidade%>&status=1&acao=editar">&nbsp;<%=cd_especialidade%></a>&nbsp;&nbsp;</td>
		<td valign="top">&nbsp;<a href="ferramentas.asp?tipo=especialidade&cd_especialidade=<%=cd_especialidade%>&nm_especialidade=<%=nm_especialidade%>&status=1&acao=editar">&nbsp;<%=nm_especialidade%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_especialidade%>')" type="button" value="A" title="Desativar"></td>		
			<%num_lista = num_lista + 1
			rs_especialidade.movenext
			loop%>
	</tr>	
</table>
<%Elseif IsNumeric (busca_especialidade) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_espec  where a057_codrac = '"&busca_especialidade&"' and a057_status = 1 ORDER BY cd_codigo "
	Set rs_especialidade = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_especialidade.EOF%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>ESPECIALIDADE</b></td>
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
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>		
		<%cd_especialidade = rs_especialidade("cd_codigo")
		nm_especialidade = rs_especialidade("nm_especialidade")
		'end if%>
		<td>&nbsp;<a href="ferramentas.asp?tipo=especialidade&cd_especialidade=<%=cd_especialidade%>&nm_especialidade=<%=nm_especialidade%>&status=1&acao=editar"><%=cd_especialidade%></a>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="ferramentas.asp?tipo=especialidade&cd_especialidade=<%=cd_especialidade%>&nm_especialidade=<%=nm_especialidade%>&status=1&acao=editar"><%=nm_especialidade%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_especialidade%>')" type="button" value="A" title="Desativar"></td>
		<%num_lista = num_lista + 1
		rs_especialidade.MoveNext
		  wend
		
		  rs_especialidade.close
		  set rs_especialidade = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (busca_especialidade) then%>
	<%xsql = "SELECT * FROM TBL_espec where nm_especialidade like '"&busca_especialidade& "%' and a057_status = 1 ORDER BY nm_especialidade"
	Set rs_especialidade = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_especialidade.EOF Then%>
		<%while not rs_especialidade.EOF%>
		<%cd_especialidade = rs_especialidade("cd_codigo")
		nm_especialidade = rs_especialidade("nm_especialidade")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>ESPECIALIDADE</b></td>
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
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td align="right"><a href="ferramentas.asp?tipo=especialidade&cd_especialidade=<%=cd_especialidade%>&nm_especialidade=<%=nm_especialidade%>&status=1&acao=editar"><%=cd_especialidade%>&nbsp;&nbsp;</a>
		<td>&nbsp;<a href="ferramentas.asp?tipo=especialidade&cd_especialidade=<%=cd_especialidade%>&nm_especialidade=<%=nm_especialidade%>&ordem=<%=ordem%>&status=1&acao=editar"><%=nm_especialidade%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_especialidade%>')" type="button" value="A" title="Desativar"></td>
	</tr>
		<%num_lista = num_lista + 1
		rs_especialidade.MoveNext
		wend
	else%>
		
	<%End if
	  rs_especialidade.close
	  set rs_especialidade = nothing%>	
</table>
<%end if%>
