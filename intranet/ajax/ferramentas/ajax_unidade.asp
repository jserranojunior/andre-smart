<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

busca_unidade = request("busca_unidade")

num_lista = 1%>

<%If  busca_unidade = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>SIGLA</b></td>
					<td>&nbsp;<b>UNIDADE</b></td>
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
<%Elseif busca_unidade = "todos" Then '*** Mostra campo normal ***%>
	<%strsql ="SELECT top 100 * FROM TBL_unidades where cd_status >= 1 ORDER BY nm_unidade"
		Set rs_unidade = dbconn.execute(strsql)
		num_lista = 1%>
	<table border="0" width="500">	
		<%Do while Not rs_unidade.EOF
			
			cd_unidade = rs_unidade("cd_codigo")
			nm_sigla = rs_unidade("nm_sigla")
			nm_unidade = rs_unidade("nm_unidade")
			cd_status = rs_unidade("cd_status")%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>SIGLA</b></td>
			<td>&nbsp;<b>UNIDADE</b></td>
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
		<td valign="top"><a href="ferramentas.asp?tipo=unidade&cd_unidade=<%=cd_unidade%>&nm_sigla=<%=nm_sigla%>&nm_unidade=<%=nm_unidade%>&ordem=<%=ordem%>&status=1&acao=editar"><%=nm_sigla%></a>&nbsp;&nbsp;</td>
		<td valign="top"><a href="ferramentas.asp?tipo=unidade&cd_unidade=<%=cd_unidade%>&nm_sigla=<%=nm_sigla%>&nm_unidade=<%=nm_unidade%>&ordem=<%=ordem%>&status=1&acao=editar"><%=nm_unidade%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_unidade%>')" type="button" value="A" title="Desativar"></td>		
			<%num_lista = num_lista + 1
			rs_unidade.movenext
			loop%>
	</tr>	
</table>
<%Elseif IsNumeric (busca_unidade) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_unidades  where cd_codigo = '"&busca_unidade&"' and cd_status >= 1 ORDER BY cd_codigo "
	Set rs_unidade = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_unidade.EOF%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>SIGLA</b></td>
			<td>&nbsp;<b>UNIDADE</b></td>
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
		<%cd_unidade = rs_unidade("cd_codigo")
		nm_sigla = rs_unidade("nm_sigla")
		nm_unidade = rs_unidade("nm_unidade")%>
		<td><a href="ferramentas.asp?tipo=unidade&cd_unidade=<%=cd_unidade%>&nm_sigla=<%=nm_sigla%>&nm_unidade=<%=nm_unidade%>&status=1&acao=editar"><%=nm_sigla%></a>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="ferramentas.asp?tipo=unidade&cd_unidade=<%=cd_unidade%>&nm_sigla=<%=nm_sigla%>&nm_unidade=<%=nm_unidade%>&status=1&acao=editar"><%=nm_unidade%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_unidade%>')" type="button" value="A" title="Desativar"></td>
		<%num_lista = num_lista + 1
		rs_unidade.MoveNext
		med_sel = ""
		  wend
		
		  rs_unidade.close
		  set rs_unidade = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (busca_unidade) then%>
	<%xsql = "SELECT * FROM TBL_unidades where nm_unidade like '"&busca_unidade& "%' and cd_status >= 1 OR nm_sigla like '"&busca_unidade& "%' and cd_status = 1 ORDER BY nm_unidade"
	Set rs_unidade = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_unidade.EOF Then%>
		<%while not rs_unidade.EOF%>
		<%cd_unidade = rs_unidade("cd_codigo")
		nm_sigla = rs_unidade("nm_sigla")
		nm_unidade = rs_unidade("nm_unidade")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>SIGLA</b></td>
			<td>&nbsp;<b>UNIDADE</b></td>
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
		<td valign="top" align="right"><a href="ferramentas.asp?tipo=unidade&cd_unidade=<%=cd_unidade%>&nm_sigla=<%=nm_sigla%>&nm_unidade=<%=nm_unidade%>&status=1&acao=editar"><%=nm_sigla%>&nbsp;&nbsp;</a>
		<td valign="top">&nbsp;<a href="ferramentas.asp?tipo=unidade&cd_unidade=<%=cd_unidade%>&nm_sigla=<%=nm_sigla%>&nm_unidade=<%=nm_unidade%>&status=1&acao=editar"><%=nm_unidade%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_unidade%>')" type="button" value="A" title="Desativar"></td>
	</tr>
		<%num_lista = num_lista + 1
		rs_unidade.MoveNext
		wend
	else%>
		
	<%End if
	  rs_unidade.close
	  set rs_unidade = nothing%>	
</table>
<%end if%>
