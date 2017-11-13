<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

busca_convenio = request("busca_convenio")

num_lista = 1	%>

<%If  busca_convenio = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD</b></td>
					<td>&nbsp;<b>CONVENIO</b></td>
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
<%Elseif busca_convenio = "todos" Then '*** Mostra campo normal ***%>
	<%strsql ="SELECT * FROM TBL_convenios where cd_status = 1 ORDER BY nm_convenio"
		Set rs_convenio = dbconn.execute(strsql)
		num_lista = 1
		Do while Not rs_convenio.EOF
			
			cd_convenio = rs_convenio("cd_codigo")
			nm_convenio = rs_convenio("nm_convenio")
			cd_status = rs_convenio("cd_status")%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>CONVÊNIO</b></td>
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
		<td align="center" valign="top"><b>&nbsp;<%'=zero(num_lista)%></b><%=cd_status%></td>
		<td valign="top"><a href="ferramentas.asp?tipo=convenio&cd_convenio=<%=cd_convenio%>&nm_convenio=<%=nm_convenio%>&status=1&acao=editar"><%=cd_convenio%></a>&nbsp;&nbsp;</td>
		<td valign="top">&nbsp;<a href="ferramentas.asp?tipo=convenio&cd_convenio=<%=cd_convenio%>&nm_convenio=<%=nm_convenio%>&status=1&acao=editar">&nbsp;<%=nm_convenio%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_convenio%>')" type="button" value="A" title="Desativar">/td>		
			<%num_lista = num_lista + 1
			rs_convenio.movenext
			loop%>
	</tr>	
</table>
<%Elseif IsNumeric (busca_convenio) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_convenios  where cd_codigo = '"&busca_convenio&"' and cd_status = 1 ORDER BY cd_codigo "
	Set rs_convenio = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_convenio.EOF%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>CONVÊNIO</b></td>
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
		<%cd_convenio = rs_convenio("cd_codigo")
		nm_convenio = rs_convenio("nm_convenio")
		'end if%>
		<td>&nbsp;<a href="ferramentas.asp?tipo=convenio&cd_convenio=<%=cd_convenio%>&nm_convenio=<%=nm_convenio%>&status=1&acao=editar"><%=cd_convenio%></a>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="ferramentas.asp?tipo=convenio&cd_convenio=<%=cd_convenio%>&nm_convenio=<%=nm_convenio%>&status=1&acao=editar"><%=nm_convenio%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_convenio%>')" type="button" value="A" title="Desativar"></td>
		<%num_lista = num_lista + 1
		rs_convenio.MoveNext
		  wend
		
		  rs_convenio.close
		  set rs_convenio = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (busca_convenio) then%>
	<%xsql = "SELECT * FROM TBL_convenios where nm_convenio like '"&busca_convenio& "%' and cd_status = 1 ORDER BY nm_convenio"
	Set rs_convenio = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_convenio.EOF Then%>
		<%while not rs_convenio.EOF%>
		<%cd_convenio = rs_convenio("cd_codigo")
		nm_convenio = rs_convenio("nm_convenio")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>CONVÊNIO</b></td>
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
		<td align="right"><a href="ferramentas.asp?tipo=convenio&cd_convenio=<%=cd_convenio%>&nm_convenio=<%=nm_convenio%>&status=1&acao=editar"><%=cd_convenio%>&nbsp;&nbsp;</a>
		<td>&nbsp;<a href="ferramentas.asp?tipo=convenio&cd_convenio=<%=cd_convenio%>&nm_convenio=<%=nm_convenio%>&ordem=<%=ordem%>&status=1&acao=editar"><%=nm_convenio%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_convenio%>')" type="button" value="A" title="Desativar"></td>
	</tr>
		<%num_lista = num_lista + 1
		rs_convenio.MoveNext
		wend
	else%>
		
	<%End if
	  rs_convenio.close
	  set rs_convenio = nothing%>	
</table>
<%end if%>
