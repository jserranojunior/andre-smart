<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

busca_rack = request("busca_rack")

num_lista = 1	%>

<%If  busca_rack = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD</b></td>
					<td>&nbsp;<b>RACK</b></td>
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
<%Elseif busca_rack = "todos" Then '*** Mostra campo normal ***%>
	<%strsql ="SELECT * FROM TBL_rack where a056_status = 1 ORDER BY a056_desrac"
		Set rs_rack = dbconn.execute(strsql)
		num_lista = 1%>
	<table border="0" width="500">		
		<%Do while Not rs_rack.EOF
			
			cd_rack = rs_rack("a056_codrac")
			nm_rack = rs_rack("a056_desrac")
			cd_status = rs_rack("a056_status")%>
	
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>RACK</b></td>
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
		<td valign="top"><a href="ferramentas.asp?tipo=rack&cd_rack=<%=cd_rack%>&nm_rack=<%=nm_rack%>&status=1&acao=editar">&nbsp;<%=cd_rack%></a>&nbsp;&nbsp;</td>
		<td valign="top">&nbsp;<a href="ferramentas.asp?tipo=rack&cd_rack=<%=cd_rack%>&nm_rack=<%=nm_rack%>&status=1&acao=editar">&nbsp;<%=nm_rack%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_rack%>')" type="button" value="A" title="Desativar"></td>		
			<%num_lista = num_lista + 1
			rs_rack.movenext
			loop%>
	</tr>	
</table>
<%Elseif IsNumeric (busca_rack) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_rack  where a056_codrac = '"&busca_rack&"' and a056_status = 1 ORDER BY a056_codrac "
	Set rs_rack = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_rack.EOF%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>RACK</b></td>
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
		<%cd_rack = rs_rack("a056_codrac")
		nm_rack = rs_rack("a056_desrac")
		'end if%>
		<td>&nbsp;<a href="ferramentas.asp?tipo=rack&cd_rack=<%=cd_rack%>&nm_rack=<%=nm_rack%>&status=1&acao=editar"><%=cd_rack%></a>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="ferramentas.asp?tipo=rack&cd_rack=<%=cd_rack%>&nm_rack=<%=nm_rack%>&status=1&acao=editar"><%=nm_rack%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_rack%>')" type="button" value="A" title="Desativar"></td>
		<%num_lista = num_lista + 1
		rs_rack.MoveNext
		  wend
		
		  rs_rack.close
		  set rs_rack = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (busca_rack) then%>
	<%xsql = "SELECT * FROM TBL_rack where a056_desrac like '"&busca_rack& "%' and a056_status = 1 ORDER BY a056_desrac"
	Set rs_rack = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_rack.EOF Then%>
		<%while not rs_rack.EOF%>
		<%cd_rack = rs_rack("a056_codrac")
		nm_rack = rs_rack("a056_desrac")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>RACK</b></td>
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
		<td align="right"><a href="ferramentas.asp?tipo=rack&cd_rack=<%=cd_rack%>&nm_rack=<%=nm_rack%>&status=1&acao=editar"><%=cd_rack%>&nbsp;&nbsp;</a>
		<td>&nbsp;<a href="ferramentas.asp?tipo=rack&cd_rack=<%=cd_rack%>&nm_rack=<%=nm_rack%>&ordem=<%=ordem%>&status=1&acao=editar"><%=nm_rack%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_rack%>')" type="button" value="A" title="Desativar"></td>
	</tr>
		<%num_lista = num_lista + 1
		rs_rack.MoveNext
		wend
	else%>
		
	<%End if
	  rs_rack.close
	  set rs_rack = nothing%>	
</table>
<%end if%>
