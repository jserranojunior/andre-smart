<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

busca_material = request("busca_material")

num_lista = 1	%>

<%If  busca_material = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD</b></td>
					<td>&nbsp;<b>material</b></td>
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
<%Elseif busca_material = "todos" Then '*** Mostra campo normal ***%>
	<%strsql ="SELECT * FROM TBL_material WHERE a061_status = 1 ORDER BY a061_nommat"
		Set rs_material = dbconn.execute(strsql)
		num_lista = 1%>
	<table border="0" width="500">		
		<%Do while Not rs_material.EOF
			
			cd_material = rs_material("a061_codmat")
			nm_material = rs_material("a061_nommat")
			cd_status = rs_material("a061_status")%>
	
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>MATERIAL</b></td>
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
		<td valign="top"><a href="ferramentas.asp?tipo=material&cd_material=<%=cd_material%>&nm_material=<%=nm_material%>&status=1&acao=editar">&nbsp;<%=cd_material%></a>&nbsp;&nbsp;</td>
		<td valign="top">&nbsp;<a href="ferramentas.asp?tipo=material&cd_material=<%=cd_material%>&nm_material=<%=nm_material%>&status=1&acao=editar">&nbsp;<%=nm_material%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_material%>')" type="button" value="A" title="Desativar"></td>		
			<%num_lista = num_lista + 1
			rs_material.movenext
			loop%>
	</tr>	
</table>
<%Elseif IsNumeric (busca_material) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_material  where a061_codmat = '"&busca_material&"' and a061_status = 1 ORDER BY a061_codmat"
	Set rs_material = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_material.EOF%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>material</b></td>
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
		<%cd_material = rs_material("a061_codmat")
		nm_material = rs_material("a061_nommat")
		'end if%>
		<td>&nbsp;<a href="ferramentas.asp?tipo=material&cd_material=<%=cd_material%>&nm_material=<%=nm_material%>&status=1&acao=editar"><%=cd_material%></a>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="ferramentas.asp?tipo=material&cd_material=<%=cd_material%>&nm_material=<%=nm_material%>&status=1&acao=editar"><%=nm_material%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_material%>')" type="button" value="A" title="Desativar"></td>
		<%num_lista = num_lista + 1
		rs_material.MoveNext
		  wend
		
		  rs_material.close
		  set rs_material = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (busca_material) then%>
	<%xsql = "SELECT * FROM TBL_material where a061_nommat like '"&busca_material& "%' and a061_status = 1 ORDER BY a061_nommat"
	Set rs_material = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_material.EOF Then%>
		<%while not rs_material.EOF%>
		<%cd_material = rs_material("a061_codmat")
		nm_material = rs_material("a061_nommat")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD</b></td>
			<td>&nbsp;<b>MATERIAIS</b></td>
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
		<td align="right"><a href="ferramentas.asp?tipo=material&cd_material=<%=cd_material%>&nm_material=<%=nm_material%>&status=1&acao=editar"><%=cd_material%>&nbsp;&nbsp;</a>
		<td>&nbsp;<a href="ferramentas.asp?tipo=material&cd_material=<%=cd_material%>&nm_material=<%=nm_material%>&ordem=<%=ordem%>&status=1&acao=editar"><%=nm_material%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_material%>')" type="button" value="A" title="Desativar"></td>
	</tr>
		<%num_lista = num_lista + 1
		rs_material.MoveNext
		wend
	else%>
		
	<%End if
	  rs_material.close
	  set rs_material = nothing%>	
</table>
<%end if%>
