<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

busca_empresa = request("busca_empresa")
busca_empresa = busca_ajax(busca_empresa)

num_lista = 1	%>

<%If  busca_empresa = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>NOME</b></td>
					<td>&nbsp;<b>TIPO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>				
				</tr>
	<%end if%>	
</table>
<%Elseif busca_empresa = "todos" Then '*** Mostra todos os registros ***%>
	<%xsql = "SELECT * FROM TBL_empresa where cd_status = 1 ORDER BY nm_empresa"
	Set rs_empresa = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_empresa.EOF Then%>
		<%while not rs_empresa.EOF%>
		<%cd_empresa = rs_empresa("cd_codigo")
		nm_empresa = rs_empresa("nm_empresa")
		cd_cnpj = rs_empresa("cd_cnpj")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>Empresa</b></td>
			<td>&nbsp;<b>TIPO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="280" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="120" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/empresa_cad.asp?cd_empresa=<%=cd_empresa%>&nm_empresa=<%=nm_empresa%>&cd_tipo=<%=cd_tipo%>&status=1&acao=editar','Altera_nome','width=500,height=265')" return false;"><%=nm_empresa%></a></td>
		<td>&nbsp;<%=cd_cnpj%></td>
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_empresa%>')" type="button" value="A" title="Desativar"></td>
		<!--td><%'=nm_equipamento_novo%></td-->
	</tr>
		<%num_lista = num_lista + 1
		rs_empresa.MoveNext
		wend
	else%>
		
	<%End if
	  rs_empresa.close
	  set rs_empresa = nothing%>	   
</table>
<%Elseif IsNumeric (busca_empresa) then ' *** Mostra Campo numérico ***%>
	<%xsql = "SELECT * FROM TBL_empresa where cd_codigo = '"&busca_empresa&"' AND cd_status = 1 ORDER BY nm_empresa"
	Set rs_empresa = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_empresa.EOF Then%>
		<%while not rs_empresa.EOF%>
		<%cd_empresa = rs_empresa("cd_codigo")
		nm_empresa = rs_empresa("nm_empresa")
		cd_cnpj = rs_empresa("cd_cnpj")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>Empresa</b></td>
			<td>&nbsp;<b>TIPO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="280" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="120" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/empresa_cad.asp?cd_empresa=<%=cd_empresa%>&nm_empresa=<%=nm_empresa%>&cd_tipo=<%=cd_tipo%>&status=1&acao=editar','Altera_nome','width=500,height=265')" return false;"><%=nm_empresa%></a></td>
		<td>&nbsp;<%=cd_cnpj%></td>
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_empresa%>')" type="button" value="A" title="Desativar"></td>
		<!--td><%'=nm_equipamento_novo%></td-->
	</tr>
		<%num_lista = num_lista + 1
		rs_empresa.MoveNext
		wend
	else%>
		
	<%End if
	  rs_empresa.close
	  set rs_empresa = nothing%>	   
</table>

<%Elseif not IsNumeric (busca_empresa) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_empresa where nm_empresa like '%"&busca_empresa&"%' AND cd_status = 1 ORDER BY nm_empresa"
	Set rs_empresa = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_empresa.EOF Then%>
		<%while not rs_empresa.EOF%>
		<%cd_empresa = rs_empresa("cd_codigo")
		nm_empresa = rs_empresa("nm_empresa")
		cd_cnpj = rs_empresa("cd_cnpj")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>Empresa</b></td>
			<td>&nbsp;<b>TIPO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="280" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="120" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/empresa_cad.asp?cd_empresa=<%=cd_empresa%>&nm_empresa=<%=nm_empresa%>&cd_tipo=<%=cd_tipo%>&status=1&acao=editar','Altera_nome','width=500,height=265')" return false;"><%=nm_empresa%></a></td>
		<td>&nbsp;<%=cd_cnpj%></td>
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_empresa%>')" type="button" value="A" title="Desativar"></td>
		<!--td><%'=nm_equipamento_novo%></td-->
	</tr>
		<%num_lista = num_lista + 1
		rs_empresa.MoveNext
		wend
	else%>
		
	<%End if
	  rs_empresa.close
	  set rs_empresa = nothing%>	   
</table>
<%else%>
	<table>
		<tr>
			<td>Nenhum</td>
		</tr>
	</table>
<%end if%>
<br>
<span style="color:silver;"><%=busca_equipamento%></span>