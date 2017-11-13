<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")
cd_user = session("cd_codigo")

busca_fornecedor = request("busca_fornecedor")

response.write(busca_fornecedor)
busca_fornecedor = busca_ajax(busca_fornecedor)

	separador1 = instr(1,busca_fornecedor,"|",1)
	fornecedor = mid(busca_fornecedor,1,separador1-1)
	status = mid(busca_fornecedor,separador1+1,len(busca_fornecedor)-separador1)


num_lista = 1	%>

<%If  fornecedor = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>NOME</b></td>
					<td>&nbsp;<b>TIPO</b></td>
					<td>&nbsp;<%=cd_user%></td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>				
				</tr>
	<%end if%>	
</table>
<%Elseif fornecedor = "todos" Then '*** Mostra campo normal ***%>
	<%strsql ="SELECT * FROM TBL_fornecedor WHERE cd_status = "&status&" ORDER BY nm_fornecedor"
		Set rs_fornecedor = dbconn.execute(strsql)
		num_lista = 1%>
	<table border="0" width="500">		
		<%Do while Not rs_fornecedor.EOF
			cd_fornecedor = rs_fornecedor("cd_codigo")
			nm_fornecedor = rs_fornecedor("nm_fornecedor")
			cd_status = rs_fornecedor("cd_status")%>
	
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>FORNECEDOR</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="30" height="1" border="0"></td>
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">			
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<!--td align="center" valign="top"><b>&nbsp;<%'=zero(num_lista)%></b><%=cd_status%></td-->
		<td valign="top">&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/fornecedores_cad.asp?cd_fornecedor=<%=cd_fornecedor%>&status=<%=cd_status%>&acao=editar','Altera_nome','width=500,height=300,scrollbars=1')" return false;">&nbsp;<%=nm_fornecedor%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_fornecedor%>')" type="button" value="A" title="Desativar"></td>		
			<%num_lista = num_lista + 1
			rs_fornecedor.movenext
			loop%>
		
	</tr>
	<!--tr><td>Todos</td></tr-->
</table>
<%Elseif IsNumeric (fornecedor) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_fornecedor where cd_codigo = '"&fornecedor&"' and cd_status = "&status&" ORDER BY nm_fornecedor"
	Set rs_fornecedor = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_fornecedor.EOF
	cd_status = rs_fornecedor("cd_status")%>
		<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>FORNECEDOR</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		</tr>
		<%end if%>
	<tr  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>		
		<%cd_fornecedor = rs_fornecedor("cd_codigo")
		nm_fornecedor = rs_fornecedor("nm_fornecedor")%>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/fornecedores_cad.asp?cd_fornecedor=<%=cd_fornecedor%>&status=<%=cd_status%>&acao=editar','Altera_nome','width=500,height=300,scrollbars=1')" return false;"><%=nm_fornecedor%></a></td>
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_fornecedor%>')" type="button" value="A" title="Desativar"></td>
		<%num_lista = num_lista + 1
		rs_fornecedor.MoveNext
		
	 wend
		if num_lista = 1 then%>
			<td><b>Nenhum registro encontrado.</b></td>
		<%end if
	rs_fornecedor.close
	set rs_fornecedor = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (fornecedor) then%>
	<%'if cd_user = 46  and status = 0 then
		'xsql = "SELECT * FROM TBL_fornecedor where nm_fornecedor like '%"&busca_inteligente_ajax(fornecedor)&"%' and cd_status = "&status&" ORDER BY nm_fornecedor"
	'else
		'xsql = "SELECT * FROM TBL_fornecedor where nm_fornecedor like '%"&busca_inteligente_ajax(fornecedor)&"%' and cd_status = "&status&" ORDER BY nm_fornecedor"
	'end if	
	xsql = "SELECT * FROM TBL_fornecedor where nm_fornecedor like '%"&busca_inteligente_ajax(fornecedor)&"%' and cd_status = "&status&" ORDER BY nm_fornecedor"
	Set rs_fornecedor = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_fornecedor.EOF Then%>
		<%while not rs_fornecedor.EOF%>
		<%cd_fornecedor = rs_fornecedor("cd_codigo")
		nm_fornecedor = rs_fornecedor("nm_fornecedor")
		cd_status = rs_fornecedor("cd_status")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>FORNECEDOR</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/fornecedores_cad.asp?cd_fornecedor=<%=cd_fornecedor%>&status=<%=cd_status%>&acao=editar','Altera_nome','width=500,height=300,scrollbars=1')" return false;"><%=nm_fornecedor%></a></td>
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_fornecedor%>')" type="button" value="A" title="Desativar"></td>
		<!--td><%'=nm_equipamento_novo%></td-->
	</tr>
		<%num_lista = num_lista + 1
		rs_fornecedor.MoveNext
		wend
	else%>
		
	<%End if
	  rs_fornecedor.close
	  set rs_fornecedor = nothing%>	   
</table>
<%else%>
	<table>
		<tr>
			<td>Nenhum</td>
		</tr>
	</table>
<%end if%>
<br>
<span style="color:silver;"><%=fornecedor%></span>