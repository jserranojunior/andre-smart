<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")
cd_user = session("cd_codigo")

busca_equipamento = request("busca_equipamento")
busca_equipamento = busca_ajax(busca_equipamento)



separador1 = instr(1,busca_equipamento,"|",1)
equipamento = mid(busca_equipamento,1,separador1-1)
status = mid(busca_equipamento,separador1+1,len(busca_equipamento)-separador1)



num_lista = 1	%>

<%If  equipamento = "" Then '*** Mostra campo normal ***%>
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
<%Elseif equipamento = "todos" Then '*** Mostra campo normal ***%>
	<%strsql ="SELECT * FROM View_equipamento WHERE cd_status = 1 ORDER BY nm_equipamento_novo"
		Set rs_equipamento = dbconn.execute(strsql)
		num_lista = 1%>
	<table border="0" width="500">		
		<%Do while Not rs_equipamento.EOF
			cd_pn = rs_equipamento("cd_pn")
			cd_equipamento = rs_equipamento("cd_codigo")
			nm_equipamento = rs_equipamento("nm_equipamento_novo")
			nm_equipamento_tipo = rs_equipamento("nm_equipamento_tipo")
			cd_tipo = rs_equipamento("cd_tipo")
			cd_status = rs_equipamento("cd_status")%>
	
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>EQUIPAMENTO</b></td>
			<td>&nbsp;<b>TIPO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="30" height="1" border="0"></td>
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">			
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<!--td align="center" valign="top"><b>&nbsp;<%'=zero(num_lista)%></b><%=cd_status%></td-->
		<td valign="top">&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/equipamento_cad.asp?tipo=equipamento&cd_pn=<%=cd_pn%>&cd_equipamento=<%=cd_equipamento%>&nm_equipamento=<%=nm_equipamento%>&cd_tipo=<%=cd_tipo%>&status=1&acao=editar','Altera_nome','width=500,height=160')" return false;">&nbsp;<%=nm_equipamento%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td>&nbsp;<%=nm_equipamento_tipo%><%'=altera%></td>
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_equipamento%>')" type="button" value="A" title="Desativar"></td>		
			<%num_lista = num_lista + 1
			rs_equipamento.movenext
			loop%>
		
	</tr>
	<!--tr><td>Todos</td></tr-->
</table>
<%Elseif IsNumeric (equipamento) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM View_equipamento where cd_codigo = '"&equipamento&"' and cd_status <> 0 ORDER BY nm_equipamento_novo"
	Set rs_equipamento = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_equipamento.EOF%>
		<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>EQUIPAMENTO</b></td>
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
	<tr  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>		
		<%cd_pn = rs_equipamento("cd_pn")
		cd_equipamento = rs_equipamento("cd_codigo")
		nm_equipamento = rs_equipamento("nm_equipamento_novo")
		nm_equipamento_tipo = rs_equipamento("nm_equipamento_tipo")
		cd_tipo = rs_equipamento("cd_tipo")
		nm_tipo = rs_equipamento("nm_equipamento_tipo")%>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/equipamento_cad.asp?tipo=equipamento&cd_pn=<%=cd_pn%>&cd_equipamento=<%=cd_equipamento%>&nm_equipamento=<%=nm_equipamento%>&cd_tipo=<%=cd_tipo%>&status=1&acao=editar','Altera_nome','width=500,height=160')" return false;"><%=nm_equipamento%></a></td>
		<td>&nbsp;<%=nm_equipamento_tipo%></td>
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_equipamento%>')" type="button" value="A" title="Desativar"></td>
		<%num_lista = num_lista + 1
		rs_equipamento.MoveNext
		
	 wend
		if num_lista = 1 then%>
			<td><b>Nenhum registro encontrado.</b></td>
		<%end if
	rs_equipamento.close
	set rs_equipamento = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (equipamento) then%>
	<%if cd_user = 46  and status = 0 then
		xsql = "SELECT * FROM View_equipamento where nm_equipamento_novo like '%"&busca_inteligente_ajax(equipamento)&"%' and cd_status = "&status&" OR nm_equipamento like '%"&busca_inteligente_ajax(equipamento)&"%' and cd_status = "&status&" ORDER BY nm_equipamento_novo"
	else
		xsql = "SELECT * FROM View_equipamento where nm_equipamento_novo like '%"&busca_inteligente_ajax(equipamento)&"%' and cd_status = "&status&" ORDER BY nm_equipamento_novo"
	end if	
	Set rs_equipamento = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_equipamento.EOF Then%>
		<%while not rs_equipamento.EOF%>
		<%cd_pn = rs_equipamento("cd_pn")
		cd_equipamento = rs_equipamento("cd_codigo")
		nm_tipo = rs_equipamento("nm_equipamento_tipo")
		nm_equipamento = rs_equipamento("nm_equipamento")
		nm_equipamento_novo = rs_equipamento("nm_equipamento_novo")
		nm_equipamento_tipo = rs_equipamento("nm_equipamento_tipo")
		
		cd_tipo = rs_equipamento("cd_tipo")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>EQUIPAMENTO</b></td>
			<%if cd_user = 46 and status = 0 then%><td>&nbsp;<b>antigo</b></td><%end if%>
			<td>&nbsp;<b>TIPO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
			<%if cd_user = 46 and status = 0 then%><td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td><%end if%>
			<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/equipamento_cad.asp?cd_equipamento=<%=cd_equipamento%>&nm_equipamento=<%=nm_equipamento%>&cd_tipo=<%=cd_tipo%>&status=1&acao=editar','Altera_nome','width=500,height=160')" return false;"><%=nm_equipamento_novo%>.</a></td>
		<%if cd_user = 46 and status = 0 then%><td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/equipamento_cad.asp?cd_equipamento=<%=cd_equipamento%>&nm_equipamento=<%=nm_equipamento%>&cd_tipo=<%=cd_tipo%>&status=1&acao=editar','Altera_nome','width=500,height=160')" return false;"><%=nm_equipamento%>.</a></td><%end if%>
		<td>&nbsp;<%=nm_equipamento_tipo%></td>
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_equipamento%>')" type="button" value="A" title="Desativar"></td>
		<!--td><%'=nm_equipamento_novo%></td-->
	</tr>
		<%num_lista = num_lista + 1
		rs_equipamento.MoveNext
		wend
	else%>
		
	<%End if
	  rs_equipamento.close
	  set rs_equipamento = nothing%>	   
</table>
<%else%>
	<table>
		<tr>
			<td>Nenhum</td>
		</tr>
	</table>
<%end if%>
<br>
<span style="color:silver;"><%=equipamento%></span>