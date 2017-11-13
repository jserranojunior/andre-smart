<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%busca_unidade = request("busca_unidade")
separador1 = instr(1,busca_unidade,"|",1)
unidade = mid(busca_unidade,1,separador1-1)
status = mid(busca_unidade,separador1+1,len(busca_unidade)-separador1)



num_lista = 1	%>

<%If  unidade = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD.</b></td>
					<td>&nbsp;<b>UNIDADE</b></td>
					<td>&nbsp;<b>TIPO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>				
				</tr>
	<%end if%>	
</table>
<%Elseif unidade = "todos" Then '*** Mostra campo normal ***%>
	<%'strsql ="SELECT * FROM View_equipamento WHERE cd_status = 1 ORDER BY nm_equipamento_novo"
		'Set rs_equipamento = dbconn.execute(strsql)
	xsql = "up_unidade_lista @cd_status='="&status&"', @cd_hospital=' between 1 and  3'"
	set rs_unid = dbconn.execute(xsql)
		num_lista = 1%>
	<table border="0" width="500">		
		<%Do while Not rs_unid.EOF
			cd_unidade = rs_unid("cd_codigo")
			nm_unidade = rs_unid("nm_unidade")
			cd_status = rs_unid("cd_status")%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD.</b></td>
			<td>&nbsp;<b>UNIDADE</b></td>
			<td>&nbsp;<b>TIPO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="25" height="1" border="0"></td>	
			<td><img src="../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="60" height="1" border="0"></td>
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">			
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td align="center" valign="top">&nbsp;<%=zero(cd_unidade)%></td>
		<td valign="top">&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/unidade_cad.asp?cd_unidade=<%=cd_unidade%>&status=<%=cd_status%>&acao=editar&jan=1','Altera_nome','width=500,height=400')" return false;">&nbsp;<%=nm_unidade%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td>&nbsp;<%=nm_equipamento_tipo%><%'=altera%></td>
		<td id="no_print" valign="top" align="center">
			<%if status = "1" Then%>
				<img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_unidade%>')" type="button" value="A" title="Desativar">
			<%else%>
				<img src="imagens/ic_del.gif" border="0" onclick="javascript:JsDelete('<%=cd_unidade%>')" type="button" value="A" title="Excluir"> &nbsp; &nbsp;
				<img src="imagens/undo2.gif" border="0" onclick="javascript:JsReativar('<%=cd_unidade%>')" type="button" value="A" title="Reativar">
			<%end if%></td>		
			<%num_lista = num_lista + 1
			rs_unid.movenext
			loop%>
		
	</tr>
	<!--tr><td>Todos</td></tr-->
</table>
<%Elseif IsNumeric (unidade) then ' *** Mostra Campo não numérico ***%>
	<%xsql = "up_unidade_busca @cd_unidade='"&unidade&"', @nm_unidade=' ', @cd_status="&status&""
	set rs_unid = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_unid.EOF%>
		<%
		cd_unidade = rs_unid("cd_codigo")
		cd_status = rs_unid("cd_status")
		if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD.</b></td>
			<td>&nbsp;<b>UNIDADE</b></td>
			<td>&nbsp;<b>TIPO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
		</tr>
		<%end if%>
	<tr  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td align="right"><%=zero(cd_unidade)%>&nbsp;&nbsp;</td>		
		<%cd_unidade = rs_unid("cd_codigo")
		nm_unidade = rs_unid("nm_unidade")%>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/unidade_cad.asp?cd_unidade=<%=cd_unidade%>&status=<%=cd_status%>&acao=editar&jan=1','Altera_nome','width=500,height=400')" return false;"><%=nm_unidade%></a></td>
		<td>&nbsp;<%=nm_equipamento_tipo%></td>
		<td id="no_print" valign="top" align="center">
			<%if status = "1" Then%>
				<img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_unidade%>')" type="button" value="A" title="Desativar">
			<%else%>
				<img src="imagens/ic_del.gif" border="0" onclick="javascript:JsDelete('<%=cd_unidade%>')" type="button" value="A" title="Excluir"> &nbsp; &nbsp;
				<img src="imagens/undo2.gif" border="0" onclick="javascript:JsReativar('<%=cd_unidade%>')" type="button" value="A" title="Reativar">
			<%end if%></td>
		<%num_lista = num_lista + 1
		rs_unid.MoveNext
		
	 wend
		if num_lista = 1 then%>
			<td><b>Nenhum registro encontrado.</b></td>
		<%end if%>
		<%rs_unid.close
		set rs_unid = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (unidade) then%>
	<%'busca_unidade = replace(busca_unidade,".","_")
	'xsql = "up_unidade_busca @cd_unidade=' ', @nm_unidade="&busca_inteligente(busca_unidade)&", @cd_status=1"
	'set rs_unid = dbconn.execute(xsql)
	strsql = "Select * From View_unidades where nm_unidade like '%"&busca_inteligente_ajax(unidade)&"%' AND cd_status="&status&""
	set rs_unid = dbconn.execute(strsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%if not rs_unid.EOF Then%>
		<%while not rs_unid.EOF%>
		<%cd_unidade = rs_unid("cd_codigo")
		nm_unidade = rs_unid("nm_unidade")
		cd_status = rs_unid("cd_status")%>
			<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD.</b></td>
					<td>&nbsp;<b>UNIDADE</b></td>
					<td>&nbsp;<b>TIPO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
				</tr>
			<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td align="right"><%=zero(cd_unidade)%>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/unidade_cad.asp?cd_unidade=<%=cd_unidade%>&status=<%=cd_status%>&acao=editar&jan=1','Altera_nome','width=500,height=400')" return false;"><%=nm_unidade%></a></td>
		<td>&nbsp;<%=nm_equipamento_tipo%></td>
		<td id="no_print" valign="top" align="center">
			<%if status = "1" Then%>
				<img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_unidade%>')" type="button" value="A" title="Desativar">
			<%else%>
				<img src="imagens/ic_del.gif" border="0" onclick="javascript:JsDelete('<%=cd_unidade%>')" type="button" value="A" title="Excluir"> &nbsp; &nbsp;
				<img src="imagens/undo2.gif" border="0" onclick="javascript:JsReativar('<%=cd_unidade%>')" type="button" value="A" title="Reativar">
			<%end if%></td>
	</tr>
		<%num_lista = num_lista + 1
		rs_unid.MoveNext
		wend
	else%>
	<%rs_unid.close
	set rs_unid = nothing%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CÓD.</b></td>
			<td>&nbsp;<b>UNIDADE</b></td>
			<td>&nbsp;<b>TIPO</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>				
		</tr>
		<tr><td colspan="5">Não encontrado</td></tr>
	<%end if%>	   
</table>
<%Else%>
	<table>
		<tr>
			<td>Nenhum</td>
		</tr>
	</table>
<%End if%>