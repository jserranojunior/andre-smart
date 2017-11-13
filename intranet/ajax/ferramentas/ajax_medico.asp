<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

busca_medico = request("busca_medico")

num_lista = 1	%>

<%If  busca_medico = "" Then '*** Mostra campo normal ***%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CRM</b></td>
					<td>&nbsp;<b>MEDICO</b></td>
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
<%Elseif busca_medico = "todos" Then '*** Mostra campo normal ***%>
	<%strsql ="SELECT * FROM TBL_medicos where a055_status = 1 ORDER BY a055_desmed"
		Set rs_medico = dbconn.execute(strsql)
		num_lista = 1%>
	<table border="0" width="500">	
		<%Do while Not rs_medico.EOF
			
			cd_medico = rs_medico("a055_codmed")
			cd_crm = rs_medico("a055_crmmed")
			nm_medico = rs_medico("a055_desmed")
			cd_status = rs_medico("a055_status")%>
	
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CRM</b></td>
			<td>&nbsp;<b>MEDICO</b></td>
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
		<td valign="top"><a href="ferramentas.asp?tipo=medico&cd_medico=<%=cd_medico%>&cd_crm=<%=num_crm%>&nm_medico=<%=nm_medico%>&ordem=<%=ordem%>&status=1&acao=editar"><%=cd_crm%></a>&nbsp;&nbsp;</td>
		<td valign="top">&nbsp;<a href="ferramentas.asp?tipo=medico&cd_medico=<%=cd_medico%>&cd_crm=<%=num_crm%>&nm_medico=<%=nm_medico%>&ordem=<%=ordem%>&status=1&acao=editar">&nbsp;<%=nm_medico%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_medico%>')" type="button" value="A" title="Desativar"></td>		
			<%num_lista = num_lista + 1
			rs_medico.movenext
			loop%>
	</tr>	
</table>
<%Elseif IsNumeric (busca_medico) then ' *** Mostra Campo não numérico ***%>
	<%'xsql = "SELECT * FROM TBL_medicos Order by a055_desmed" '*** Mostra pelo nome
	xsql = "SELECT * FROM TBL_medicos  where a055_crmmed = '"&busca_medico&"' and a055_status = 1 ORDER BY a055_crmmed "
	Set rs_crm = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">
	<%while not rs_crm.EOF%>
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>CRM</b></td>
			<td>&nbsp;<b>MEDICO</b></td>
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
		<%cd_medico = rs_crm("a055_codmed")
		num_crm = rs_crm("a055_crmmed")
		nm_medico = rs_crm("a055_desmed")
		if int(cd_crm) = int(num_crm) then%><%med_sel = "selected"%><%end if%>
		<td><a href="ferramentas.asp?tipo=medico&cd_medico=<%=cd_medico%>&cd_crm=<%=num_crm%>&nm_medico=<%=nm_medico%>&ordem=<%=ordem%>&status=1&acao=editar"><%=num_crm%></a>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="ferramentas.asp?tipo=medico&cd_medico=<%=cd_medico%>&cd_crm=<%=num_crm%>&nm_medico=<%=nm_medico%>&ordem=<%=ordem%>&status=1&acao=editar"><%=nm_medico%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_medico%>')" type="button" value="A" title="Desativar"></td>
		<%num_lista = num_lista + 1
		rs_crm.MoveNext
		med_sel = ""
		  wend
		
		  rs_crm.close
		  set rs_crm = nothing%>
	</tr>
</table>
	
<%Elseif not IsNumeric (busca_medico) then%>
	<%'xsql = "SELECT * FROM TBL_medicos where a055_desmed like '"&cd_crm& "%' order by a055_desmed"
	xsql = "SELECT * FROM TBL_medicos where a055_desmed like '"&busca_medico& "%' and a055_status = 1 ORDER BY a055_desmed"
	Set rs_crm = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%if not rs_crm.EOF Then%>
		<%while not rs_crm.EOF%>
		<%cd_medico = rs_crm("a055_codmed")
		num_crm = rs_crm("a055_crmmed")
		nm_medico = rs_crm("a055_desmed")%>		
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>CRM</b></td>
			<td>&nbsp;<b>MEDICO</b></td>
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
		<td align="right"><a href="ferramentas.asp?tipo=medico&cd_medico=<%=cd_medico%>&cd_crm=<%=num_crm%>&nm_medico=<%=nm_medico%>&ordem=<%=ordem%>&status=1&acao=editar"><%=num_crm%>&nbsp;&nbsp;</a>
		<td>&nbsp;<a href="ferramentas.asp?tipo=medico&cd_medico=<%=cd_medico%>&cd_crm=<%=num_crm%>&nm_medico=<%=nm_medico%>&ordem=<%=ordem%>&status=1&acao=editar"><%=nm_medico%></a></td>
		<td id="no_print" valign="top"><img src="imagens/ic_reprovado.gif" onclick="javascript:JsDesativar('<%=cd_medico%>')" type="button" value="A" title="Desativar"></td>
	</tr>
		<%num_lista = num_lista + 1
		rs_crm.MoveNext
		wend
	else%>
		
	<%End if
	  rs_crm.close
	  set rs_crm = nothing%>	
</table>
<%end if%>
