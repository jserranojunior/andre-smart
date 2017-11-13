<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<br>
<table align="center">
	<tr><td></td></tr>



	<tr>
		<td colspan="100%" align="center">RELATÓRIO DO USO DE ÓTICAS</td>
	</tr>
	<tr style="color:#ffffff;" bgcolor="#808080">
		<td bgcolor="#808080">&nbsp;</td>
		<td>&nbsp;Ótica</td>
		<td>&nbsp;Data</td>
		<td>&nbsp;Unidade</td>
		<td>&nbsp;N/S</td>
		<td>&nbsp;Patrimonio</td>
		<td>&nbsp;Protocolo</td>
	</tr>
	<%xsql ="SELECT count(cd_patrimonio) as conta, cd_patrimonio FROM View_patrimonio_uso where nm_tipo = 'o' group by cd_patrimonio order by cd_patrimonio"
	Set rs = dbconn.execute(xsql)
	
	while not rs.EOF
	
	conta = rs("conta")
	cd_patrimonio = rs("cd_patrimonio")
	linha = 1
	
		xsql_registros ="SELECT * FROM View_patrimonio_uso where nm_tipo = 'o' and cd_patrimonio = '"&cd_patrimonio&"' order by cd_patrimonio"
		Set rs_registros = dbconn.execute(xsql_registros)
		
		while not rs_registros.EOF
		nm_equipamento = rs_registros("nm_equipamento")
		dt_data = rs_registros("a002_datpro")
		nm_unidade = rs_registros("nm_unidade")
		cd_ns = rs_registros("cd_ns")
		cd_patrimonio = rs_registros("cd_patrimonio")
		'protocolo = digitov(rs_registros("cd_unidade")&"."&proto(rs_registros("cd_protocolo")))
		protocolo = digitov(zero(rs_registros("cd_unidade"))&"."&proto(rs_registros("cd_protocolo")))%>
			<tr>
				<td style="color:#ffffff;" bgcolor="#808080" align="right"><%=linha%></td>
				<td>&nbsp;<%=nm_equipamento%></td>
				<td>&nbsp;<%=dt_data%></td>
				<td>&nbsp;<%=nm_unidade%></td>
				<td align="right">&nbsp;<%=cd_ns%></td>
				<td align="center">&nbsp;<%=cd_patrimonio%></td>
				<td align="right">&nbsp;<%=protocolo%></td>
			</tr>
		<%rs_registros.movenext
		wend%>
		<tr><td colspan="7" bgcolor="#c0c0c0">&nbsp;Subtotal: <%'=linha%><%=conta%></td></tr>
		<tr><td>&nbsp;</td></tr>
	<%%>
	<%linha = linha + 1
	rs.movenext
	wend%>
</table>


</body>
</html>
