<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<br>
<table align="center" border="1" style="border-collapse:collapse;">

	<tr><td></td></tr>



	<tr>
		<td colspan="8" align="center">RELATÓRIO DO USO DE ÓTICAS e procedimentos cirurgicos</td>
	</tr>
	<%'xsql_registros = "SELECT * FROM VIEW_protocolo_patrimonio_uso where nm_tipo = 'o' order by cd_patrimonio"
	linha = 1
	xsql_registros = "SELECT COUNT(no_patrimonio) AS conta, cd_patrimonio, no_patrimonio, cd_equipamento, nm_equipamento_novo, cd_ns, cd_status FROM VIEW_protocolo_patrimonio_uso where no_patrimonio is not null GROUP BY cd_patrimonio, no_patrimonio, cd_equipamento, nm_equipamento_novo, cd_ns, cd_status ORDER BY cd_status, cd_equipamento, nm_equipamento_novo, conta DESC"
	
	set rs_registros = dbconn.execute(xsql_registros)
		
		while not rs_registros.EOF
		conta = rs_registros("conta")
		cd_equipamento = rs_registros("cd_equipamento")
		nm_equipamento = rs_registros("nm_equipamento_novo")
		cd_ns = rs_registros("cd_ns")
		no_patrimonio = rs_registros("no_patrimonio")
		cd_patrimonio = rs_registros("cd_patrimonio")
		cd_status = rs_registros("cd_status")
			if cd_status = 1 then
				nm_status = "Ativo"
			elseif cd_status = 2 then
				nm_status = "Descartado"
			else
				nm_status = " - - "
			end if
			
			
			if ultimo_equipamento <> cd_equipamento AND ultimo_equipamento <> "" Then%>
				<tr><td colspan="6">&nbsp;<%'=ultimo_equipamento&" - "&cd_equipamento%></td></tr>
			<%end if
		if ultimo_status <> cd_status then
			if ultimo_status <> "" Then%>
				<tr><td colspan="6" align="center">&nbsp;</td></tr>
			<%end if%>
		<tr><td colspan="6" align="center"><b><%=nm_status%></b></td></tr>
		<tr style="color:#ffffff;" bgcolor="#808080">
			<td bgcolor="#808080">&nbsp;</td>
			<td>&nbsp;Qtd</td>
			<td>&nbsp;Patrimonio</td>
			<td>&nbsp;Ótica</td>
			<td>&nbsp;N/S</td>
			<td>&nbsp;Status</td>
		</tr>
		<% ultimo_status = cd_status
		end if%>	
			<tr>
				<td style="color:#ffffff;" bgcolor="#808080" align="right"><%=linha%></td>
				<td>&nbsp;<%=conta%></td>
				<td align="center"><span onclick="window.open('patrimonio/patrimonio_consulta.asp?cd_pat_codigo=<%=cd_patrimonio%>', 'janela_vdlap', 'width=580, height=500, toolbar=0, location=0, status=0, scrollbars=0,resizable=1'); return false;">&nbsp;O<%=no_patrimonio%></span></td>
				<td>&nbsp;<%=nm_equipamento%></td>
				<td align="center">&nbsp;<%=cd_ns%></td>
				<td align="center">&nbsp;<%=nm_status%></td>
			</tr>			
		<%ultimo_equipamento = cd_equipamento
		linha = linha + 1
		rs_registros.movenext
		wend%>
		<tr><td colspan="6" align="center">&nbsp;</td></tr>
</table>


</body>
</html>
