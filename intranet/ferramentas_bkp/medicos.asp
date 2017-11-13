<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<html>
<head>
	<title>Untitled</title>
</head>

<body>
<%xsql ="SELECT COUNT(A055_crmmed) AS conta, A055_crmmed, A055_desmed FROM TBL_Medicos where a055_status = 1 GROUP BY A055_crmmed, A055_desmed order by A055_desmed"
Set rs = dbconn.execute(xsql)%>
	
	<table style="border:1px solid silver; border-collapse:collapse">
	<tr><td colspan="6" style="border:1px solid silver; background-color:silver; color:white;" align="center"><%=ucase("Relatório de médicos cadastrados no banco de dados")%></td></tr>
	<tr>
		<td style="border:1px solid silver;">&nbsp;</td>
		<td style="border:1px solid silver;"><%=ucase("CRM")%></td>
		<td style="border:1px solid silver;"><%=ucase("Cirurgião")%></td>		
	</tr>
	<tr>
	<%linha = 1
	num = 1
	
	while not rs.EOF
	conta = rs("conta")
	crm = rs("A055_crmmed")
	nome = rs("A055_desmed")
		if linha = 1 then
			cor_linha = "silver"
		else
			cor_linha = "white"
		end if%>
			
			
		<tr>
			<td style="background-color:<%=cor_linha%>;></td>;"><%=num%></td>
			<td style="background-color:<%=cor_linha%>;></td>;"><%=crm%></td>
			<td style="background-color:<%=cor_linha%>;"><%=nome%></td>			
		</tr>
			
			
	<%linha = linha + 1
	if linha > 2 then
		linha = 1
	end if
	
	num = num + 1
	registros_pag = registros_pag + 1
	rs.movenext
	wend%>
	</tr>
	<tr>
		<td><img src="../imagens/blackdot.gif" alt="" width="40" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="60" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td>	
	</tr>
	<tr><td colspan="3">TOTAL: <%=num-1%></td></tr>
</table>
</body>
</html>
