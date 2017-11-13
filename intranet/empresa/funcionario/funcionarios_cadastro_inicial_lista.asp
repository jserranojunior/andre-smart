<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Untitled</title>
</head>

<body>
<table border="1">
	<tr>
		<td colspan="4"> Listagem dos candidatos aprovados</td>
	</tr>
	<tr>
		<td>Nº</td>
		<td>Nome</td>
		<td>RG</td>
		<td>CPF</td>
	</tr>
	<%numero = 1
	xsql ="SELECT * FROM TBL_funcionario_inicial"
	Set rs = dbconn.execute(xsql)
	
		while not rs.EOF
			cd_funcionario = rs("cd_codigo")
			nm_nome = rs("nm_nome")
			nm_rg = rs("nm_rg")
			nm_cpf = rs("nm_cpf")%>
			<tr>
				<td><%=numero%></td>
				<td><%=nm_nome%></td>
				<td><%=nm_rg%></td>
				<td><%=nm_cpf%></td>
				<td><img src="../../imagens/redo.gif" alt="" width="25" height="24" border="0" onclick="window.open('empresa/funcionario/funcionarios_cadastro.asp?cod_ini=<%=cd_funcionario%>', 'janela_vdlap', 'width=800, height=600,  toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;">&nbsp;</td>
			</tr>
		<%numero = numero + 1
		rs.movenext
		wend%>
	
</table>


</body>
</html>
