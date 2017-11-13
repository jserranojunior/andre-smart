<!--#include file="../../includes/inc_area_restrita.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%'cd_codigo = request("cd_codigo")
msg = request("msg")
%>
<html>
<table width="300" cellspacing="2" cellpadding="2" border="1" class="textopadrao">
<tr><td colspan="5" bgcolor="#808080" align="center"><font color="#ffffff">LISTAGEM DE USUÁRIOS</font></td></tr>
	<td bgcolor="#c5c5c5">Código</td>
	<td bgcolor="#c5c5c5">Nome</td>
	<td bgcolor="#c5c5c5">E-mail</td>
	<td bgcolor="#c5c5c5">&nbsp;</td>
	</tr>
	<%
	'xsql_1 ="up_adm_lista @cd_codigo=''"
	'Set rs1 = dbconn.execute(xsql_1)
	
	xsql1 = "SELECT * FROM TBL_ADM_usuario where cd_status > 0 order by nm_nome"
	Set rs1 = dbconn.execute(xsql1)
	
	
	Do while NOT rs1.EOF

	cd_cod = rs1("cd_codigo")
	nm_nome = rs1("nm_nome")
	nm_email = rs1("nm_email")
	nm_senha = rs1("nm_senha")
	%>
<tr bgcolor="#eeeeee">
	<td align="center"><%=zero(cd_cod)%></td>
	<td><a href="usuarios.asp?tipo=editar&cd_cod_usuario=<%=cd_cod%>"><%=nm_nome%></a></td>
	<td><%=nm_email%></td>
	<td>&nbsp;</td>
</tr>
	
	<%
	rs1.movenext
	loop
	
	rs1.close
	set rs1 = nothing
	
	cd_codigo = ""
	%>
<tr><td colspan="100%">&nbsp;</td></tr>
<tr>
	<td align="center"><a href="usuarios.asp?tipo=novo">Novo</a></td>
</tr>

</table>

</body>
</html>




