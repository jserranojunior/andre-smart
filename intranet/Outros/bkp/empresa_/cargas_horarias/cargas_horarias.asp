<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->

<html>
<head>
	<title>Untitled</title>
</head>
<script language="Javascript">
<!-- 
  function mOvr(src,clrOver) {
    if (!src.contains(event.fromElement)) {
	  src.style.cursor = 'hand';
	  src.bgColor = clrOver;
	}
  }
  function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	  src.style.cursor = 'default';
	  src.bgColor = clrIn;
	}
  }
 // -->	
</script>
<body>
<%'xsql = "select * from view_cargas_todos"
xsql = "select * from TBL_funcionario where cd_regime_trabalho=1 order by nm_nome, nm_sobrenome"
set rs = dbconn.execute(xsql)%>
<table cellspacing="0" cellpadding="0" border="1">
<tr bgcolor="yellow">
    <td>&nbsp;Codigo&nbsp;</td>
    <td>&nbsp;Nome&nbsp;</td>
    <td>&nbsp;C&nbsp;</td>
	<td>&nbsp;D&nbsp;</td>
	<td colspan=15>&nbsp;Mes - Carga Horária - Carga Diaria&nbsp;</td>
	
</tr>
<%do while not rs.eof
	cd_codigo = rs("cd_codigo")
	nm_nome = rs("nm_nome")
	nm_sobrenome = rs("nm_sobrenome")
'	cd_mes = rs("cd_mes")
'	cd_carga_horaria = rs("cd_carga_horaria")
'	cd_carga_diaria = rs("cd_carga_diaria")
	dt_contratado = rs("dt_contratado")
	dt_demissao = rs("dt_demissao")
%>

<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=strcor%>');">
    <td>&nbsp;<%=cd_codigo%>&nbsp;</td>
    <td>&nbsp;<%=nm_nome%>&nbsp;<%=nm_sobrenome%>&nbsp;</td>
    <td>&nbsp;<%=month(dt_contratado)%>&nbsp;</td>
	<td>&nbsp;<%=right(dt_demissao,2)%>&nbsp;</td>
	<%xsql_2 = "select * from TBL_carga_horaria Where cd_funcionario = '"&cd_codigo&"' order by cd_mes"
	set rs_2 = dbconn.execute(xsql_2)
	do while not rs_2.EOF
	cd_mes = rs_2("cd_mes")
	cd_carga_horaria = rs_2("cd_carga_horaria")
	cd_carga_diaria = rs_2("cd_carga_diaria")
	%>
	<td>&nbsp;0<%=cd_mes%></td>
	<td>&nbsp;<%=total_mes(cd_carga_horaria)%>&nbsp;</td>
	<td align=right>&nbsp;<%=total_mes(cd_carga_diaria)%>&nbsp;</td>
	<%rs_2.movenext
	loop%>
	
	
</tr>
<%rs.movenext
loop%>
</table>




</body>
</html>
