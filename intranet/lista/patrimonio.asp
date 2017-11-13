<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<style type="text/css" media="print">
<!--
#no_print{ visibility:hidden; display: none;}
#ok_print{ visibility:visible; display:block;}
-->
</style>
<style type="text/css" media="screen">
<!--
#ok_print{ visibility:hidden; display: none;}
#no_print{ visibility:visible; display:block;}

-->
</style>
<html>
<head>
	<title>Untitled</title>
</head>

<body>
<table border="1">
	<tr style="background-color:gray; color:white;"><td colspan="3"> &nbsp; Lista dos nomes da tabela de equipamentos/instrumentos/óticas/materiais &nbsp; </td></tr>
	<tr style="background-color:silver;"><td>Nº</td><td>Status</td><td>Item</td><!--<td align="center">M</td>--><td align="center">P</td><!--<td>Cod</td>--></tr>
	<%num_linha = 0
	num_linha_pag = 0
	xsql_eqp = "Select * From TBL_equipamento where cd_status=1 and nm_equipamento_novo is not Null order by nm_equipamento_novo"
		Set rs_eqp = dbconn.execute(xsql_eqp)
		Do while not rs_eqp.EOF
			cd_equip = rs_eqp("cd_codigo")
			nm_equip = rs_eqp("nm_equipamento_novo")
			
			cd_status = rs_eqp("cd_status")
				if int(cd_equipamento) = cd_equip then
					selected = "selected" 
				End if
				num_linha = num_linha + 1
				num_linha_pag = num_linha_pag + 1%>
			<tr>
				<td><%=num_linha%></td>
				<td align="center" valign="middle"><%if cd_status = 1 then%><img src="../imagens/blackdot.gif" alt="" width="10" height="10" border="0"><%Else%>&nbsp;<%end if%></td>
				<td><%'=cd_equip%><%=nm_equip%></td>
				<!--<td align="center">&nbsp;<%'xsql = "Select distinct cd_equipamento From VIEW_os_lista where cd_equipamento = '"&cd_equip&"'"
					'Set rs = dbconn.execute(xsql)
					'if not rs.EOF then
					'	cd_os = rs("cd_equipamento")
					'	
					'		if int(cd_os) <> "" then%>X<%'end if
					'end if
					'cd_os = ""%>&nbsp;</td>-->
					<td align="center">&nbsp;<%xsql = "Select distinct cd_equipamento From VIEW_patrimonio_lista where cd_equipamento = '"&cd_equip&"'"
					Set rs = dbconn.execute(xsql)
					if not rs.EOF then
						cd_os = rs("cd_equipamento")
						
							if int(cd_os) <> "" then%>X<%end if
					end if
					cd_os=""%>&nbsp;</td>
					<!--<td><%'=cd_equip%></td>-->
			</tr>
			</option>
		<%if num_linha_pag = 40 then
			num_linha_pag = 0
			cabeca = 0%>
				<tr><td style="page-break-after:always;" colspan="5" id="ok_print" colspan="5">&nbsp;</td></tr>
				<tr style="background-color:gray; color:white;" id="ok_print"><td colspan="5"> &nbsp; Lista dos nomes da tabela de equipamentos/instrumentos/óticas/materiais &nbsp; </td></tr>
				<tr style="background-color:silver;" id="ok_print"><td>Nº</td><td>Status</td><td>Item</td><!--<td align="center">M</td>--><td align="center">P</td><!--<td>Cod</td>--></tr>
		<%end if
		rs_eqp.movenext
		selected = ""
		loop%>
</table>
</body>
</html>
