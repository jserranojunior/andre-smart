<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<head>
	<title>Untitled</title>
</head>
<%str_titulo = request("titulo")
str_campo = request("campo")
str_variavel = request("variavel")
str_condicao = request("condicao")
	str_condicao = replace(str_condicao,"|","'")
	'*** tratamento dos elementos ***
	gestao_null = instr(1,str_condicao,"(cd_gestao = )",0)
	if gestao_null > 0 then
		str_condicao = replace(str_condicao,"(cd_gestao = )","(cd_gestao is null)")
	end if
%>
<body>
<table border="1" style="border:1px solid black; border-collapse:collapse;" align="center">
	<tr style="background-color:gray; color:white;">
		<td colspan="5" align="center"><%=str_titulo%></td>
	</tr>
	
<%if teste = "" Then

num = 0
	'strsql = "SELECT  * FROM  dbo.View_os_lista  where "&str_campo&"='"&str_variavel&"' ORDER BY dt_entrada"
	strsql = "SELECT  * FROM  View_os_lista2 "&str_condicao&" ORDER BY cd_unidade,dt_entrada,num_os"
	Set rs_manutencao = dbconn.execute(strsql)
			
		if rs_manutencao.EOF Then%>
			<tr style="background-color:white; color:red;">
				<td colspan="5" align="center"><b>Nenhuma ordem de serviço encontrada</b></td>
			</tr>
		<%else%>
			<tr>
				<td colspan="55" align="center">Lista O.S.</td>
			</tr>
			<tr style="background-color:silver; color:black;">
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td align="center">O.S.</td>
				<td align="center">Data</td>
				<td align="center">Item</td>
				<td ></td>
			</tr>
			<%While not rs_manutencao.EOF
				num = num + 1
				cd_codigo = rs_manutencao("cd_codigo")
				cd_unidade = rs_manutencao("cd_unidade")
				num_os = rs_manutencao("num_os")
				dt_entrada = rs_manutencao("dt_os")
				'nm_equipamento = rs_manutencao("nm_equipamento")
				nm_equipamento = rs_manutencao("nm_equipamento_novo")
				cd_patrimonio = rs_manutencao("cd_patrimonio")
				fecha = rs_manutencao("fecha")
					if fecha = 1 Then
						status = "F"
						cor_status = "red"
					else
						status = "A"
						cor_status = "green"
					end if%>
				<tr>
					<td align="center"><%=zero(num)%></td>
					<td  style="color:<%=cor_status%>;"><%=status%></td>
					<td align="right" style="cursor:pointer; cursor:hand;" onclick="window.open('manutencao_ver_janela2.asp?cd_codigo=<%=cd_codigo%>&campo=cd_codigo&visual=1&jan=1', 'principal', 'width=600, height=500,scrollbars=1'); return false;">
						<%=zero(cd_unidade)%>.<%=manutencao(num_os)%></td>
					<td align="center"><%=zero(day(dt_entrada))&"/"&zero(month(dt_entrada))&"/"&year(dt_entrada)%></td>
					<td><%=nm_equipamento%> <%'=cd_patrimonio%></td>
				</tr>
				
			<%rs_manutencao.movenext
			
			wend%>
			<tr><td colspan="5"><b>Total:</b> <%=zero(num)%> Orde<%if num <= 1 then%>m<%else%>ns<%end if%> de Serviço</td></tr>
<%end if%>
			<tr>
				<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="10" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="225" height="1" border="0"></td></tr>
		<%end if%>
		<!--tr><td colspan="5">SELECT  * <br>FROM  dbo.View_os_lista  <br><%'=str_condicao%> <br>ORDER BY dt_entrada,num_os</td></tr-->
</table>


</body>
</html>
