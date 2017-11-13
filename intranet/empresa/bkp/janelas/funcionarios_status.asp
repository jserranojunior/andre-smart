<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Atualiza situação</title>
</head>
<%cd_funcionario = request("cd_funcionario")
cd_dia = request("cd_dia")&"/"
cd_mes = request("cd_mes")&"/"
cd_ano = request("cd_ano")
	dt_desliga = cd_mes&cd_dia&cd_ano
	
dia = day(now)
mes = month(now)


	xsql = "Select * From View_funcionario Where cd_codigo = "&cd_funcionario&""
	Set rs = dbconn.execute(xsql)
	
	if not rs.EOF Then
		nm_nome = rs("nm_nome")
		cd_codigo = rs("cd_codigo")
		
		xsql = "Select * From View_funcionario_status Where cd_funcionario = "&cd_funcionario&" and cd_suspenso <> 1 order by dt_atualizacao DESC"
		Set rs_status = dbconn.execute(xsql)
			if not rs_status.EOF then
				nm_status = rs_status("nm_status")
			end if
	end  if
%>
<body onUnload="window.opener.location.reload()">
<table cellspacing="1" cellpadding="1" border="0" align="center">
<form action="../acoes/funcionarios_acao.asp" method="post" name="funcao" id="funcao">
<input type="hidden" name="acao" value="status">
<input type="hidden" name="cod" value="<%=cd_funcionario%>">
<tr><td bgcolor="#808080" style="color:white;" align="center" colspan="2"><b>SITUAÇÃO</b></td></tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Nome: &nbsp; </td>
	<td>&nbsp;<%=nm_nome%></td>
</tr>
<tr bgcolor="#e4e4e4">
    <td colspan="2">&nbsp;</td>
</tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Situação Atual: &nbsp; </td>
	<td>&nbsp;<%=nm_status%></td>
</tr>
<% strsql = "TBL_status"' order by nm_status"
  Set rs_status = dbconn.execute(strsql)
  'nm_funcao =  rs_func("nm_funcao")%>
  <%'if strcod = "" Then%>								
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Nova situação: &nbsp; </td>
	<td><select name="cd_status" class="inputs"> 
	<option value="">Selecione</option>
	<%Do While Not rs_status.EOF
	'rs_func("cd_codigo")%>
	<option value="<%=rs_status("cd_codigo")%>"<%if strcd_status = CInt(rs_status("cd_codigo")) then%>selected<%End if%>><%=rs_status("nm_status")%></option>
	  <%rs_status.movenext
		loop

		rs_status.close
		Set rs_status = nothing
	  %></td>
</tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Data: &nbsp; </td>		
    <td><select name="cd_dia" class="inputs">
							<%for i = 1 to 31%>
							<option value="<%=i%>" <%if i = dia then response.write("selected") End if%>> <%=i%></option>
							<%next%>
						</select>
						<select name="cd_mes" class="inputs">
							<%for i = 1 to 12%>
							<option value="<%=i%>" <%if i = mes then response.write("selected") End if%>><%=mes_selecionado(i)%></option>
							<%next%>
						</select>
						<input type="text" name="cd_ano" size="5" maxlength="4" class="inputs" value="<%=year(now)%>"></td>
</tr>
<tr bgcolor="#e4e4e4">
    <td align="center" colspan="2"><br><input type="submit" value="OK"></td>
</tr>
</table>
<br>
<table border="0" align="center">
	<tr>
		<td colspan="4" bgcolor="#808080" style="color:white;" align="center">Histórico (ordem decrescente)</td>
	</tr>
	<tr bgcolor="silver" style="color:white;">
		<td>&nbsp;</td>
		<td>&nbsp;Status</td>
		<td>&nbsp;Data</td>
		<td>&nbsp;Status</td>
	</tr>
	<% strsql = "SELECT TOP 5 * FROM View_funcionario_status WHERE cd_funcionario='"&cd_funcionario&"' ORDER BY dt_atualizacao DESC"
  	Set rs = dbconn.execute(strsql)
  	
	nr = 1
	While not rs.EOF 
  	nm_status = rs("nm_status")
	dt_atualizacao = rs("dt_atualizacao")
	cd_suspenso = rs("cd_suspenso")
	cd_status = rs("cd_codigo")
	
	
	if row_color = 1 then
		cor_linha = "#FFFFFF"
	else
		cor_linha = "#ccffff"
	end if
  %>
	<tr>
		<td>&nbsp;<%=zero(nr)%></td>
		<td>&nbsp;<%=nm_status%></td>
		<td>&nbsp;<%=day(dt_atualizacao)%>/<%=mesdoano(Month(dt_atualizacao))%>/<%=year(dt_atualizacao)%></td>
		<td align="center">
			<%if cd_suspenso = 1 then%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_status%>&cd_suspenso=0&tbl=status&acao=suspende&cd_funcionario=<%=cd_funcionario%>"><img src="../../imagens/ic_pendente.gif" alt="<%=cd_codigo%>" width="10" height="12" border="0"></a>	
			<%else%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_status%>&cd_suspenso=1&tbl=status&acao=suspende&cd_funcionario=<%=cd_funcionario%>"><img src="../../imagens/check_ok.gif" alt="<%=cd_codigo%>" width="25" height="12" border="0"></a>
			<%end if%></td>
	</tr>
	<%
	nr = nr + 1
	rs.movenext
	wend%>
</table>
</form>
</body>
</html>