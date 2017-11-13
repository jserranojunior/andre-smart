<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Atualiza cargos</title>
</head>
<%cd_funcionario = request("cd_funcionario")
cd_dia = request("cd_dia")&"/"
cd_mes = request("cd_mes")&"/"
cd_ano = request("cd_ano")
	dt_desliga = cd_mes&cd_dia&cd_ano
	
dia = day(now)
mes = month(now)

'	if dt_desliga <> "//" Then
'	response.write("É")
'	end if
	
	xsql = "Select * From View_funcionario Where cd_codigo = "&cd_funcionario&""
	Set rs = dbconn.execute(xsql)
	
	if not rs_EOF Then
		nm_nome = rs("nm_nome")&" "&rs("nm_sobrenome")
		
	end  if
%>
<body onUnload="window.opener.location.reload()">
<table cellspacing="1" cellpadding="1" border="1" align="center">
<form action="../acoes/funcionarios_acao.asp" method="post" name="funcao" id="funcao">
<input type="hidden" name="acao" value="altera">
<input type="hidden" name="cod" value="<%=cd_funcionario%>">
<input type="hidden" name="cd_status" value="4">
<tr>
    <td>Nome: <%=nm_nome%></td>
</tr>
<tr>
    <td>&nbsp;</td>
</tr>
<tr>
    <td>Selecione a data da dispensa:</td>
</tr>
<tr>			
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
<tr>
    <td align="center"><br><input type="submit" value="Dispensar"></td>
</tr>
</table>
<br>
<table border="1" align="center">
	<tr>
		<td colspan="3">Histórico (ordem decrescente)</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;Status</td>
		<td>&nbsp;Data</td>
	</tr>
	<% strsql = "SELECT TOP 6 * FROM View_funcionario_status WHERE cd_funcionario='"&cd_funcionario&"' ORDER BY cd_funcionario, dt_atualizacao DESC"
  	Set rs = dbconn.execute(strsql)
  	
	nr = 1
	While not rs.EOF 
  	nm_status = rs("nm_status")
	dt_atualizacao = rs("dt_atualizacao")
  %>
	<tr>
		<td>&nbsp;<%=zero(nr)%></td>
		<td>&nbsp;<%=nm_status%></td>
		<td>&nbsp;<%=day(dt_atualizacao)%>/<%=mesdoano(Month(dt_atualizacao))%>/<%=year(dt_atualizacao)%></td>
	</tr>
	<%
	nr = nr + 1
	rs.movenext
	wend%>
</table>
</form>
</body>
</html>