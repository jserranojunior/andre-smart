<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Atualiza Contrato</title>
</head>
<%cd_funcionario = request("cd_funcionario")

	
dia = day(now)
mes = month(now)

	xsql = "Select * From View_funcionario Where cd_codigo = "&cd_funcionario&""	
	Set rs = dbconn.execute(xsql)
	
	if not rs.EOF Then
	
		nm_nome = rs("nm_nome")
				xsql_contrato = "Select * From View_funcionario_contrato_lista Where cd_funcionario = "&cd_funcionario&" ORDER BY dt_admissao desc"
				Set rs_contrato = dbconn.execute(xsql_contrato)
					if not rs_contrato.EOF then
						'nm_nome = rs_contrato("nm_nome")
						nm_regime_trab = rs_contrato("nm_regime_trab")
					end if
	end  if
%>

<body onUnload="window.opener.location.reload()">
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../../js/mascarainput.js"></SCRIPT>

<table cellspacing="1" cellpadding="1" border="0" align="center">
<form  name="cargo" action="../acoes/funcionarios_acao.asp" method="post" id="funcao">
<input type="hidden" name="acao" value="contrato">
<input type="hidden" name="cd_funcionario" value="<%=cd_funcionario%>">
<tr><td bgcolor="#808080" style="color:white;" align="center" colspan="2"><b>Contrato</b></td></tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Nome: &nbsp; </td>
	<td>&nbsp;<%=nm_nome%></td>
</tr>
<tr bgcolor="#e4e4e4">
    <td colspan="2">&nbsp;</td>
</tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Contrato Atual &nbsp; </td>
	<td>&nbsp;<%=nm_regime_trab%></td>
</tr>
<% strsql = "TBL_tipo_contrato order by nm_regime_trab"
  Set rs_contr = dbconn.execute(strsql)
%>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Novo Contrato &nbsp; </td>
	<td><select name="cd_regime_trabalho" class="inputs"> 
	<option value="">Selecione</option>
	<%Do While Not rs_contr.EOF%>
	<option value="<%=rs_contr("cd_codigo")%>"<%if strcd_funcao = CInt(rs_contr("cd_codigo")) then%>selected<%End if%>><%=rs_contr("nm_regime_trab")%></option>
	  <%rs_contr.movenext
		loop

		rs_contr.close
		Set rs_contr = nothing
	  %></td>	  
</tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Data de admissão: &nbsp; </td>		
    <!--td><select name="cd_dia" class="inputs">
							<%for i = 1 to 31%>
							<option value="<%=i%>" <%if i = dia then response.write("selected") End if%>> <%=i%></option>
							<%next%>
						</select>
						<select name="cd_mes" class="inputs">
							<%for i = 1 to 12%>
							<option value="<%=i%>" <%if i = mes then response.write("selected") End if%>><%=mes_selecionado(i)%></option>
							<%next%>
						</select>
						<input type="text" name="cd_ano" size="5" maxlength="4" class="inputs" value="<%=year(now)%>"></td-->
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
	<tr bgcolor="#808080" style="color:white;" align="center">
		<td colspan="4">Histórico (ordem decrescente)</td>
	</tr>
	<tr bgcolor="silver" style="color:white;">
		<td>&nbsp;</td>
		<td>&nbsp;Empresa</td>
		<td>&nbsp;Data admissão</td>
		<td>&nbsp;Data demissão</td>
		<!--td>&nbsp;Status</td-->
	</tr>
	<%strsql = "SELECT TOP 5 * FROM View_funcionario_contrato_lista WHERE cd_funcionario='"&cd_funcionario&"' ORDER BY dt_admissao DESC"
  	Set rs_contr_lista = dbconn.execute(strsql)
  	
	nr = 1
	row_color = 1
	
	While not rs_contr_lista.EOF 
	'cd_codigo = rs_contr_lista("cd_codigo")
	
	cd_funcionario = rs_contr_lista("cd_funcionario")
	nm_regime_trab = rs_contr_lista("nm_regime_trab")
	dt_admissao = rs_contr_lista("dt_admissao")
	dt_demissao = rs_contr_lista("dt_demissao")
	
	'dt_atualizacao = rs("dt_atualizacao")
	'cd_suspenso = rs("cd_suspenso")
	'cd_funcionario = rs("cd_funcionario")
	
	if row_color = 1 then
		cor_linha = "#FFFFFF"
	else
		cor_linha = "#ccffff"
	end if
  %>
	<tr bgcolor="<%=cor_linha%>">
		<td>&nbsp;<%=zero(nr)%></td>
		<td>&nbsp;<%=nm_regime_trab%></td>
		<td>&nbsp;<%=zero(day(dt_admissao))%>/<%=mesdoano(Month(dt_admissao))%>/<%=year(dt_admissao)%></td>
		<td>&nbsp;<%=zero(day(dt_demissao))%>/<%=mesdoano(Month(dt_demissao))%>/<%=year(dt_demissao)%></td>	
		<!--td align="center">
			<%if cd_suspenso = 1 then%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_codigo%>&cd_suspenso=0&tbl=cargo&acao=suspende&cd_funcionario=<%=cd_funcionario%>"><img src="../../imagens/ic_pendente.gif" alt="<%=cd_codigo%>" width="10" height="12" border="0"></a>	
			<%else%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_codigo%>&cd_suspenso=1&tbl=cargo&acao=suspende&cd_funcionario=<%=cd_funcionario%>"><img src="../../imagens/check_ok.gif" alt="<%=cd_codigo%>" width="25" height="12" border="0"></a>
			<%end if%></td-->
	</tr>
	<%
	nr = nr + 1
	
	row_color = row_color + 1
	  	if row_color > 2 then
			row_color = 1
		end if
	rs_contr_lista.movenext
	wend%>
</table>
</form>


	<SCRIPT LANGUAGE="javascript">
	{
	   	MaskInput(document.cargo.dt_atualizacao, "99/99/9999");	   
	 }
	</SCRIPT>
							
</body>
</html>