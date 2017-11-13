<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Recisão de Contrato</title>
</head>
<%cd_funcionario = request("cd_funcionario")
func_ativos = request("func_ativos")
cod_contrato = request("cod_contrato")	
dia = day(now)
mes = month(now)
ano = year(now)
				
				'xsql_contrato = "sp_funcionarios_s02 @funcao='"&func_ativos&"', @dt_listagem='"&ano&"-"&mes&"-"&dia&"', @nm_funcionario='"&nome&"', @cd_funcionario='"&cd_funcionario&"'"
				xsql_contrato = "Select top 1 * From VIEW_funcionario_contrato_lista Where cd_codigo = "&cd_funcionario&" "
				Set rs_contrato = dbconn.execute(xsql_contrato)
					if not rs_contrato.EOF then
						'cd_funcionario = rs_contrato("cd_funcionario")
						cd_funcionario = rs_contrato("cd_codigo")
						nm_nome = rs_contrato("nm_nome")
						
						cd_contrato = rs_contrato("cd_contrato")
						nm_regime_trab = rs_contrato("nm_regime_trab")
						'nm_regime_trab = rs_contrato("nm_contrato_atual_abrv")
						'dt_admissao = rs_contrato("dt_admissao")
						dt_admissao = rs_contrato("dt_admissao_geral")
					end if
	'end  if
%>

<body onload="document.form.cd_dia.focus();" onUnload="window.opener.location.reload()">
<!--SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../../js/mascarainput.js"></SCRIPT-->

<table cellspacing="1" cellpadding="1" border="0" align="center">
<form  name="form" action="../acoes/funcionarios_acao.asp" method="post" id="funcao">
<input type="hiddens" name="cod" value="<%=cd_contrato%>">
<input type="hidden" name="acao" value="recisao">
<input type="hidden" name="func_ativos" value="<%=func_ativos%>">
<input type="hidden" name="cd_funcionario" value="<%=cd_funcionario%>">
<tr><td bgcolor="#ff6666" style="color:white; font-family:verdana; font-size:12px;" align="center" colspan="2"><b>Recisão de contrato</b></td></tr>
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td>&nbsp; Nome: &nbsp; </td>
	<td>&nbsp;<%=nm_nome%></td>
</tr>
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td colspan="2">&nbsp;</td>
</tr>
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td>&nbsp; Empresa &nbsp; </td>
	<td>&nbsp;<%=nm_regime_trab%></td>
</tr>
<% strsql = "TBL_tipo_contrato order by nm_regime_trab"
  Set rs_contr = dbconn.execute(strsql)
%>
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td>&nbsp; Data de admissão:&nbsp;</td>
	<td>&nbsp;<%=zero(day(dt_admissao))%>/<%=mesdoano(month(dt_admissao))%>/<%=year(dt_admissao)%></td>
</tr>
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td>&nbsp; Data de demissão: &nbsp; </td>		
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
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td align="center" colspan="2"><br><input type="submit" value="OK" style="color:black; font-family:verdana; font-size:12px;"></td>
</tr>
<tr>
	<td><img src="../../imagens/px.gif" alt="" width="135" height="1" border="0"></td>
	<td><img src="../../imagens/px.gif" alt="" width="205" height="1" border="0"></td>
</tr>
</table>
<br>
<table border="0" align="center">
	<tr bgcolor="#808080" style="color:white;" align="center" style="color:black; font-family:verdana; font-size:12px;">
		<td colspan="4">Histórico (ordem decrescente)</td>
	</tr>
	<tr bgcolor="silver" style="color:white;" style="color:black; font-family:verdana; font-size:12px;">
		<td>&nbsp;</td>
		<td>&nbsp;Empresa</td>
		<td>&nbsp;Data admissão</td>
		<td>&nbsp;Data demissão</td>
		<!--td>&nbsp;Status</td-->
	</tr>
	<%'strsql = "SELECT TOP 5 * FROM View_funcionario_contrato_lista WHERE cd_codigo='"&cd_contrato&"' ORDER BY dt_admissao DESC"
	strsql = "sp_funcionarios_s02 @funcao='SC', @dt_listagem='"&ano&"-"&mes&"-"&dia&"', @nm_funcionario='"&nome&"', @cd_funcionario='"&cd_funcionario&"'"
	
  	Set rs_contr_lista = dbconn.execute(strsql)
  	
	nr = 1
	row_color = 1
	
	While not rs_contr_lista.EOF
	cd_codigo = rs_contr_lista("cd_codigo")
	
	'cd_funcionario = rs_contr_lista("cd_funcionario")
	cd_funcionario = rs_contr_lista("cd_codigo")
	'nm_regime_trab = rs_contr_lista("nm_regime_trab")
	nm_empresa_abrv  = rs_contr_lista("nm_empresa_abrv")
	
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
	<tr bgcolor="<%=cor_linha%>" style="color:black; font-family:verdana; font-size:12px; background-color:#e9e9e9;">
		<td>&nbsp;<%=zero(nr)%></td>
		<td>&nbsp;<%=nm_empresa_abrv%><%'&nm_regime_trab%></td>
		<td>&nbsp;<%=day(dt_admissao)%>/<%=mesdoano(Month(dt_admissao))%>/<%=year(dt_admissao)%></td>
		<td align="center">
		<%if not isnull(dt_demissao) then%>
			<%=day(dt_demissao)%>/<%=mesdoano(Month(dt_demissao))%>/<%=year(dt_demissao)%>
		<%else%>
			---
		<%end if%></td>	
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


	<!--SCRIPT LANGUAGE="javascript">
	{
	   	MaskInput(document.cargo.dt_atualizacao, "99/99/9999");	   
	 }
	</SCRIPT-->
							
</body>
</html>