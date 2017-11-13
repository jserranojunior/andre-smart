<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Altera Contrato</title>
</head>
<%cd_funcionario = request("cd_funcionario")
strcod_contrato = request("cod_contrato")
	
dia = day(now)
mes = month(now)

	xsql = "Select * From View_funcionario Where cd_codigo = "&cd_funcionario&""	
	Set rs = dbconn.execute(xsql)
	
	if not rs.EOF Then
	
		nm_nome = rs("nm_nome")
				'xsql_contrato = "Select * From View_funcionario_contrato_lista Where cd_funcionario = "&cd_funcionario&" ORDER BY dt_admissao desc"
				'xsql_contrato = "Select * From View_funcionario_contrato_lista Where cd_codigo = "&strcd_codigo&" ORDER BY dt_admissao desc"
				xsql_contrato = "Select * From View_funcionario_contrato_lista Where cod_contrato = '"&strcod_contrato&"' ORDER BY dt_admissao desc"
				Set rs_contrato = dbconn.execute(xsql_contrato)
					if not rs_contrato.EOF then
						cd_funcionario_contrato = rs_contrato("cd_codigo")
						cod_contrato = rs_contrato("cod_contrato")
						cd_contrato = rs_contrato("cd_contrato")
						cd_regime_trab = rs_contrato("cd_regime_trab")
						nm_regime_trab = rs_contrato("nm_regime_trab")
					end if
	end  if
%>

<body onload="document.form.cd_regime_trabalho.focus();" onUnload="window.opener.location.reload()">
<!--SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../../js/mascarainput.js"></SCRIPT-->

<table cellspacing="1" cellpadding="1" border="0" align="center">
<%if session("cd_codigo") = 46 then%>
	<tr><td colspan="2" align="center"><span style="font-size=8px;color:red;">empresa/janelas/funcionarios_contrato_altera.asp</span></td></tr>
<%end if%>
<form  name="form" action="../acoes/funcionarios_acao.asp" method="post" id="funcao">
<input type="hidden" name="acao" value="contrato_altera">
<input type="hidden" name="cd_funcionario_contrato" value="<%=cd_funcionario_contrato%>">
<input type="hidden" name="cod_contrato" value="<%=strcod_contrato%>">
<input type="hidden" name="fecha_janela" value="1">
<input type="hidden" name="cd_funcionario" value="<%=cd_funcionario%>">
<tr><td bgcolor="#808080" style="color:white; font-family:verdana; font-size:14px;" align="center" colspan="2"><b>Altera Contrato</b></td></tr>
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td>&nbsp; Nome: &nbsp; </td>
	<td>&nbsp;<%=nm_nome%></td>
</tr>
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td colspan="2">&nbsp;</td>
</tr>
<% strsql = "TBL_tipo_contrato order by nm_regime_trab"
  Set rs_contr = dbconn.execute(strsql)
%><tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td>&nbsp; Empresa &nbsp; </td>
	<td><select name="cd_regime_trabalho" class="inputs"> 
	<option value="">Selecione</option>
	<%Do While Not rs_contr.EOF%>
	<option value="<%=rs_contr("cd_codigo")%>"<%if int(cd_regime_trab) = CInt(rs_contr("cd_codigo")) then%>selected<%End if%>><%=rs_contr("nm_regime_trab")%></option>
	  <%rs_contr.movenext
		loop

		rs_contr.close
		Set rs_contr = nothing
	  %></td>	  
</tr>
<%' strsql = "TBL_tipo_contrato order by nm_regime_trab"
  'Set rs_contr = dbconn.execute(strsql)
%>
<%'strsql = "SELECT * FROM View_funcionario_contrato_lista WHERE cd_codigo='"&cd_funcionario_contrato&"'"
strsql = "SELECT * FROM View_funcionario_contrato_lista WHERE cod_contrato='"&strcod_contrato&"'"

  	Set rs_contr_lista = dbconn.execute(strsql)
  	
	nr = 1
	row_color = 1
	
	'While not rs_contr_lista.EOF 
	if not rs_contr_lista.EOF then
	cd_codigo = rs_contr_lista("cd_codigo")
	
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
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td>&nbsp; Admissão: &nbsp; </td>		
    <td><select name="cd_dia" class="inputs">
			<%for i = 1 to 31%>
			<option value="<%=i%>" <%if i = day(dt_admissao) then response.write("selected") End if%>> <%=i%></option>
			<%next%>
		</select>
		<select name="cd_mes" class="inputs">
			<%for i = 1 to 12%>
			<option value="<%=i%>" <%if i = month(dt_admissao) then response.write("selected") End if%>><%=mes_selecionado(i)%></option>
			<%next%>
		</select>
		<input type="text" name="cd_ano" size="5" maxlength="4" class="inputs" value="<%=year(dt_admissao)%>"></td>
	
</tr>
<%if dt_demissao <> "" then%>
<tr bgcolor="#e4e4e4" style="color:black; font-family:verdana; font-size:12px;">
    <td>&nbsp; Demissão: &nbsp; </td>		
	<td><select name="cd_dia_d" class="inputs">
			<option value=""></option>
			<%for i = 1 to 31%>
			<option value="<%=i%>" <%if i = day(dt_demissao) then response.write("selected") End if%>> <%=i%></option>
			<%next%>
		</select>
		<select name="cd_mes_d" class="inputs">
			<option value=""></option>
			<%for i = 1 to 12%>
			<option value="<%=i%>" <%if i = month(dt_demissao) then response.write("selected") End if%>><%=mes_selecionado(i)%></option>
			<%next%>
		</select>
		<input type="text" name="cd_ano_d" size="5" maxlength="4" class="inputs" value="<%=year(dt_demissao)%>"></td>
	
</tr>
<%end if%>
<%
	nr = nr + 1
	
	row_color = row_color + 1
	  	if row_color > 2 then
			row_color = 1
		end if
	rs_contr_lista.movenext
	end if
	'wend%>
<tr bgcolor="#e4e4e4" style="color:black; font-size:12px;">
    <td align="center" colspan="2"><br><input type="submit" value="Altera" style="color:black; font-size:12px;"><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0">
	<img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsContratoDelete('<%=strcod_contrato%>')"></td>
</tr>
</table>
<br>

</form>
		
		<SCRIPT LANGUAGE="javascript">
			function JsContratoDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse contrato?"))
				  {
				  document.location.href('../../empresa/acoes/funcionarios_acao.asp?cod_contrato='+cod+'&acao=contrato_delete');
					}
					  } 
		</SCRIPT>
		
	<!--SCRIPT LANGUAGE="javascript">
	{
	   	MaskInput(document.form.dt_atualizacao, "99/99/9999");	   
	 }
	</SCRIPT-->
							
</body>
</html>