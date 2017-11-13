<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Atualiza funções</title>
</head>
<%cd_funcionario = request("cd_funcionario")
	
dia = day(now)
mes = month(now)

	xsql = "Select * From View_funcionario Where cd_codigo = "&cd_funcionario&""	
	Set rs = dbconn.execute(xsql)
	
	if not rs.EOF Then
		nm_nome = rs("nm_nome")
				xsql_funcao = "Select * From View_funcionario_funcao Where cd_funcionario = "&cd_funcionario&" and cd_suspenso <> 1 ORDER BY dt_atualizacao desc"
				Set rs_funcao = dbconn.execute(xsql_funcao)
					if not rs_funcao.EOF then
						nm_funcao = rs_funcao("nm_funcao")
					end if
	end  if
%>

<body onLoad="document.form.cd_funcao.focus();" onUnload="window.opener.location.reload()">
<!--SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../../js/mascarainput.js"></SCRIPT-->

<table cellspacing="1" cellpadding="1" border="0" align="center">
<form  name="form" action="../acoes/funcionarios_acao.asp" method="post" id="funcao">
<input type="hidden" name="acao" value="funcoes">
<input type="hidden" name="cod" value="<%=cd_funcionario%>">
<tr><td bgcolor="#808080" style="color:white;" align="center" colspan="2"><b>Funções</b></td></tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Nome: &nbsp; </td>
	<td>&nbsp;<%=nm_nome%></td>
</tr>
<tr bgcolor="#e4e4e4">
    <td colspan="2">&nbsp;</td>
</tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Função Atual &nbsp; </td>
	<td>&nbsp;<%=nm_funcao%></td>
</tr>
<% strsql = "TBL_cargos where cd_especificidade = 4 order by nm_cargo"
  Set rs_funcao = dbconn.execute(strsql)
%>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Novo Cargo &nbsp; </td>
	<td><select name="cd_funcao" class="inputs"> 
	<option value="">Selecione</option>
	<%Do While Not rs_funcao.EOF%>
	<option value="<%=rs_funcao("cd_codigo")%>"<%if strcd_funcao = CInt(rs_funcao("cd_codigo")) then%>selected<%End if%>><%=rs_funcao("nm_cargo")%></option>
	  <%rs_funcao.movenext
		loop

		rs_funcao.close
		Set rs_funcao = nothing
	  %></td>	  
</tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Data: &nbsp; </td>		
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
</form>
<br>
<table border="0" align="center">
	<tr bgcolor="#808080" style="color:white;" align="center">
		<td colspan="4">Histórico (ordem decrescente)</td>
	</tr>
	<tr bgcolor="silver" style="color:white;">
		<td>&nbsp;</td>
		<td>&nbsp;Situação</td>
		<td>&nbsp;Data</td>
		<td>&nbsp;Status</td>
	</tr>
	<%strsql = "SELECT TOP 5 * FROM View_funcionario_funcao WHERE cd_funcionario='"&cd_funcionario&"' ORDER BY dt_atualizacao DESC"
  	Set rs = dbconn.execute(strsql)
  	
	nr = 1
	row_color = 1
	
	While not rs.EOF 
	cd_codigo = rs("cd_codigo")
	cd_funcao = rs("cd_funcao")
	nm_funcao = rs("nm_funcao")
	dt_atualizacao = rs("dt_atualizacao")
	cd_suspenso = rs("cd_suspenso")
	cd_funcionario = rs("cd_funcionario")
	
	if row_color = 1 then
		cor_linha = "#FFFFFF"
	else
		cor_linha = "#ccffff"
	end if
  %>
	<tr bgcolor="<%=cor_linha%>">
		<td>&nbsp;<%=zero(nr)%></td>
		<td>&nbsp;<!--a href=""--><%=nm_funcao%><!--/a--></td>
		<td>&nbsp;<%=day(dt_atualizacao)%>/<%=mesdoano(Month(dt_atualizacao))%>/<%=year(dt_atualizacao)%></td>	
		<td align="center">
			<%if cd_suspenso = 1 then%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_codigo%>&cd_suspenso=0&tbl=cargo&acao=suspende&cd_funcionario=<%=cd_funcionario%>"><img src="../../imagens/ic_pendente.gif" alt="<%=cd_codigo%>" width="10" height="12" border="0"></a>	
			<%else%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_codigo%>&cd_suspenso=1&tbl=cargo&acao=suspende&cd_funcionario=<%=cd_funcionario%>"><img src="../../imagens/check_ok.gif" alt="<%=cd_codigo%>" width="25" height="12" border="0"></a>
			<%end if%></td>
		<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsFuncaoDelete('<%=cd_funcao%>')"><%'=cd_funcao%></td>
	</tr>
	<%
	nr = nr + 1
	
	row_color = row_color + 1
	  	if row_color > 2 then
			row_color = 1
		end if
	rs.movenext
	wend%>
</table>

	<SCRIPT LANGUAGE="javascript">
		function JsFuncaoDelete(cod)
			   {
			    if (confirm ("Tem certeza que deseja excluir essa função?"))
			  {
			  location.href='../../empresa/acoes/funcionarios_acao.asp?cod_funcao='+cod+'&acao=funcao_delete';
			  //window.location('../../empresa/acoes/funcionarios_acao.asp?cod_funcao='+cod+'&acao=funcao_delete');
			 
				}
				  } 
	</SCRIPT>

	<!--SCRIPT LANGUAGE="javascript">
	{
	   	MaskInput(document.cargo.dt_atualizacao, "99/99/9999");	   
	 }
	</SCRIPT-->
							
</body>
</html>