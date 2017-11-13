<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Atualiza Unidades</title>
</head>
<%cd_funcionario = request("cd_funcionario")

cd_dia = request("cd_dia")&"/"
cd_mes = request("cd_mes")&"/"
cd_ano = request("cd_ano")
	dt_desliga = cd_mes&cd_dia&cd_ano
	
dia = day(now)
mes = month(now)
ano = year(now)

cod_unidade = request("cod_unidade")
	if cod_unidade <> "" Then
		xsql = "Select * From TBL_funcionario_unidade Where cd_codigo = "&cod_unidade&""
		Set rs_unidade = dbconn.execute(xsql)
			if not rs_unidade.EOF Then
				dt_atualizacao = month(rs_unidade("dt_atualizacao"))&"/"&day(rs_unidade("dt_atualizacao"))&"/"&year(rs_unidade("dt_atualizacao"))
				dia = month(dt_atualizacao)
				mes = day(dt_atualizacao)
				ano = year(dt_atualizacao)
				strcd_unidade = rs_unidade("cd_unidade")
			end if
			
			modo = " - <b style='color:yellow;'>Alterar</b>"
			acao = "unidade_alterar"
	Else
		acao = "unidade"
	end if

	xsql = "Select * From View_funcionario Where cd_codigo = "&cd_funcionario&""
	'strsql = "TBL_funcionario_funcao where cd_funcionario='"&cd_funcaionario&"' order by cd_codigo Desc"
	Set rs = dbconn.execute(xsql)
	
	if not rs.EOF Then
		nm_nome = rs("nm_nome")
		cd_codigo = rs("cd_codigo")
		
		xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" And cd_suspenso <> 1 order by dt_atualizacao DESC"
		Set rs_unid = dbconn.execute(xsql)
			if not rs_unid.EOF Then
				nm_unidade_atual = rs_unid("nm_unidade")
			end if
	end  if
%>
<body onUnload="window.opener.location.reload()">
<table cellspacing="1" cellpadding="1" border="0" align="center">
<form action="../acoes/funcionarios_acao.asp" method="post" name="funcao" id="funcao">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="cod" value="<%=cd_funcionario%>">
<input type="hidden" name="cod_unidade" value="<%=cod_unidade%>">
<input type="hidden" name="fecha_janela" value="1">

<tr><td bgcolor="#808080" style="color:white;" align="center" colspan="2"><b>UNIDADE</b> <%=modo%></td></tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Nome: &nbsp;</td>
	<td>&nbsp; <%=nm_nome%></td>
</tr>
<tr bgcolor="#e4e4e4">
    <td colspan="2">&nbsp;</td>
</tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Unidade Atual: </td>
	<td>&nbsp;<%=nm_unidade_atual%></td>
</tr>
<%strsql = "SELECT * FROM TBL_unidades WHERE cd_status = 1 order by nm_unidade"
  Set rs_und = dbconn.execute(strsql)
  'nm_funcao =  rs_func("nm_funcao")%>
  <%'if strcod = "" Then%>								
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Nova Unidade &nbsp;</td>
	<td><select name="cd_unidade" class="inputs"> 
	<option value="">Selecione</option>
	<%Do While Not rs_und.EOF%>
	<option value="<%=rs_und("cd_codigo")%>"<%if strcd_unidade = rs_und("cd_codigo") then%>selected<%End if%>><%=rs_und("nm_unidade")%></option>
	  <%rs_und.movenext
		loop

		'rs_func.close
		'Set rs_func = nothing
	  %></td>
</tr>
<tr bgcolor="#e4e4e4">
    <td>&nbsp; Data: &nbsp;</td>		
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
						<input type="text" name="cd_ano" size="5" maxlength="4" class="inputs" value="<%=ano%>"></td>
</tr>
<tr bgcolor="#e4e4e4">
    <td align="center" colspan="2"><br><input type="submit" value="OK"></td>
</tr>
</form>
</table>
<br>
<table border="0" align="center">
	<tr>
		<td colspan="4" bgcolor="#808080" style="color:white;" align="center">Histórico (ordem decrescente)</td>
	</tr>
	<tr bgcolor="silver" style="color:white;">
		<td>&nbsp;</td>
		<td>&nbsp;Unidade</td>
		<td>&nbsp;Data</td>
		<td>&nbsp;Status</td>
	</tr>
	<% strsql = "SELECT TOP 5 * FROM View_funcionario_unidade WHERE cd_funcionario='"&cd_funcionario&"' ORDER BY dt_atualizacao DESC"
  	Set rs = dbconn.execute(strsql)
  	
	nr = 1
	While not rs.EOF 
	cod_unidade = rs("cd_codigo")
  	nm_unidade = rs("nm_unidade")
	nm_sigla = rs("nm_sigla")
	dt_atualizacao = rs("dt_atualizacao")
	cd_suspenso = rs("cd_suspenso")
  	%>
	<tr>
		<td valign="top">&nbsp;<a href="funcionarios_unidade.asp?cd_funcionario=<%=cd_funcionario%>&cod_unidade=<%=cod_unidade%>&acao=<%=acao%>"><%=zero(nr)%></a></td>
		<td valign="top">&nbsp;<%=nm_sigla%></td>
		<td valign="top">&nbsp;<%=day(dt_atualizacao)%>/<%=mesdoano(Month(dt_atualizacao))%>/<%=year(dt_atualizacao)%></td>
		<td valign="top">
			<%if cd_suspenso = 1 then%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_unidade%>&cd_suspenso=0&tbl=unidade&acao=suspende&cd_funcionario=<%=cd_funcionario%>&cd_unidade=<%=cd_unidade%>"><img src="../../imagens/ic_pendente.gif" alt="<%=cd_codigo%>" width="10" height="12" border="0"></a>	
			<%else%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_unidade%>&cd_suspenso=1&tbl=unidade&acao=suspende&cd_funcionario=<%=cd_funcionario%>"><img src="../../imagens/check_ok.gif" alt="<%=cd_codigo%>" width="25" height="12" border="0"></a>
			<%end if%></td>
		<td valign="top"><a href="funcionarios_unidade.asp?cd_funcionario=<%=cd_funcionario%>&cod_unidade=<%=cod_unidade%>&acao=unidade_apagar"></a><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsUnidadeDelete('<%=cod_unidade%>')"></td>
	</tr>
	<%
	nr = nr + 1
	rs.movenext
	wend%>
</table>
	<SCRIPT LANGUAGE="javascript">
		function JsUnidadeDelete(cod)
			   {
				if (confirm ("Tem certeza que deseja excluir essa Unidade?"))
			  {
			  document.location.href('../../empresa/acoes/funcionarios_acao.asp?cod_unidade='+cod+'&acao=unidade_delete');
				}
				  } 
	</SCRIPT>
</body>
</html>