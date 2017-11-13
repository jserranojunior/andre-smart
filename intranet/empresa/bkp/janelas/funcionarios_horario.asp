<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Atualiza horários</title>
</head>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../../js/mascarainput.js"></SCRIPT>
<%cd_funcionario = request("cd_funcionario")
cd_dia = request("cd_dia")&"/"
cd_mes = request("cd_mes")&"/"
cd_ano = request("cd_ano")
	dt_desliga = cd_mes&cd_dia&cd_ano
	
dia = day(now)
mes = month(now)


	xsql = "Select * From View_funcionario Where cd_codigo = "&cd_funcionario&""
	'strsql = "TBL_funcionario_funcao where cd_funcionario='"&cd_funcaionario&"' order by cd_codigo Desc"
	Set rs = dbconn.execute(xsql)
	
	if not rs_EOF Then
		nm_nome = rs("nm_nome")
		cd_codigo = rs("cd_codigo")
		
		xsql = "Select * From View_funcionario_horario Where cd_funcionario = "&cd_funcionario&" and cd_suspenso <> 1 order by dt_atualizacao DESC"
		Set rs_hora = dbconn.execute(xsql)
			if not rs_hora.EOF Then
				hr_entrada = rs_hora("hr_entrada")
				hr_saida = rs_hora("hr_saida")
				nm_intervalo = rs_hora("nm_intervalo")
			end if
	end  if
%>
<body onUnload="window.opener.location.reload()">
<table cellspacing="1" cellpadding="1" border="0" align="center">
<form name="form" action="../acoes/funcionarios_acao.asp" method="post">
<input type="hidden" name="acao" value="horario">
<input type="hidden" name="cod" value="<%=cd_funcionario%>">
<tr bgcolor="#e4e4e4"">
    <td>Nome:</td>
	<td colspan="3">&nbsp; <%=nm_nome%></td>
</tr>
<tr bgcolor="#e4e4e4"">
    <td colspan="4">&nbsp;</td>
</tr>
<tr bgcolor="#e4e4e4"">
	<td>&nbsp;</td>
	<td align="center">entrada</td>
	<td align="center">intervalo</td>
	<td align="center">saída</td>
</tr>
<tr bgcolor="#e4e4e4"">
    <td>Horário Atual:</td>
	<td align="center">&nbsp;<%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%></td>
	<td align="center"><%=nm_intervalo%></td>
	<td align="center">&nbsp;<%=zero(hour(hr_saida))&":"&zero(minute(hr_saida))%></td>
</tr>
<% strsql = "TBL_status order by nm_status"
  Set rs_status = dbconn.execute(strsql)
  'nm_funcao =  rs_func("nm_funcao")%>
  <%'if strcod = "" Then%>								
<tr bgcolor="#e4e4e4"">
    <td>Novo horário</td>
	<td><input type="text" name="hr_entrada" size="6" maxlength="5"  class="inputs"></td>
	<td><!--input type="text" name="nm_intervalo" size="10" maxlength="10"  class="inputs"-->
		<select name="nm_intervalo" size="1">
			<option value=""></option>
			<option value="1 Hora">1 Hora</option>
			<option value="15 minutos">15 minutos</option>
		</select>
	</td>
	<td><input type="text" name="hr_saida" size="6" maxlength="5"  class="inputs"></td>
</tr>
<tr bgcolor="#e4e4e4"">
    <td>Data:</td>		
    <td colspan="3"><select name="cd_dia" class="inputs">
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
	<!--td colspan="3"><input type="text" name="dt_atualizacao" size="10" maxlength="10" class="inputs" value=""></td-->
</tr>
<tr bgcolor="#e4e4e4"">
    <td align="center" colspan="4"><br><input type="submit" value="OK"></td>
</tr>
</table>
<br>
<table border="0" align="center">
	<tr bgcolor="#808080" style="color:white;" align="center">
		<td colspan="6">Histórico (ordem decrescente)</td>
	</tr>
	<tr bgcolor="silver" style="color:white;">
		<td>&nbsp;</td>
		<td align="center">Entrada</td>
		<td align="center">Intervalo</td>
		<td align="center">Saída</td>
		<td align="center">Data</td>
		<td align="center">status</td>
	</tr>
	<% strsql = "SELECT TOP 5 * FROM View_funcionario_horario WHERE cd_funcionario='"&cd_funcionario&"' ORDER BY dt_atualizacao DESC"
  	Set rs = dbconn.execute(strsql)
  	
	nr = 1
	row_color = 1
	
	While not rs.EOF 
	cd_codigo = rs("cd_codigo")
	cd_funcionario = rs("cd_funcionario")
  	hr_entrada = rs("hr_entrada")
	hr_saida = rs("hr_saida")
	nm_intervalo = rs("nm_intervalo")
	dt_atualizacao = rs("dt_atualizacao")
	cd_suspenso = rs("cd_suspenso")
	
	if row_color = 1 then
		cor_linha = "#FFFFFF"
	else
		cor_linha = "#ccffff"
	end if%>
	<tr bgcolor="<%=cor_linha%>">
		<td>&nbsp;<%=zero(nr)%></td>
		<td align="center"><%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%></td>
		<td align="center"><%=nm_intervalo%></td>
		<td align="center"><%=zero(hour(hr_saida))&":"&zero(minute(hr_saida))%></td>
		<td>&nbsp;<%=day(dt_atualizacao)%>/<%=mesdoano(Month(dt_atualizacao))%>/<%=year(dt_atualizacao)%></td>
		<td align="center">
			<%if cd_suspenso = 1 then%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_codigo%>&cd_suspenso=0&tbl=horario&acao=suspende&cd_funcionario=<%=cd_funcionario%>"><img src="../../imagens/ic_pendente.gif" alt="" width="10" height="12" border="0"></a>	
			<%else%>
				<a href="../acoes/funcionarios_acao.asp?cod=<%=cd_codigo%>&cd_suspenso=1&tbl=horario&acao=suspende&cd_funcionario=<%=cd_funcionario%>"><img src="../../imagens/check_ok.gif" alt="<%=cd_codigo%>" width="25" height="12" border="0"></a>
			<%end if%></td>
		<td><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsHorarioDelete('<%=cd_codigo%>')"><%'=cd_codigo%></td>
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
	 {									   	
		MaskInput(document.form.hr_entrada, "99:99");
		MaskInput(document.form.hr_saida, "99:99");
		//MaskInput(document.form.dt_atualizacao, "99/99/9999");
	 }
	 
	 function Jsmuda(cod1, cod2, cod3, cod4)
		{
		  if (confirm ("Tem certeza que deseja excluir a folga?"))
			  {
				document.location.href('acoes/horas_trabalhadas_acao.asp?cd_codigo='+cod1+'&cd_cod_folga='+cod2+'&cd_folga='+cod3+'&acao=excluir&mesentrada=<%=mes_entrada%>&anoentrada=<%=ano_entrada%>');
			  }
		}
		
		
		function JsHorarioDelete(cod)
			   {
				if (confirm ("Tem certeza que deseja excluir esse horário?"))
			  {
			  document.location.href('../../empresa/acoes/funcionarios_acao.asp?cod_horario='+cod+'&acao=horario_delete');
				}
				  } 
</SCRIPT>
</SCRIPT>
</form>


</body>
</html>