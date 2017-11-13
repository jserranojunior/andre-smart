<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
<%
cd_valor = request("cd_valor")
xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_codigo='"&cd_valor&"'"
	Set rs_strvalor = dbconn.execute(xsql)
	while not rs_strvalor.EOF
		strcd_valor = rs_strvalor("cd_codigo")
		strnr_valor = rs_strvalor("nr_valor")
		strnm_valor = rs_strvalor("nm_valor")
		strcd_tipo_valor = rs_strvalor("cd_tipo")
		strdt_atualizacao = rs_strvalor("dt_atualizacao")
		strnm_obs = rs_strvalor("nm_obs")
	rs_strvalor.movenext
	wend

if strcd_valor = "" Then
	acao = "valor_padrao"
else
	acao = "valor_padrao_altera"
end if
%>
</head>

<body><br>
<table align="center">
	<form action="empresa/acoes/funcionarios_acao.asp" method="post">
	<input type="text" name="acao" value="<%=acao%>">
	<input type="text" name="cd_valor" value="<%=strcd_valor%>">
	<!--input type="hidden" name="cd_sessao" value="valor_padrao"-->
	<tr>
		<td colspan="2" align="center"><b>FUNCIONÁRIOS: VALORES PADRÕES</b></td>
	</tr>
	<tr><td colspan="2"><img src="../../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
	<tr>
		<td colspan="5"></td>
	</tr>
	<tr>
		<td><b>Tipo</b></td>
		<td><select name="cd_padrao_valor_tipo" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_funcionario_valores_tipos order by nm_valor"
										Set rs_valor = dbconn.execute(xsql)
										while not rs_valor.EOF
										cd_valor = rs_valor("cd_codigo")
										nm_valor = rs_valor("nm_valor")%>
											<option value="<%=cd_valor%>" <%if int(strcd_tipo_valor) = cd_valor then%>Selected<%end if%>><%=nm_valor%></option>
										<%rs_valor.movenext
										Wend%>
									</select> &nbsp; </td>
	</tr>
	<tr>
		<td><b>Valor R$</b></td>
		<td><input type="text" name="nr_padrao_valor" size="10" maxlength="20" class="inputs" value="<%=trim(strnr_valor)%>"></td>
	</tr>
	<tr>
		<td><b>Data</b></td>
		<td><input type="text" name="dt_padrao_atualizacao" size="10" maxlength="10" class="inputs" value="<%=strdt_atualizacao%>"></td>
	</tr>
	<tr>
		<td><b>Obs.</b></td>
		<td><textarea cols="40" rows="3" name="nm_padrao_obs" class="inputs"><%=strnm_obs%></textarea></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="submit" value="OK"> &nbsp; &nbsp; <input type="button" value="Novo" onclick="location.href='empresa.asp?tipo=valor_padrao'"></td>
	</tr>
	</form>
</table>
<br>
<br>
<table align="center" border="0" style="border-">
	<tr><td colspan="5"><img src="../../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
		<%xsql = "Select distinct cd_tipo From TBL_funcionario_valores_padroes"
		Set rs_tipo = dbconn.execute(xsql)
		while not rs_tipo.EOF
			cd_tipo = rs_tipo("cd_tipo")%>
			
			<%xsql = "Select top 3 * From VIEW_funcionario_valores_padroes Where cd_tipo='"&cd_tipo&"' order by dt_atualizacao desc"
				Set rs_valor = dbconn.execute(xsql)
				while not rs_valor.EOF
					cd_valor = rs_valor("cd_codigo")
					nr_valor = rs_valor("nr_valor")
					nm_valor = rs_valor("nm_valor")
					dt_atualizacao = rs_valor("dt_atualizacao")
					nm_obs = rs_valor("nm_obs")
					
					if cabeca = 0 Then
						if registros > 0 then%>
							<tr><td colspan="4">&nbsp;</td></tr>
						<%end if%>
					<tr>
						<td colspan="5" align="center" style="background:gray; color:white; font-weight:bold;">VALORES PADRÕES</td>
					</tr>
					<tr style="background:gray; color:white;">
						<td>Tipo</td>
						<td>Valor R$</td>
						<td>Data</td>
						<td>Obs.</td>
						<td>&nbsp;</td>
					</tr>
					<%end if%>
					<tr style="color:<%if registroatual > 0 Then%>silver;  font-weight:italic<%else%>black; font-weight:bold;<%end if%>">
						<td><%=nm_valor%>&nbsp;</td>
						<td align="right"><%=formatNumber(nr_valor)%>&nbsp;</td>
						<td><%=zero(day(dt_atualizacao))&"/"&zero(month(dt_atualizacao))&"/"&year(dt_atualizacao)%>&nbsp;</td>
						<td><%=nm_obs%>&nbsp;</td>
						<td><img src="../../imagens/ic_editar.gif" alt="Editar" width="13" height="14" border="0" onclick="location.href='empresa.asp?tipo=valor_padrao&cd_valor=<%=cd_valor%>'">&nbsp;
						<img src="../../imagens/ic_del.gif" alt="Apagar" width="13" height="14" border="0" onClick="javascript:JsValorpadraoDelete('<%=cd_valor%>')"></td>
					</tr>
				<%cabeca = cabeca + 1
				registroatual = registroatual + 1
				registros = registros + 1
				rs_valor.movenext
				Wend%>
			
	
		<%cabeca = 0
		registroatual = 0
		rs_tipo.movenext
		Wend%>
	
	
</table>

<SCRIPT LANGUAGE="javascript">
		function JsValorpadraoDelete(cod)
			   {
				if (confirm ("Tem certeza que deseja excluir este valor padrão?"))
			  {
			  document.location.href('../../empresa/acoes/funcionarios_acao.asp?cd_valor='+cod+'&acao=valor_padrao_especifico_delete');
				}
				  } 
	</SCRIPT>

</body>
</html>
