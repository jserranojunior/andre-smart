<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_plan_prev = request("cd_plan_prev")
mes_atribuido = request("mes_atribuido")
ano_atribuido = request("ano_atribuido")
tipo_manut = request("tipo_manut")
jan = request("jan")

'str_patrimonio = request("cd_patrimonio")
'nm_tipo = request("nm_tipo")
''cd_equipamento =  request("cd_equipamento")
'n'm_equipamento = request("nm_equipamento")
'cd_especialidade = request("cd_especialidade")
'	if cd_especialidade = "" Then 
'		cd_especialidade = "0" 
'	end if
'cd_marca =  request("cd_marca")
'nm_marca =  request("nm_marca")
'cd_ns =  request("cd_ns")
'cd_pn =  request("cd_pn")
'cd_unidade =  request("cd_unidade")
'nm_sigla =  request("nm_sigla")

'dt_data =  request("dt_data")
'ordem = request("ordem")
'	if ordem = "" Then
'	ordem = "nm_tipo, cd_patrimonio"
'	end if
acao = request("acao")
	if acao = "" or acao = "inserir" Then
	acao = "inserir"
	Else
	acao = "editar"
	end if

mensagem = request("mensagem")
%>


<head>
	<title>Inventário</title>
</head>
<body onUnload="window.opener.form.submit()">

<table border="0" class="txt" id="no_print" style="font-size:12px; font-family:arial;" align="center">
<%'sql = "SELECT * FROM "

			'if cd_plan_prev <> "" Then
			if acao = "editar" then
				strsql ="SELECT * FROM View_patrimonio_manutencoes WHERE cd_codigo = "&cd_plan_prev&""'@ordem=10" '"&ordem&"', @condicao = "&condicao&""
			  	Set rs_patrimonio = dbconn.execute(strsql)
				
				IF not rs_patrimonio.EOF THEN
		 			cd_pat_codigo = rs_patrimonio("cd_codigo")
					no_patrimonio = rs_patrimonio("no_patrimonio")
					cd_equipamento = rs_patrimonio("cd_equipamento")
					nm_equipamento_novo = rs_patrimonio("nm_equipamento_novo")
					'tipo_manut = rs_patrimonio("tipo_manut")
					
						dt_plan_prev = rs_patrimonio("dt_plan_prev")
						dt_patrimonio_garantia = rs_patrimonio("dt_garantia")
					
					condicao = " Where dt_descarte is null"
					ordem = " "&ordem&" "
						'strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"'"
					  	'Set rs_espec = dbconn.execute(strsql_espec)
						'	if not rs_espec.EOF then
						'		nm_especialidade = rs_espec("nm_especialidade")
						'	end if
				End if
			elseif acao = "inserir" then
				'strsql ="SELECT * FROM View_patrimonio_manutencoes WHERE cd_codigo = "&cd_plan_prev&""'@ordem=10" '"&ordem&"', @condicao = "&condicao&""
			  	strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_codigo = "&cd_plan_prev&""
				Set rs_patrimonio = dbconn.execute(strsql)
			'	
					IF not rs_patrimonio.EOF THEN
						'cd_patrimonio = rs_patrimonio("cd_patrimonio")
						no_patrimonio = rs_patrimonio("no_patrimonio")
						cd_pat_codigo = rs_patrimonio("cd_codigo")
						nm_equipamento_novo = rs_patrimonio("nm_equipamento_novo")
						
					End if
				
			End if

%>
<form action="acoes/patrimonio_acao.asp" method="post" name="form" onsubmit="return validar_patrimonio(document.form);" onSubmit="return checa(this);">
<input type="hidden" name="jan" value="<%=jan%>">
<input type="hidden" name="cd_tipo" value="manutencao">
<input type="hidden" name="cd_plan_prev" value="<%=cd_plan_prev%>">
<input type="hidden" name="tipo_manut" value="<%=tipo_manut%>">
<input type="hidden" name="acao" value="<%=acao%>">



<tr>
	<td colspan="3">&nbsp;<b>Inventário - <font color="red">Cadastro de equipamentos/instrumentos</font></b><%'=cd_plan_prev%></td>
	<td align="right">&nbsp;</td>
</tr>
<tr bgcolor=#cococo>
	<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
</tr>
<tr>
	<td>&nbsp;<b> Manutenção:</b></td>
	<td><b>
		<%if tipo_manut = 1 Then
			response.write("PREVENTIVA")
		elseif tipo_manut = 2 then
			response.write("CALIBRAÇAO")
		elseif tipo_manut = 3 then
			response.write("SEGURANÇA ELÉTRICA")
		end if
	%></b></td>
</tr>
<tr>
    <td>&nbsp;<b> Item:</b></td>
	<td><%=nm_equipamento_novo%></td>
</tr>
<tr>
    <td>&nbsp;<b> Nº Patrimônio:</b></td>
	<td><%=no_patrimonio%></td>
</tr>
<tr><td align=center colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<%if acao <> "inserir" then%>
<tr>
	<td>&nbsp;<b>Agendada para:</b></td>
	<td><b><%=ucase(left(mes_selecionado(month(dt_plan_prev)),3))%> / <%=year(dt_plan_prev)%></b></td>
	<td colspan="2" align="right"> &nbsp; <img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="javascript:JsDelete('<%=cd_pat_codigo%>')"><%'=cd_pat_codigo%></td>
</tr>
<%end if%>

<tr>
	<td>&nbsp;<b><%if acao <> "inserir" then%>Alterar para:<%else%>Agendar para<%'=mes_atribuido&"/"&ano_atribuido%><%end if%></b>
	</td>
	<td colspan="3">
	<select name="dt_plan_mes_manut" onFocus="nextfield ='dt_plan_ano_manut';">
		<option value=""></option>
	<%for i = 1 to 12%>	
		<option value="<%=i%>" <%if int(mes_atribuido) = i then%>Selected<%end if%>><%=ucase(left(mes_selecionado(i),3))%></option>
	<%next%>
	</select>/<input type="text" name="dt_plan_ano_manut" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='cd_calibracao';" value="<%=ano_atribuido%>">
	<%if cd_patrimonio <> "" Then
		strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = 1 order by dt_plan_prev desc"
			Set rs_prev = dbconn.execute(strsql)
		while not rs_prev.EOF
			dt_plan_prev = rs_prev("dt_plan_prev")%>
		<b style="color:#006600;"><%=Ucase(left(mes_selecionado(month(dt_plan_prev)),3))%>/<%=year(dt_plan_prev)%></b>
		<%rs_prev.movenext
		wend%>
	<%end if%></td>
</tr>
<tr>
	<td>&nbsp;<b>Garantia</b></td>
	<td><input type="checkbox" name="cd_patrimonio_garantia" value="1" <%if dt_patrimonio_garantia <> "" then%>checked<%end if%>> Sim</td>
</tr>
<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<tr>
    <td>&nbsp;</td>
    <td><input type="submit" name="envia" value=" OK "> <br><font color="red"><%=mensagem%></font></td>
	
</tr>
<%if cd_pat_codigo <> "" Then%>
<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<%end if%>
</form>
<tr>
    <td colspan="4">&nbsp;</td>
</tr>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
</tr>
</table>
<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1)
	{
	  if (confirm ("Tem certeza que deseja excluir?"))
		  {
			document.location.href('acoes/patrimonio_acao.asp?cd_plan_prev='+cod1+'&cd_tipo=manutencao&acao=apagar');
		  }
	}
</SCRIPT>
</body>
</html>
