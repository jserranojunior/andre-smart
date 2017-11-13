<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->

<html>
<%
tipo_manut = request("tipo_manut")
dt_obs = "01/01/"&request("ano_atribuido")
cd_unidade = request("cd_unidade")

cd_obs = request("cd_obs")
'acao = 1
	if cd_obs = "" then
		acao = "inserir"
	else
		acao = "editar"
	end if
%>

<head>
	<title>Manutenções - Patrimônio</title>
</head>

<body onUnload="window.opener.form.submit()">
<table border="0" class="txt" id="no_print" style="font-size:12px; font-family:arial;" align="center">
<%'sql = "SELECT * FROM "

			'if cd_plan_prev <> "" Then
			if acao = "editar" then
				strsql ="SELECT * FROM TBL_patrimonio_manut_obs WHERE cd_codigo = "&cd_obs&""'@ordem=10" '"&ordem&"', @condicao = "&condicao&""
			  	Set rs_obs = dbconn.execute(strsql)
				
				IF not rs_obs.EOF THEN
		 			nm_obs = rs_obs("nm_obs")
				End if
			elseif acao = "inserir_" then
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
<form action="acoes/patrimonio_acao.asp" method="post" name="patrimonio" onsubmit="return validar_patrimonio(document.form);" onSubmit="return checa(this);">
<input type="hidden" name="janela" value="1">

<input type="hidden" name="cd_tipo" value="manut_obs">
<input type="hidden" name="tipo_manut" value="<%=tipo_manut%>">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="dt_obs" value="<%=dt_obs%>">
<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
<input type="hidden" name="cd_codigo" value="<%=cd_obs%>">




<tr>
	<td colspan="3">&nbsp;<b>Inventário - <font color="red">Cadastro de observações</font></b><%'=cd_plan_prev%></td>
	<td align="right">&nbsp;</td>
</tr>
<tr bgcolor=#cococo>
	<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
</tr>
<tr>
	<td>&nbsp;<b>Observação:</b></td>
	<td>&nbsp;</td>
</tr>
<tr>
    <td><textarea cols="35" rows="5" name="nm_obs" onKeyDown="conta_caract"><%=nm_obs%></textarea></td>
</tr>
<tr>
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
			document.location.href('acoes/patrimonio_acao.asp?cd_codigo='+cod1+'&cd_tipo=manut_obs&acao=apagar');
		  }
	}
</SCRIPT>
</body>
</html>
