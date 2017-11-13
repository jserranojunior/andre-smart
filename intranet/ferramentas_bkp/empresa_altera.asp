<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_pn = request("cd_pn")
cd_equipamento = request("cd_equipamento")
nm_equipamento = request("nm_equipamento")
status = request("status")
	if status = "" or status = 1 then
		str_status = " where cd_status = 1 "
	elseif status = 0 then
		str_status = " where cd_status = 0 "
	end if

acao = request("acao")
	if acao = "" Then
	acao = "inserir"
	Else
	acao = "editar"
	end if

mensagem = request("mensagem")
%>

<head>
	<title>Cadastro de material</title>
</head>
<script language="JavaScript">
	function validar_equipamento(shipinfo){
			if (shipinfo.nm_equipamento.value.length==""){window.alert ("O equipamento deve ser digitado.");shipinfo.nm_equipamento.focus();return (false);}	
		return (true);}	
</script>

<script language="javascript">
<!--
  function mOvr(src,clrOver) {
    if (!src.contains(event.fromElement)) {
	  src.style.cursor = 'hand';
	  src.bgColor = clrOver;
	}
  }
  function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	  src.style.cursor = 'default';
	  src.bgColor = clrIn;
	}
  }


// -->	
</script>

<!--#include file="../js/ajax.js"-->
<!--#include file="../js/ferramentas.js"-->

<body onload="foco()"><br>
<table class="txt" align="center" id="no_print" width="500">
<form action="../acoes/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_equipamento(document.Form);">
<!--input type="hidden" name="janela" value="1"-->
<input type="hidden" name="modo" value="equipamento">
<input type="hidden" name="cd_equipamento" value="<%=cd_equipamento%>">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="status" value="<%=status%>">
<%if status = 0 then
	atividade = "#ccffff"
Elseif status = 1 then
	atividade = "#ffcc99"
end if%>

<tr bgcolor="<%=atividade%>">
	<td colspan="100%" style="font-size=15px;">&nbsp;<b>Equipamento - <font color="red">Manutenção do cadastro de equipamento</font></b></td>
</tr>
<tr bgcolor=#cococo>
	<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
</tr>
<tr bgcolor="<%=atividade%>"><td align=center colspan="100%">&nbsp;</td></tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp; PART NUMBER:</td>
    <td colspan="3"><input type="text" name="cd_pn" size="35" value="<%=cd_pn%>" class="inputs"></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp; EQUIPAMENTO:</td>
    <td colspan="3"><input type="text" name="nm_equipamento" size="60" value="<%=nm_equipamento%>" class="inputs"></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td colspan="4">&nbsp;</td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;</td>
    <td colspan="3"><input type="submit" name="envia" value="OK">&nbsp; 
	<a href="ferramentas.asp?tipo=equipamento&status=1">Novo</a> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	(Ir para: <%if status = 0then%>
	<a href="ferramentas.asp?tipo=equipamento&status=1">Ativos</a>
	<%elseif status = 1  or status = "" then%>
	<a href="ferramentas.asp?tipo=equipamento&status=0">Inativos</a>
	<%end if%>)
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(<a href="ferramentas.asp">Ferramentas</a>)
	<br>
	<font color="red"><%=mensagem%></font></td>
</tr>
<tr>
    <td colspan="4"><img src="../imagens/blackdot.gif" alt="" width="500" height="1" border="0"></td>
</tr>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="330" height="1" border="0"></td>
</tr>
</table>
<br><br><!-- LISTAGEM DE equipamentos -->
<table class="txt" border="0" width="500" align="center">
<%if status <> 0 or status = "" then%>
<tr>
	<td colspan="100%">&nbsp;<b>Busca: <input type="text" name="busca_equipamento" class="inputs" size="10" maxlength="30" value="<%=cd_equipamento%>" onKeyup="mostra_equipamento();"></td>
</tr>
<%end if%>
<!--tr bgcolor="#C0C0C0">
	<td align="center">Nº</td>
	<td>&nbsp;<b>CRM</b></td>
	<td>&nbsp;<b>MÉDICO</b></td>	
	<!--td colspan="2">&nbsp;<b><a href="ferramentas.asp?tipo=medico&ordem=a055_datcad">CADASTRADO EM</b></a></td-->	
</tr-->
<tr>
	<td colspan="100%"><img src="../imagens/px.gif" alt="" width="100%" height="1" border="0"></td>	
</tr>

<tr>
	<td colspan="100%">
		<span id='divEquip'>
			<%if status = 1 then%>			
			<table border="0" width="500" align="center">
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD</b></td>
					<td>&nbsp;<b>EQUIPAMENTO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>						
				</tr>
			</table>
			<%elseif status = 0 then
			strsql ="SELECT * FROM TBL_equipamento where cd_status = 0 ORDER BY nm_equipamento"
			Set rs_equipamento = dbconn.execute(strsql)
			num_lista = 1%>
			<table border="0" width="500">
			<%Do while Not rs_equipamento.EOF
				
				cd_pn = rs_equipamento("cd_pn")
				cd_equipamento = rs_equipamento("cd_codigo")
				nm_equipamento = rs_equipamento("nm_equipamento")
				cd_status = rs_equipamento("cd_status")
				nm_equipamento_novo = rs_equipamento("nm_equipamento_novo")
					'if not nm_equipamento_novo is null then
					'	sql = "UPDATE TBL_equipamento SET nm_equipamento_novo = '"&nm_equipamento&"' WHERE (cd_codigo = '"&cd_equipamento&"')"
					'	dbconn.execute(sql)
					'end if
					
				%>
			
				<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>PART NUMBER</b></td>
					<td>&nbsp;<b>EQUIPAMENTO</b></td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>						
				</tr>
				<%end if%>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">			
					<td valign="middle" align="right"><%=num_lista%></td>
					<td valign="middle">&nbsp;<%=cd_pn%></td>
					<td valign="middle">&nbsp;<%=nm_equipamento%></td>
					<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
					<td id="no_print" valign="middle">					
					<img src="imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_equipamento%>')" type="button" value="A" title="Excluir">
					<img src="imagens/undo.gif" onclick="javascript:JsReativar('<%=cd_equipamento%>')" type="button" value="A" title="Reativar"></td>	
					<td><%=nm_equipamento_novo%></td>
						<%num_lista = num_lista + 1
						rs_equipamento.movenext
						loop%>
				</tr>					
			</table>
			<%end if%>
		</span>
	</td>
</tr>
<!--tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
	<td><img src="../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>	
	<td><img src="../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
	<td><img src="../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td>
	<td><img src="../imagens/blackdot.gif" alt="" width="155" height="1" border="0"></td>	
	<td id="no_print"><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
</tr-->
<tr><td colspan="100%">&nbsp;</td></tr>

</form>
</table>

<!--/td></tr></table-->
<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=equipamento&cd_equipamento='+cod1+'&modo=equipamento&status=2&acao=excluir');
	  }
}
function JsDesativar(cod1)
{
  if (confirm ("Tem certeza que deseja tornar o equipamento inativo?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=equipamento&cd_equipamento='+cod1+'&modo=equipamento&status=0&acao=excluir');
	  }
}
function JsReativar(cod1)
{
  if (confirm ("Tem certeza que deseja reativar o cadastro do equipamento?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=equipamento&cd_equipamento='+cod1+'&modo=equipamento&status=0&acao=editar');
	  }
}
</SCRIPT>
</body>
</html>
