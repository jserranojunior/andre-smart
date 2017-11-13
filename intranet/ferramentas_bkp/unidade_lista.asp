<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_pn = request("cd_pn")
cd_equipamento = request("cd_equipamento")
nm_equipamento = request("nm_equipamento")
cd_tipo = request("cd_tipo")
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
<!--#include file="../ferramentas/js/ferramentas.js"-->

<body onload="foco()"><br>

<form action="../acoes/cadastros_acao.asp" method="post" name="Form">
<!--input type="hidden" name="janela" value="1"-->
<input type="hidden" name="modo" value="equipamento">
<input type="hidden" name="cd_equipamento" value="<%=cd_equipamento%>">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="status" value="<%=status%>">
<%if status = 0 then
	atividade = "#ccffff"
Elseif status = 1 then
	atividade = "#ffcc00"
end if%>






<table class="txt" border="0" width="500" align="center">
<tr bgcolor="<%=atividade%>">
	<td align="center" colspan="100%" style="color:red;font-size=13px;"><b>Listagem de Unidades</b></td>
</tr>
<tr>
	<td >&nbsp;<%'if status <> 0 or status = "" then%><b>BUSCA: <input type="text" name="busca_unidade" class="inputs" size="10" maxlength="30" value="<%=cd_equipamento%>" onKeyup="mostra_unidade(<%=status%>);"><%'end if%></td>
	<td align="center"><a href="javascript:void(0);" onClick="window.open('ferramentas/unidade_cad.asp','unidade_nomes','width=500,height=450')" return false;">Novo</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		(Ir para: <%if status = 0 then%>
		<a href="ferramentas.asp?tipo=unidade&status=1">Ativos</a>
		<%elseif status = 1  or status = "" then%>
		<a href="ferramentas.asp?tipo=unidade&status=0">Inativos</a>
	<%end if%>)	
	</td>
</tr>
<tr>
	<td colspan="100%"><img src="../imagens/px.gif" alt="" width="100%" height="1" border="0"></td>	
</tr>

<tr>
	<td colspan="100%">
		<span id='divUnid'>
			<%if status = 1 then%>			
			<table border="0" width="500" align="center">
				
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>NOME</b></td>
					<td>&nbsp;<b>TIPO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>						
				</tr>
			</table>
			<%elseif status = 0 then
			strsql ="SELECT * FROM View_unidades where cd_status = 0 ORDER BY nm_unidade"
			Set rs_unid = dbconn.execute(strsql)
			num_lista = 1%>
			<table border="0" width="500">
			<%Do while Not rs_unid.EOF
				
				cd_unidade = rs_unid("cd_codigo")
				nm_unidade = rs_unid("nm_unidade")
				cd_status = rs_unid("cd_status")
				nm_unidade_tipo = rs_unid("nm_unidade_tipo")%>
			
				<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD.</b></td>
					<td>&nbsp;<b>UNIDADE</b></td>
					<td>&nbsp;<b>TIPO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>			
				</tr>
				<%end if%>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">			
					<td valign="middle" align="right"><%=num_lista%></td>
					<td valign="middle">&nbsp;<%=zero(cd_unidade)%></td>
					<td valign="middle">&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/unidade_cad.asp?cd_unidade=<%=cd_unidade%>&status=<%=cd_status%>&acao=editar&jan=1','Altera_nome','width=500,height=400')" return false;">&nbsp;<%=nm_unidade%></a></td>
					<td valign="middle">&nbsp;<%=nm_unidade_tipo%></td>
					<td align="center" valign="middle" id="no_print">
					<img src="imagens/ic_del.gif" border="0" onclick="javascript:JsDelete('<%=cd_unidade%>')" type="button" value="A" title="Excluir"> &nbsp; &nbsp;
					<img src="imagens/undo2.gif" border="0" onclick="javascript:JsReativar('<%=cd_unidade%>')" type="button" value="A" title="Reativar"></td>	
						<%num_lista = num_lista + 1
						rs_unid.movenext
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
		document.location.href('ferramentas/acao/cadastros_acao.asp?tipo=unidade&cd_unidade='+cod1+'&modo=unidade&acao=excluir');
	  }
}

function JsDesativar(cod1)
{
  if (confirm ("Tem certeza que deseja tornar a unidade inativa?"))
	  {
		document.location.href('ferramentas/acao/cadastros_acao.asp?tipo=unidade&cd_unidade='+cod1+'&modo=unidade&cd_status=0&acao=inativar');
	  }
}
function JsReativar(cod1)
{
  if (confirm ("Tem certeza que deseja reativar o cadastro da unidade?"))
	  {
		document.location.href('ferramentas/acao/cadastros_acao.asp?tipo=unidade&cd_unidade='+cod1+'&modo=unidade&cd_status=1&acao=ativar');
	  }
}
</SCRIPT>
</body>
</html>
