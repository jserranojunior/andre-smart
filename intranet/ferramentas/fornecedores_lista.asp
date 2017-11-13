<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_user = session("cd_codigo")
pasta_loc = "ferramentas\"
arquivo_loc = "fornecedores_lista.asp"


cd_pn = request("cd_pn")
cd_fornecedor = request("cd_fornecedor")
nm_fornecedor = request("nm_fornecedor")
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

<form action="../acoes/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_fornecedor(document.Form);">
<!--input type="hidden" name="janela" value="1"-->
<input type="hidden" name="modo" value="fornecedor">
<input type="hidden" name="cd_fornecedor" value="<%=cd_fornecedor%>">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="status" value="<%=status%>">
<%if status = 0 then
	atividade = "#ccffff"
Elseif status = 1 then
	atividade = "#ffcc00"
end if%>
<!--#include file="../includes/arquivo_loc.asp"-->
<table class="txt" border="0" width="500" align="center">
<tr bgcolor="<%=atividade%>">
	<td align="center" colspan="4" style="color:red;font-size=13px;"><b>Listagem de fornecedores.</b></td>
</tr>
<tr>
	<td >&nbsp;<%'if status <> 0 or status = "" then%><b>BUSCA: <input type="text" name="busca_fornecedor" class="inputs" size="10" maxlength="30" value="<%=cd_fornecedor%>" onKeyup="mostra_fornecedor(<%=status%>);"><%'end if%></td>
	<td align="center"><a href="javascript:void(0);" onClick="window.open('ferramentas/fornecedores_cad.asp','itens_nomes','width=500,height=400')" return false;">Novo</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		(Ir para: <%if status = 0 then%>
		<a href="ferramentas.asp?tipo=fornecedores&status=1">Ativos</a>
		<%elseif status = 1  or status = "" then%>
		<a href="ferramentas.asp?tipo=fornecedores&status=0">Inativos</a>
	<%end if%>)	
	</td>
</tr>

<tr>
	<td colspan="100%"><img src="../imagens/px.gif" alt="" width="100%" height="1" border="0"></td>	
</tr>
<tr>
	<td colspan="100%">
		<span id='divForn'>
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
					<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>						
				</tr>
			</table>
			<%elseif status = 0 then
			strsql ="SELECT * FROM TBL_fornecedor where cd_status = 0 ORDER BY nm_fornecedor"
			Set rs_fornecedor = dbconn.execute(strsql)
			num_lista = 1%>
			<table border="0" width="500">
			<%Do while Not rs_fornecedor.EOF
				
				cd_fornecedor = rs_fornecedor("cd_codigo")
				nm_fornecedor = rs_fornecedor("nm_fornecedor")
				cd_status = rs_fornecedor("cd_status")%>
			
				<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>FORNECEDOR</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>					
				</tr>
				<%end if%>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td valign="middle" align="right"><%=num_lista%></td>
					<td valign="middle">&nbsp;<%=nm_fornecedor%></td>
					<td id="no_print" valign="middle">
					<img src="imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_fornecedor%>')" type="button" value="A" title="Excluir">
					<img src="imagens/undo.gif" onclick="javascript:JsReativar('<%=cd_fornecedor%>')" type="button" value="A" title="Reativar"></td>
						<%num_lista = num_lista + 1
						rs_fornecedor.movenext
						loop%>
				</tr>
			</table>
			<%end if%>
		</span>
	</td>
</tr>
<tr><td colspan="100%">&nbsp;</td></tr>

</form>
</table>

<!--/td></tr></table-->
<SCRIPT LANGUAGE="javascript">
/*function JsDelete(cod1)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('../acoes/cadastros_acao.asp?tipo=fornecedor&cd_fornecedor='+cod1+'&modo=fornecedor&status=2&acao=excluir');
	  }
}
function JsDesativar(cod1)
{
  if (confirm ("Tem certeza que deseja tornar o fornecedor inativo?"))
	  {
		document.location.href('ferramentas/acao/cadastros_acao.asp?tipo=fornecedor&cd_fornecedor='+cod1+'&modo=fornecedor&status=0&acao=desativar');
	  }
}
function JsReativar(cod1)
{
  if (confirm ("Tem certeza que deseja reativar o cadastro do fornecedor?"))
	  {
		document.location.href('ferramentas/acao/cadastros_acao.asp?tipo=fornecedor&cd_fornecedor='+cod1+'&modo=fornecedor&cd_status=0&acao=ativar');
	  }
}*/
</SCRIPT>
</body>
</html>
