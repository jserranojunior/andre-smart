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
	<title>Cadastro de nomes de itens</title>
</head>
<script language="JavaScript">
	function validar_equipamento(shipinfo){
			if (shipinfo.cd_tipo.value=="0"){window.alert ("O tipo de equipamento deve ser informado.");shipinfo.cd_tipo.focus();return (false);}
			if (shipinfo.nm_equipamento.value.length==""){window.alert ("O equipamento deve ser digitado.");shipinfo.nm_equipamento.focus();return (false);}
			if (shipinfo.existe.value.length==""){window.alert ("Digite um nome diferente.");shipinfo.nm_equipamento.focus();return (false);}
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
<!--#include file="../ferramentas/js/ferramentas.js"-->

<body onload="foco()">
<table class="txt" align="center" id="no_print" width="300" style="font-family:verdana; font-size:10px;">
<tr>
	<td coslpan="2" align="center"><b>Cadastro de nomes</b></td>
</tr>
<form action="acao/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_equipamento(document.Form);">
<!--input type="hidden" name="janela" value="1"-->
<input type="hidden" name="modo" value="equipamento">
<input type="hidden" name="cd_equipamento" value="<%=cd_equipamento%>">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="status" value="<%=status%>">
<input type="hidden" name="jan" value="1">
<%if status = 0 then
	atividade = "#ccffff"
Elseif status = 1 then
	atividade = "#ffcc99"
end if%>


<tr bgcolor=#cococo>
	<td align=center colspan="2"><img src="imagens/px.gif"  height=1></td>
</tr>

<tr bgcolor="<%=atividade%>">
    <td>&nbsp; TIPO:</td>
	<td><select name="cd_tipo">
						<option value="0"></option>
						<%strsql ="SELECT * FROM TBL_equipamento_tipo ORDER BY nm_equipamento_tipo"
						Set rs = dbconn.execute(strsql)
						 while not rs.EOF
						 	cd_equipamento_tipo = rs("cd_codigo")
							nm_equipamento_tipo = rs("nm_equipamento_tipo")%>
							<option value="<%=cd_equipamento_tipo%>" <%if int(cd_tipo) = cd_equipamento_tipo then%>Selected<%end if%>><%=nm_equipamento_tipo%></option>
							<%rs.movenext
						 wend%>
				</select></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp; EQUIPAMENTO:</td>
    <td><input type="text" name="nm_equipamento" size="40" value="<%=nm_equipamento%>" class="inputs" <%if acao <> "editar" then%>onKeyup="verifica_nome();"<%end if%>></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td colspan="2">&nbsp;</td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td align="left" colspan="2"> &nbsp; <span id='divEquip_novo'><%if acao = "editar" then%><input type="submit" name="envia" value="Altera"><%end if%></span></td>
</tr>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="330" height="1" border="0"></td>
</tr>
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
