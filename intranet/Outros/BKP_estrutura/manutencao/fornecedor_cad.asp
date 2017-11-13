<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"--> 
<!--#include file="../includes/inc_area_restrita.asp"-->   
<!--#include file="../css/estilo.asp"-->  
<html>
<head>
	<title>Untitled</title>
</head>
<style type="text/css">
<!--
.hide{visibility:hidden; display: none;}
.show{visibility:visible; display:block;}
-->
</style>
<%cd_codigo = request("cd_codigo")
	if cd_codigo = "" then
	acao = "inserir"
	Else
	acao = "editar"
	end if
action = request("action")
tipo = request("tipo")
janela = request("janela")

if janela = "pop_up" then
	janela = "hide"
end if

strsql ="SELECT * FROM TBL_fornecedor WHERE cd_codigo='"&cd_codigo&"'"' up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
Set rs = dbconn.execute(strsql)
do while not rs.eof
cd_codigo = rs("cd_codigo")
nm_fornecedor = rs("nm_fornecedor")
nm_endereco = rs("nm_endereco")
nm_numero = rs("nm_numero")
nm_bairro = rs("nm_bairro")
nm_cidade = rs("nm_cidade")
nm_estado = rs("nm_estado")
nm_contato = rs("nm_contato")
cd_prazo = rs("cd_prazo")

rs.movenext
loop
%>

<body>
<%=janela%>
<table border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>&nbsp;</td><td><br>

<form action="../adm/adm2_acao.asp" method="post">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
<input type="hidden" name="tipo" value="forn">
<input type="hidden" name="janela" value="<%=janela%>">
<table border="0" align="center">
<tr id="no_print">		
	<td class="txt_cinza" colspan="12" bgcolor="#C0C0C0">
	<b><font color="red">Fornecedores</font></b></td>
</tr>
<tr>
    <td>COD.:</td>
    <td><b><%=zero(cd_codigo)%></b>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr>
    <td>Fornecedor:</td>
    <td colspan="5"><input type="text" name="nm_fornecedor" value="<%=nm_fornecedor%>" size="60" maxlength="200"></td>
</tr>
<tr>
    <td>Endereço:</td>
    <td><input type="text" name="nm_endereco" value="<%=nm_endereco%>" size="30" maxlength="200"></td>
    <td>Nº:</td>
    <td><input type="text" name="nm_numero" value="<%=nm_numero%>" size="6" maxlength="6"></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr>
    <td>Bairro:</td>
    <td><input type="text" name="nm_bairro" value="<%=nm_bairro%>" size="30" maxlength="50"></td>
    <td>Cidade:</td>
    <td><input type="text" name="nm_cidade" value="<%=nm_cidade%>" size="15" maxlength="50"></td>
    <td>UF</td>
    <td><input type="text" name="nm_estado" value="<%=nm_estado%>" size="2" maxlength="2"></td>
</tr>
<tr>
    <td>Contato:</td>
    <td><input type="text" name="nm_contato" value="<%=nm_contato%>" size="30" maxlength="50"></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr>
    <td>Prazo:</td>
    <td><input type="text" name="cd_prazo" size="3" maxlength="3" value="<%=cd_prazo%>"> dias</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr>
    <td><input type="submit" name="ok" value="OK"></td>
    <td>&nbsp;</td>
    <td><a href="fornecedores.asp?cd_codigo="><img src="../imagens/ic_novo.gif" itle="Novo" border="0"></a></td>
    <td align="right">&nbsp;<%if cd_codigo <> "" then%><img src="../imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_codigo%>')" type="button" value="A" title="Apagar"><%end if%></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr><td colspan="100%"><font color="red"><b><%=action%></b></font></td></tr>
</table>
</form>
 <%'=tipo%>
<table cellspacing="0" cellpadding="1" border="1"  class="<%=janela%>">
<tr bgcolor="silver">
    <td align="center"><b>Cód.</b></td>
    <td align="center"><b>Empresa</b></td>
    <td align="center"><b>Prazo</b></td>
    <td align="center"><b>&nbsp;</b></td>
	<td rowspan="100%">&nbsp;</td>
	<td align="center"><b>Cód.</b></td>
    <td align="center"><b>Empresa</b></td>
    <td align="center"><b>Prazo</b></td>
    <td align="center"><b>&nbsp;</b></td>
</tr>
<tr>
	<td><img src="imagens/px.gif" alt="" width="25" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="250" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="30" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="20" height="1" border="0"></td>
	<!--td><img src="imagens/px.gif" alt="" width="30" height="1" border="0"></td-->
	<td><img src="imagens/px.gif" alt="" width="25" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="250" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="30" height="2" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
</tr>
<%xsql ="SELECT * FROM TBL_fornecedor order by nm_fornecedor"'WHERE cd_codigo='"&cd_avaliacao&"'"' up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
Set rs = dbconn.execute(xsql)%>
<tr><%num = 1

do while not rs.eof
cd_codigo = rs("cd_codigo")
nm_fornecedor = rs("nm_fornecedor")
nm_endereco = rs("nm_endereco")
nm_numero = rs("nm_numero")
nm_bairro = rs("nm_bairro")
nm_cidade = rs("nm_cidade")
nm_estado = rs("nm_estado")
nm_contato = rs("nm_contato")
cd_prazo = rs("cd_prazo")
%>
    <td align="center">&nbsp;<b><%=zero(cd_codigo)%></b></td>
    <td>&nbsp;<a href="fornecedores.asp?cd_codigo=<%=cd_codigo%>"><%=nm_fornecedor%></a></td>
    <td align="center">&nbsp;<%=cd_prazo%></td>
    <td align="center"><img src="imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_codigo%>')" type="button" value="A" title="Apagar"></td>
<%num = num + 1

if num = 3 Then
num = 1%>
</tr><tr>
<%end if


rs.movenext
loop
%>	
</tr>
<tr bgcolor="#c0c0c0" id="no_print">
	<td><img src="imagens/px.gif" alt="" width="25" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="250" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="30" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="20" height="1" border="0"></td>
	<!--td><img src="imagens/px.gif" alt="" width="30" height="1" border="0"></td-->
	<td><img src="imagens/px.gif" alt="" width="25" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="250" height="2" border="0"></td>
	<td><img src="imagens/px.gif" alt="" width="30" height="2" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
</tr>

</table>
</td></tr></table><br>&nbsp;



<SCRIPT LANGUAGE="javascript">			
	function JsDelete(cod)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('adm/adm2_acao.asp?cd_codigo='+cod+'&acao=excluir&tipo=forn');
	  }
}
</SCRIPT>
</body>
</html>