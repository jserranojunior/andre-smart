<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../css/estilo.asp"-->  

<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>

<%
cd_codigo = request("cd_codigo")
campo = request("campo")
list = request("list")
action = request("action")
acao = "inserir"
janela = request("janela")
metodo = request("metodo")
%>

<body>
<%'=metodo%>
<table width="405" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr bgcolor="#f2f2f2">
		<td>&nbsp;</td><td><br>

<form action="../acoes/cadastros_acao.asp" method="post">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
<input type="hidden" name="modo" value="unidade">
<input type="hidden" name="janela" value="<%=janela%>">
<input type="hidden" name="metodo" value="<%=metodo%>">
<table border="0" align="center">
<tr id="no_print">		
	<td class="txt_cinza" colspan="12" bgcolor="#C0C0C0">
	<b>Unidade</b></td>
</tr>
<tr>
    <td>COD.:</td>
    <td><b><%=zero(cd_codigo)%></b>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr bgcolor="#f5f5f5">
	<td  align="left">Sigla</td>
	<td><input type="text" name="nm_sigla" size="10" maxlength="50" class="inputs" value=""></td>
</tr>
<tr bgcolor="#f5f5f5">	
	<td  align="left">Unidade</td>
	<td><input type="text" name="nm_unidade" size="45" maxlength="50" class="inputs" value=""></td>
	
</tr><tr>
    <td><input type="submit" name="ok" value="OK"></td>
    <td>&nbsp;</td>
    <td><!--a href="fornecedores.asp?cd_codigo="><img src="../imagens/ic_novo.gif" itle="Novo" border="0"></a--></td>
    <td align="right">&nbsp;<%if cd_codigo <> "" then%><img src="../imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_codigo%>')" type="button" value="A" title="Apagar"><%end if%></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr><td colspan="100%"><font color="red"><b><%=action%></b></font></td></tr>
</table>
</form>