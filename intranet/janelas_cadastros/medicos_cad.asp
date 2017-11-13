<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../css/estilo.asp"-->  

<script language="JavaScript">
function validar_medico(shipinfo){
	if (shipinfo.cd_crm.value.length==""){window.alert ("O crm deve ser digitado.");shipinfo.cd_crm.focus();return (false);}
	if (shipinfo.nm_medico.value.length==""){window.alert ("O Nome do médico deve ser digitado.");shipinfo.nm_medico.focus();return (false);}	
	return (true);}
	
	// onsubmit="return validar_medico(document.Form);"
</script>

<%
'cd_codigo = request("cd_codigo")
'campo = request("campo")
'list = request("list")
'action = request("action")
acao = "inserir"
janela = request("janela")
metodo = request("metodo")
%>

<body>
<table width="445" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr bgcolor="#f2f2f2">
		<td>&nbsp;</td><td><br>

<form action="../acoes/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_medico(document.Form);">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
<input type="hidden" name="modo" value="medico">
<input type="hidden" name="janela" value="<%=janela%>">
<input type="hidden" name="metodo" value="<%=metodo%>">
<table border="0" align="center">
<tr id="no_print">		
	<td class="txt_cinza" colspan="12" bgcolor="#C0C0C0">
	<b>Medicos</b></td>
</tr>
<tr bgcolor="#f5f5f5">
	<td  align="left">CRM</td>
	<td><input type="text" name="cd_crm" size="10" maxlength="50" class="inputs" value=""></td>
</tr>
<tr bgcolor="#f5f5f5">	
	<td  align="left">Médico</td>
	<td><input type="text" name="nm_medico" size="45" maxlength="50" class="inputs" value=""></td>
	
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