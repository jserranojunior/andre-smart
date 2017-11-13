<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../css/estilo.asp"-->  

<script language="JavaScript">
function validar_procedimento(shipinfo){
	if (shipinfo.cod_amb.value.length==""){window.alert ("Digite o código AMB.");shipinfo.cod_amb.focus();return (false);}
	if (shipinfo.nm_procedimento.value.length==""){window.alert ("Digite o procedimento.");shipinfo.nm_procedimento.focus();return (false);}
	return (true);}
</script>

<%
cd_codigo = request("cd_codigo")
campo = request("campo")
list = request("list")
action = request("action")
acao = "inserir"
janela = request("janela")
metodo = request("metodo")

data_atual = month(now)&"/"&day(now)&"/"&year(now)
%>

<body>
<%'=metodo%>
<table width="445" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr bgcolor="#f2f2f2">
		<td>&nbsp;</td><td><br>

<form action="../acoes/cadastros_acao.asp" method="post" onsubmit="return validar_procedimento(document.Form);" name="Form">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="modo" value="procedimento">
<input type="hidden" name="janela" value="<%=janela%>">
<input type="hidden" name="metodo" value="<%=metodo%>">
<table border="0" align="center">
<tr id="no_print">		
	<td class="txt_cinza" colspan="12" bgcolor="#C0C0C0">
	<b>Inserir procedimento</b></td>
</tr>
<tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr bgcolor="#f5f5f5">	
	<td  align="left">Cód. AMB</td>
	<td><input type="text" name="cod_amb" size="20" maxlength="20" class="inputs" value=""></td>	
</tr>
<tr bgcolor="#f5f5f5">	
	<td  align="left">Procedimento</td>
	<td><input type="text" name="nm_procedimento" size="45" maxlength="50" class="inputs" value=""></td>
	<input type="hidden" name="cd_status" value="1">
	<input type="hidden" name="dt_datcad" value="<%=data_atual%>">
	<input type="hidden" name="dt_datatu" value="<%=data_atual%>">
</tr>
<tr>
    <td><input type="submit" name="ok" value="OK"></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
</tr>
<tr><td colspan="100%"><font color="red"><b><%=action%></b></font></td></tr>
</table>
</form>