<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%cd_user = session("cd_codigo")
pasta_loc = "ferramentas\"
arquivo_loc = "fornecedores_cad.asp"


cd_pn = request("cd_pn")
strcd_fornecedor = request("cd_fornecedor")
nm_fornecedor = request("nm_fornecedor")
'cd_tipo = request("cd_tipo")
status = request("status")
	'if status = "2" or status = 1 then
	'if status = 1 then
	'	str_status = " where cd_status = 1 "
	'elseif status = 0 then
	'	str_status = " where cd_status = 0 "
	'end if

acao = request("acao")
	if acao = "" Then
	acao = "inserir"
	Else
	acao = "editar"
	end if

mensagem = request("mensagem")
%>

<head>
	<title>Cadastro de fornecedores</title>
</head>
<script language="JavaScript">
	function validar_fornecedor(shipinfo){
			//if (shipinfo.cd_tipo.value=="0"){window.alert ("O tipo de equipamento deve ser informado.");shipinfo.cd_tipo.focus();return (false);}
			if (shipinfo.nm_fornecedor.value.length==""){window.alert ("O fornecedor deve ser digitado.");shipinfo.nm_fornecedor.focus();return (false);}
			if (shipinfo.existe.value.length==""){window.alert ("O fornecedor já existe, digite um nome diferente.");shipinfo.nm_fornecedor.focus();return (false);}
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
<!--#include file="../includes/arquivo_loc.asp"-->
<table class="txt" align="center" id="no_print" width="300" style="font-family:verdana; font-size:10px;">
<tr>
	<td align="center" colspan="2" align="center"><b>CADASTRO DE FORNECEDORES</b></td>
</tr>
<form action="acao/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_fornecedor(document.Form);">
<!--input type="hidden" name="janela" value="1"-->
<input type="hidden" name="modo" value="fornecedor">
<input type="hidden" name="cd_fornecedor" value="<%=strcd_fornecedor%>">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="cd_status" value="<%=status%>">
<input type="hidden" name="jan" value="1">
<%if status = 0 then
	atividade = "#ccffff"
Elseif status = 1 then
	atividade = "#ffcc99"
end if%>

<%strsql = "Select * From TBL_fornecedor Where cd_codigo = '"&strcd_fornecedor&"'"
	Set rs = dbconn.execute(strsql)
		while not rs.EOF
			cd_fornecedor = rs("cd_codigo")
			nm_fornecedor = rs("nm_fornecedor")
			nm_endereco = rs("nm_endereco")
			nm_numero = rs("nm_numero")
			nm_bairro = rs("nm_bairro")
			nm_cep = rs("nm_cep")
				if nm_cep <> "" Then
					separador = instr(1,nm_cep,"-",1)
					
					nm_cep1 = mid(nm_cep,1,separador-1)
					nm_cep2 = mid(nm_cep,separador+1,len(nm_cep))
				end if
			nm_cidade = rs("nm_cidade")
			nm_estado = rs("nm_estado")
			nm_cnpj = rs("nm_cnpj")
			nm_ie = rs("nm_ie")
			
		rs.movenext
		wend
				
		rs.close
		Set rs = nothing
%>
<tr bgcolor=#cococo>
	<td align="center" colspan="2"><img src="imagens/px.gif"  height=1></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Fornecedor:</td>
    <td><input type="text" name="nm_fornecedor" size="48" value="<%=nm_fornecedor%>" class="inputs" <%if acao <> "editar" then%>onKeyup="verifica_nome();"<%end if%>></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Endereço:</td>
	<td><input type="text" name="nm_endereco" size="30" value="<%=nm_endereco%>" class="inputs"> 
		&nbsp; Nº:<input type="text" name="nm_numero" size="8" value="<%=nm_numero%>" class="inputs"></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Bairro:</td>
	<td><input type="text" name="nm_bairro" size="20" value="<%=nm_bairro%>" class="inputs">
		&nbsp; CEP:<input type="text" name="nm_cep1" size="5" maxlegth="5" value="<%=nm_cep1%>" class="inputs">-<input type="text" name="nm_cep2" size="3" maxlegth="3" value="<%=nm_cep2%>" class="inputs"></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Cidade:</td>
	<td><input type="text" name="nm_cidade" size="20" value="<%=nm_cidade%>" class="inputs">
		&nbsp; &nbsp; UF:<input type="text" name="nm_estado" size="2" value="<%=nm_estado%>" class="inputs"></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;CNPJ:</td>
	<td><input type="text" name="nm_cnpj" size="20" value="<%=nm_cnpj%>" class="inputs">
		&nbsp; &nbsp; IE:<input type="text" name="nm_ie" size="2" value="<%=nm_ie%>" class="inputs"></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td colspan="2">&nbsp;</td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td align="left" colspan="2"> &nbsp; <span id='divEquip_novo'><%if acao = "inserir" then%><input type="submit" name="envia" value="OK"><%elseif acao = "editar" then%><input type="submit" name="envia" value="Altera"><%end if%></span></td>
</tr>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="330" height="1" border="0"></td>
</tr>
</table>
<!--/td></tr></table-->
<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('acoexs/cadastros_acao.asp?tipo=fornecedor&cd_equipamento='+cod1+'&modo=fornecedor&status=1&acao=excluir');
	  }
}
function JsDesativar(cod1)
{
  if (confirm ("Tem certeza que deseja tornar o fornecedor inativo?"))
	  {
		document.location.href('acoexs/cadastros_acao.asp?tipo=fornecedor&cd_fornecedor='+cod1+'&modo=fornecedor&status=0&acao=editar');
	  }
}
function JsReativar(cod1)
{
  if (confirm ("Tem certeza que deseja reativar o cadastro do equipamento?"))
	  {
		document.location.href('acoexs/cadastros_acao.asp?tipo=equipamento&cd_equipamento='+cod1+'&modo=equipamento&status=0&acao=editar');
	  }
}
</SCRIPT>
</body>
</html>
