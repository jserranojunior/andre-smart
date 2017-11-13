<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%cd_user = session("cd_codigo")
pasta_loc = "ferramentas\"
arquivo_loc = "aparelhos_cad.asp"


strcd_aparelho = request("cd_aparelho")

status = request("status")
	if status = "" Then status = 0
acao = request("acao")
	if acao = "" Then
	acao = "inserir"
	Else
	acao = "editar"
	end if

mensagem = request("mensagem")
%>

<head>
	<title>Cadastro de Aparelhos Celulares</title>
</head>
<script language="JavaScript">
	function validar_aparelho(shipinfo){
			//if (shipinfo.cd_tipo.value=="0"){window.alert ("O tipo de equipamento deve ser informado.");shipinfo.cd_tipo.focus();return (false);}
		//	if (shipinfo.nm_fornecedor.value.length==""){window.alert ("O fornecedor deve ser digitado.");shipinfo.nm_fornecedor.focus();return (false);}
		//	if (shipinfo.existe.value.length==""){window.alert ("O fornecedor já existe, digite um nome diferente.");shipinfo.nm_fornecedor.focus();return (false);}
		//return (true);}
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

<body><!-- onload="foco()"-->
<!--#include file="../includes/arquivo_loc.asp"-->
<table class="txt" align="center" id="no_print" width="300" style="font-family:verdana; font-size:10px;">
<tr>
	<td align="center" colspan="2" align="center"><b>CADASTRO DE APARELHOS CELULARES</b></td>
</tr>
<form action="acao/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_aparelho(document.Form);">
<!--input type="hidden" name="janela" value="1"-->
<input type="hidden" name="modo" value="aparelho">
<input type="hidden" name="cd_codigo" value="<%=strcd_aparelho%>">
<input type="hidden" name="acao" value="<%=acao%>">
<!--input type="hidden" name="cd_status" value="<%=status%>"-->
<input type="hidden" name="jan" value="1">
<%

if status = "" OR status = 0 then
	atividade = "#ccffff"
Elseif status = 1 then
	atividade = "#ffcc99"
end if%>

<%if strcd_aparelho <> "" Then

	strsql ="SELECT * FROM View_aparelhos_comunicacao WHERE cd_codigo = "&strcd_aparelho&" "
		Set rs = dbconn.execute(strsql)
		while not rs.EOF
			cd_aparelho = rs("cd_codigo")
			nr_numero = trim(rs("nr_numero"))
				if nr_numero <> "" Then
				nr_numero_pref = mid(nr_numero,1,1)
				nr_numero_1 = mid(nr_numero,2,4)
				nr_numero = nr_numero_pref&"-"&nr_numero_1&"-"&mid(nr_numero,6,4)
				end if
			cd_modelo = rs("cd_modelo")
			nm_modelo = rs("nm_modelo")
			nm_apelido = rs("nm_apelido")
			nr_imei_1 = rs("nr_imei_1")
			nr_imei_2 = rs("nr_imei_2")
			cd_status = rs("cd_status")
			
		rs.movenext
		wend
				
		rs.close
		Set rs = nothing
end if
%>
<tr bgcolor=#cococo>
	<td align="center" colspan="2"><img src="imagens/px.gif"  height=1></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Número:</td>
    <td><input type="text" name="nr_numero" size="15" value="<%=trim(nr_numero)%>" class="inputs" <%if acao <> "editar" then%>onKeyup="verifica_nome();"<%end if%>></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Imei 1:</td>
	<td><input type="text" name="nr_imei_1" size="20" value="<%=trim(nr_imei_1)%>" class="inputs"></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Imei 2:</td>
	<td><input type="text" name="nr_imei_2" size="20" value="<%=trim(nr_imei_2)%>" class="inputs"></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Modelo:</td>
	<td><select name="cd_modelo">
			<option value="1" SELECTED></option>
			<%strsql ="SELECT * FROM TBL_aparelhos_comunicacao_modelos WHERE cd_status = 1 "
				Set rs = dbconn.execute(strsql)
				while not rs.EOF
					cd_mod = rs("cd_codigo")
					nm_mod = rs("nm_modelo")
				%>
					<option value="<%=cd_mod%>" <%if int(cd_modelo)=int(cd_mod) then response.write("SELECTED")%>><%=nm_mod%></option>	
				<%rs.movenext
				wend
				
				rs.close
				Set rs = nothing%>
			
		</select>
	</td>
</tr>
<%if strcd_aparelho <> "" Then%>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Apelido:</td>
	<td><%=nm_apelido%></td>
</tr>
<%end if%>
<%if strcd_aparelho <> "" Then%>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;Status:</td>
	<td><input type="radio" name="cd_status" value="1" <%if cd_status = 1 then response.write("checked")%>> Ativo &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="radio" name="cd_status" value="0" <%if cd_status = 0 then response.write("checked")%>> Inativo</td>
</tr>
<%end if%>
<tr bgcolor="<%=atividade%>">
    <td colspan="2">&nbsp;</td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td align="left"> &nbsp; 
		<span id='divEquip_novo'>
			<%if acao = "inserir" then%>
				<input type="submit" name="envia" value="OK">
			<%elseif acao = "editar" then%>
				<input type="submit" name="envia" value="Altera">
			<%end if%>
		</span>
	</td>
	<td align="right">&nbsp;
		<%if cd_aparelho <> "" Then%>
		<img src="../imagens/ic_cancela.gif" onclick="javascript:JsDelete('<%=strcd_aparelho%>')" type="button" value="A" title="Excluir">
		<%end if%>			
	</td>
</tr>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="330" height="1" border="0"></td>
</tr>
</table>
<!--/td></tr></table-->


</body>
</html>
<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('acao/cadastros_acao.asp?cd_codigo='+cod1+'&modo=aparelho&acao=excluir');
	  }
}
</SCRIPT>
