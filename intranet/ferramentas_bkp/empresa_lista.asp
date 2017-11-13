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
	<title>Cadastro de Empresas</title>
</head>
<script language="JavaScript">
	function validar_equipamento(shipinfo){
			if (shipinfo.cd_tipo.value=="0"){window.alert ("O tipo de equipamento deve ser informado.");shipinfo.cd_tipo.focus();return (false);}
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
<!--#include file="../ferramentas/js/ferramentas.js"-->

<body onload="foco()"><br>

<form action="../acoes/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_eempresa(document.Form);">
<!--input type="hidden" name="janela" value="1"-->
<input type="hidden" name="modo" value="empresa">
<input type="hidden" name="cd_empresa" value="<%=cd_empresa%>">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="status" value="<%=status%>">
<%if status = 0 then
	atividade = "#ccffff"
Elseif status = 1 then
	atividade = "#ffcc99"
end if%>






<table class="txt" border="0" width="500" align="center">
<tr bgcolor="<%=atividade%>">
	<td align="center" colspan="100%" style="color:red;font-size=13px;"><b>Listagem de empresas</b></td>
</tr>


<tr>
	<td >&nbsp;<%if status <> 0 or status = "" then%><b>BUSCA: <input type="text" name="busca_empresa" class="inputs" size="10" maxlength="30" value="<%=cd_empresa%>" onKeyup="mostra_empresa();"><%end if%></td>
	<td align="center"><a href="javascript:void(0);" onClick="window.open('ferramentas/empresa_cad.asp','itens_nomes','width=500,height=270')" return false;">Novo</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		(Ir para: <%if status <= 0 then%>
		<a href="ferramentas.asp?tipo=equipamento&status=1">Ativos</a>
		<%elseif status >= 1  or status = "" then%>
		<a href="ferramentas.asp?tipo=equipamento&status=0">Inativos</a>
	<%end if%>)	
	</td>
</tr>

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
		<span id='divEmpresa'>
			<%if status = 1 then%>			
			<table border="0" width="500" align="center">
				
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>EMPRESA</b></td>
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
			strsql ="SELECT * FROM TBL_empresa where cd_status = 0 ORDER BY nm_empresa"
			Set rs_empresa = dbconn.execute(strsql)
			num_lista = 1%>
			<table border="0" width="500">
			<%Do while Not rs_empresa.EOF
				
				'cd_pn = rs_empresa("cd_pn")
				cd_empresa = rs_empresa("cd_codigo")
				nm_empresa = rs_empresa("nm_empresa")
				cd_status = rs_empresa("cd_status")%>
			
				<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>PART NUMBER</b></td>
					<td>&nbsp;<b>EQUIPAMENTO</b></td>
					<td>&nbsp;<b>Nome antigo</b></td>
					<td>&nbsp;<b>TIPO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>				
				</tr>
				<%end if%>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">			
					<td valign="middle" align="right"><%=num_lista%></td>
					<td valign="middle">&nbsp;<%=cd_pn%></td>
					<td valign="middle">&nbsp;<%=nm_empresa%></td>
					<td valign="middle" style="color:silver;">&nbsp;<%=nm_eempresa%></td>
					<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
					<td valign="middle">&nbsp;<%=cd_tipo%></td>
					<td id="no_print" valign="middle">
					<img src="imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_empresa%>')" type="button" value="A" title="Excluir">
					<img src="imagens/undo.gif" onclick="javascript:JsReativar('<%=cd_empresa%>')" type="button" value="A" title="Reativar"></td>	
						<%num_lista = num_lista + 1
						rs_empresa.movenext
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
