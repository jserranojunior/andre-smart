<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_user = session("cd_codigo")
pasta_loc = "ferramentas\"
arquivo_loc = "aparelhos_lista.asp"


cd_pn = request("cd_pn")
cd_aparelho = request("cd_aparelho")
nr_numero = request("nr_numero")
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
	<title>Cadastro de Aparelhos Celulares</title>
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
<!--#include file="js/ferramentas.js"-->

<body><!-- onload="foco()"><br-->
<!--# include file="../includes/arquivo_loc.asp"-->
<form action="../acoes/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_aparelho(document.Form);">
<!--input type="hidden" name="janela" value="1"-->
<input type="hidden" name="modo" value="aparelhos">
<input type="hidden" name="cd_aparelho" value="<%=cd_aparelho%>">
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
	<td align="center" colspan="4" style="color:red;font-size=13px;"><b>Listagem de Aparelhos Celulares.</b></td>
</tr>
<tr class="no_print">
	<td >&nbsp;<%'if status <> 0 or status = "" then%><b>BUSCA: <input type="text" name="busca_aparelho" class="inputs" size="10" maxlength="30" value="<%=cd_aparelho%>" onKeyup="mostra_aparelho(<%=status%>);"><%'end if%></td>
	<td align="center"><a href="javascript:void(0);" onClick="window.open('ferramentas/aparelhos_cad.asp','itens_nomes','width=500,height=400')" return false;">Novo</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		(Ir para: <%if status = 0 then%>
		<a href="ferramentas.asp?tipo=aparelhos&status=1">Ativos</a>
		<%elseif status = 1  or status = "" then%>
		<a href="ferramentas.asp?tipo=aparelhos&status=0">Inativos</a>
	<%end if%>)	
	</td>
</tr>

<tr>
	<td colspan="100%"><img src="../imagens/px.gif" alt="" width="100%" height="1" border="0"></td>	
</tr>
<tr>
	<td colspan="100%">
		<span id='divAparel'>
			<%if status = 1 then%>			
			<table border="0" width="500" align="center">
				
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>APARELHO</b></td>
					
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
				</tr>
			</table>
			<%elseif status = 0 then
			strsql ="SELECT * FROM TBL_Aparelhos_comunicacao where cd_status = 0 ORDER BY nr_numero"
			Set rs = dbconn.execute(strsql)
			num_lista = 1%>
			<table border="0" width="500">
			<%Do while Not rs.EOF
				
				cd_aparelho = rs("cd_codigo")
				nr_numero = rs("nr_numero")
				cd_status = rs("cd_status")%>
			
				<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center"><b>Nº</b></td>
					<td>&nbsp;<b>APARELHO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					<td><img src="../imagens/apx.gif" alt="" width="200" height="1" border="0"></td>					
				</tr>
				<%end if%>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td valign="middle" align="right"><%=num_lista%></td>
					<td valign="middle">&nbsp;<%=nr_numero%></td>
					<td class="no_print" valign="left">
					<%if status = 0 then%><%=status%>
					<img src="../imagens/ic_cancela.gif" onclick="javascript:JsDelete('<%=cd_aparelho%>')" type="button" value="A" title="Excluir" class="no_print">
					<img src="../imagens/ic_pendente.gif" onclick="javascript:JsAtivaDesativa('<%=cd_aparelho_%>')" type="button" value="A" title="ativa/desativa" class="no_print"><%'=cd_status%>
					<img src="imagens/undo.gif" onclick="javascript:JsReativar('<%=cd_aparelho%>')" type="button" value="A" title="Reativar" class="no_print"></td>
					<%end if%>
						<%num_lista = num_lista + 1
						rs.movenext
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
		document.location.href('ferramentas/acao/cadastros_acao.asp?cd_codigo='+cod1+'&modo=aparelho&acao=excluir');
	  }
}

function JsAtivaDesativa(cod1)
{
  if (confirm ("Tem certeza que deseja ativar?"))
	  {
		document.location.href('ferramentas/acao/cadastros_acao.asp?cd_codigo='+cod1+'&modo=aparelho&status=<%=cd_status%>&acao=ativadesativa');
	  }
}
</SCRIPT>
</body>
</html>
