<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_convenio = request("cd_convenio")
nm_convenio = request("nm_convenio")
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
	<title>Cadastro de Convenios</title>
</head>
<script language="JavaScript">
	function validar_convenio(shipinfo){
			if (shipinfo.nm_convenio.value.length==""){window.alert ("O nome do convênio deve ser digitado.");shipinfo.nm_convenio.focus();return (false);}	
		return (true);}	
</script>

<script language="javascript">
<!--
VerifiqueTAB=true;
function Mostra(quem, tammax) {
if ( (quem.value.length == tammax) && (VerifiqueTAB) ) { 
var i=0,j=0, indice=-1;
for (i=0; i<document.forms.length; i++) { 
for (j=0; j<document.forms[i].elements.length; j++) { 
if (document.forms[i].elements[j].name == quem.name) { 
indice=i;
break;
} 
} 
if (indice != -1) break; 
} 
for (i=0; i<=document.forms[indice].elements.length; i++) { 
if (document.forms[indice].elements[i].name == quem.name) { 
while ( (document.forms[indice].elements[(i+1)].type == "hidden") &&
(i < document.forms[indice].elements.length) ) { 
i++;
} 
document.forms[indice].elements[(i+1)].focus();
VerifiqueTAB=false;
break;
} 
} 
} 
} 
function PararTAB(quem) { VerifiqueTAB=false; } 
function ChecarTAB() { VerifiqueTAB=true; } 


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

function foco(){
document.form.cd_medico.focus(); }
// -->	
</script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "nm_convenio"; // nome do primeiro campo do site
netscape = "";
ver = navigator.appVersion; len = ver.length;
for(iln = 0; iln < len; iln++) if (ver.charAt(iln) == "(") break;
	netscape = (ver.charAt(iln+1).toUpperCase() != "C");
	function keyDown(DnEvents) {
		// ve quando e o netscape ou IE
		k = (netscape) ? DnEvents.which : window.event.keyCode;
		if (k == 13) { // preciona tecla enter
		if (nextfield == 'done') {
			//alert("viu como funciona?");
			eval('document.form.' + nextfield + '.focus()');
			return false;
			//return true; // envia quando termina os campos
		} else {
			// se existem mais campos vai para o proximo
			eval('document.form.' + nextfield + '.focus()');
			return false;
		}
	}
}
document.onkeydown = keyDown; // work together to analyze keystrokes
if (netscape) document.captureEvents(Event.KEYDOWN|Event.KEYUP);
// End -->
</script>
<!--#include file="../js/ajax.js"-->
<!--#include file="../js/ferramentas.js"-->

<body onload="foco()"><br>
<table class="txt" align="center" id="no_print" width="500">
<form action="../acoes/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_convenio(document.Form);">
<!--input type="hidden" name="janela" value="1"-->
<input type="hidden" name="modo" value="convenio">
<input type="hidden" name="cd_convenio" value="<%=cd_convenio%>">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="status" value="<%=status%>">
<%if status = 0 then
	atividade = "#ffffcc"
Elseif status = 1 then
	atividade = "#ccffcc"
end if%>

<tr bgcolor="<%=atividade%>">
	<td colspan="100%" style="font-size=15px;">&nbsp;<b>Convenios - <font color="red">Manutenção do cadastro de convenios</font></b></td>
</tr>
<tr bgcolor=#cococo>
	<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
</tr>
<tr bgcolor="<%=atividade%>"><td align=center colspan="100%">&nbsp;</td></tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp; Nome:</td>
    <td colspan="3"><input type="text" name="nm_convenio" size="60" value="<%=nm_convenio%>" class="inputs"></td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td colspan="4">&nbsp;</td>
</tr>
<tr bgcolor="<%=atividade%>">
    <td>&nbsp;</td>
    <td colspan="3"><input type="submit" name="envia" value="OK">&nbsp; 
	<a href="ferramentas.asp?tipo=convenio&status=1">Novo</a> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	(Ir para: <%if status = 0then%>
	<a href="ferramentas.asp?tipo=convenio&status=1">Ativos</a>
	<%elseif status = 1  or status = "" then%>
	<a href="ferramentas.asp?tipo=convenio&status=0">Inativos</a>
	<%end if%>)
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(<a href="ferramentas.asp">Ferramentas</a>)
	<br>
	<font color="red"><%=mensagem%></font></td>
</tr>
<tr>
    <td colspan="4"><img src="../imagens/blackdot.gif" alt="" width="500" height="1" border="0"></td>
</tr>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="330" height="1" border="0"></td>
</tr>
</table>
<br><br><!-- LISTAGEM DE convenios -->
<table class="txt" border="0" width="500" align="center">
<%if status <> 0 or status = "" then%>
<tr>
	<td colspan="100%">&nbsp;<b>Busca: <input type="text" name="busca_convenio" class="inputs" size="10" maxlength="30" value="<%=cd_convenio%>" onKeyup="mostra_convenio();"></td>
</tr>
<%end if%>
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
		<span id='divConv'>
			<%if status = 1 then%>			
			<table border="0" width="500" align="center">
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD</b></td>
					<td>&nbsp;<b>CONVÊNIO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>						
				</tr>
			</table>
			<%elseif status = 0 then
			strsql ="SELECT * FROM TBL_convenios where cd_status = 0 ORDER BY nm_convenio"
			Set rs_convenio = dbconn.execute(strsql)
			num_lista = 1%>
			<table border="0" width="500">
			<%Do while Not rs_convenio.EOF
				
				cd_convenio = rs_convenio("cd_codigo")
				nm_convenio = rs_convenio("nm_convenio")
				cd_status = rs_convenio("cd_status")%>
			
				<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>CÓD</b></td>
					<td>&nbsp;<b>CONVÊNIO</b></td>
					<td>&nbsp;</td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>						
				</tr>
				<%end if%>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">			
					<td valign="middle" align="right"><%=num_lista%></td>
					<td valign="middle"><a href="ferramentas.asp?tipo=medico&cd_medico=<%=cd_convenio%>&nm_convenio=<%=nm_convenio%>&acao=editar">&nbsp;<%=cd_convenio%></a></td>
					<td valign="middle"><a href="ferramentas.asp?tipo=medico&cd_medico=<%=cd_convenio%>&nm_convenio=<%=nm_convenio%>&acao=editar">&nbsp;<%=nm_convenio%></td>
					<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
					<td id="no_print" valign="middle">
					<img src="imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_convenio%>')" type="button" value="A" title="Excluir">
					<img src="imagens/undo.gif" onclick="javascript:JsReativar('<%=cd_convenio%>')" type="button" value="A" title="Reativar"></td>	
						<%num_lista = num_lista + 1
						rs_convenio.movenext
						loop%>
				</tr>					
			</table>
			<%end if%>
		</span>
	</td>
</tr>
<!--tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
	<td><img src="../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>	
	<td><img src="../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td>
	<td><img src="../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td>
	<td><img src="../imagens/blackdot.gif" alt="" width="155" height="1" border="0"></td>	
	<td id="no_print"><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
</tr-->
<tr><td colspan="100%">&nbsp;</td></tr>

</form>
</table>

<!--/td></tr></table-->
<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=convenio&cd_convenio='+cod1+'&modo=convenio&status=2&acao=excluir');
	  }
}
function JsDesativar(cod1)
{
  if (confirm ("Tem certeza que deseja tornar o convenio inativo?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=medico&cd_convenio='+cod1+'&modo=convenio&status=0&acao=excluir');
	  }
}
function JsReativar(cod1)
{
  if (confirm ("Tem certeza que deseja reativar o cadastro do convenio?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=convenio&cd_convenio='+cod1+'&modo=convenio&status=0&acao=editar');
	  }
}
</SCRIPT>
</body>
</html>
