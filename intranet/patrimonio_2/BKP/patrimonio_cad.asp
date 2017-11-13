<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_pat_codigo = request("cd_pat_codigo")
cd_patrimonio = request("cd_patrimonio")
cd_equipamento =  request("cd_equipamento")
nm_equipamento = request("nm_equipamento")
cd_marca =  request("cd_marca")
nm_marca =  request("nm_marca")
cd_ns =  request("cd_ns")
cd_pn =  request("cd_pn")
cd_unidade =  request("cd_unidade")
nm_sigla =  request("nm_sigla")
dt_data =  request("dt_data")
ordem = request("ordem")
	if ordem = "" Then
	ordem = "cd_patrimonio"
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
	<title>Inventário</title>
</head>
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
document.form.cd_patrimonio.focus(); }
// -->	
</script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "cd_patrimonio"; // nome do primeiro campo do site
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
<script language="JavaScript">
function validar_patrimonio(shipinfo){
	if (shipinfo.cd_equipamento.value.length==""){window.alert ("Selecione um equipamento.");shipinfo.cd_equipamento.focus();return (false);
	} else {
	if (shipinfo.cd_marca.value.length==""){window.alert ("Selecione a marca do produto.");shipinfo.cd_marca.focus();return (false);
	} else {
	if (shipinfo.cd_dia.value.length==""){window.alert ("O campo Dia não foi preenchido.");shipinfo.cd_dia.focus();return (false);
	} else {
	if (shipinfo.cd_mes.value.length==""){window.alert ("O campo Mês não foi preenchido.");shipinfo.cd_mes.focus();return (false);
	} else {
	if (shipinfo.cd_ano.value.length==""){window.alert ("O campo Ano não foi preenchido.");shipinfo.cd_ano.focus();return (false);
	} else {
	}}}}}return (true);}
</script>

<body onload="foco()"><br>
<table align="center">
<tr><td>&nbsp;&nbsp;&nbsp;</td>
<td align="center">
<!--table cellspacing="0" cellpadding="0" border="1" rules="groups" class="txt" id="no_print"-->
<table class="txt" id="no_print">
<form action="../acoes/patrimonio_acao.asp" method="post" name="form" onsubmit="return validar_patrimonio(document.form);" onSubmit="return checa(this);">
<input type="hidden" name="janela" value="1">
<input type="hidden" name="cd_tipo" value="patrimonio">
<input type="hidden" name="cd_pat_codigo" value="<%=cd_pat_codigo%>">
<input type="hidden" name="acao" value="<%=acao%>">
<tr>
	<td colspan="100%">&nbsp;<b>Inventário - <font color="red">Cadastro de equipamentos/instrumentos</font></b></td>
</tr>

<tr bgcolor=#cococo>
	<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
</tr>
<tr><td align=center colspan="100%">&nbsp;</td></tr>
<tr>
    <td>&nbsp;<b> Nº Patrimônio:</b></td>
    <td colspan="3"><input type="text" name="cd_patrimonio" value="<%=cd_patrimonio%>" class="inputs" onFocus="nextfield ='cd_pn';"></td>
</tr>
<tr>
    <td>&nbsp; Cód. Fabricante:</td>
    <td colspan="3"><input type="text" name="cd_pn" value="<%=cd_pn%>" class="inputs" onFocus="nextfield ='cd_equipamento';"></td>
</tr>
<tr>
    <td>&nbsp; Equip/Instr: </td>
		<%	selecao = " top 100 percent * "
			tabela = " TBL_equipamento"
			condicoes = ""
			criterios = " Order by nm_equipamento"
			
			strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
		  	Set rs_equip = dbconn.execute(strsql)%>
    <td colspan="3"><select name="cd_equipamento" class="inputs" onFocus="nextfield ='cd_marca';">
					<option value=""></option>
					<%Do while Not rs_equip.EOF
					if int(cd_equipamento) = rs_equip("cd_codigo") then%><%eqp_check = "selected"%><%end if%>
					<option value="<%=rs_equip("cd_codigo")%>" <%=eqp_check%>><%=rs_equip("nm_equipamento")%></option>
					<%rs_equip.movenext
					eqp_check = ""
					Loop%>
					</select>
					<a href="javascript:;" onclick="window.open('janelas_cadastros/equipamento_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a>
	</td>
								
</tr>
<tr>
	<td>&nbsp; Marca:</td>
	<%
	selecao = " top 100 percent * "
	tabela = " TBL_marca"
	condicoes = ""
	criterios = " Order by nm_marca"
	
	strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
  	Set rs_marca = dbconn.execute(strsql)%>
	<td><select name="cd_marca" size="1" class="inputs" onFocus="nextfield ='cd_ns';">
	<option value=""></option>
	<%Do while Not rs_marca.EOF
	if int(cd_marca) = rs_marca("cd_codigo") then%><%marca_check = "selected"%><%end if%>
	<option value="<%=rs_marca("cd_codigo")%>" <%=marca_check%>><%=rs_marca("nm_marca")%></option>	
	<%rs_marca.movenext
	marca_check = ""
  	Loop%>
	</td>
</tr>
<tr>
    <td>&nbsp; Num. Série:</td>
    <td colspan="3"><input type="text" name="cd_ns" size="30" value="<%=cd_ns%>" class="inputs" onFocus="nextfield ='cd_unidade';">
	</td>
</tr>
<tr>
    <td>&nbsp; Unidade:</td>
		<%strsql ="up_ADM_unidades_lista"
		Set rs_uni = dbconn.execute(strsql)%>
    <td colspan="3"><select name="cd_unidade" class="inputs"onFocus="nextfield ='cd_dia';">
					<option value=""></option>
					<%Do While Not rs_uni.eof
					if int(cd_unidade) = rs_uni("cd_codigo") then%><%unidade_check = "selected"%><%end if%>
					<option value="<%=rs_uni("cd_codigo")%>" <%=unidade_check%>><%=rs_uni("nm_unidade")%></option>
					<%rs_uni.movenext
					unidade_check = ""
					loop%>
					</select>
					<a href="javascript:;" onclick="window.open('adm/unidade_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a>
	</td>
</tr>


<tr>
    <td>&nbsp; Data de compra:</td>
					<%if cd_pat_codigo <> "" Then%>
						<%dt_dia = day(dt_data)%>
						<%dt_mes = month(dt_data)%>
						<%dt_ano = year(dt_data)%>
					<%end if%>
    <td colspan="3"><input type="text" name="cd_dia" size="2" maxlength="2" value="<%=dt_dia%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" onFocus="nextfield ='cd_mes';">/
					<input type="text" name="cd_mes" size="2" maxlength="2" value="<%=dt_mes%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" onFocus="nextfield ='cd_ano';">/
					<input type="text" name="cd_ano" size="4" maxlength="4" value="<%=dt_ano%>" class="inputs" onFocus="nextfield ='done';"></td>
</tr>
<tr>
    <td colspan="4">&nbsp;</td>
</tr>
<tr>
    <td>&nbsp;</td>
    <td colspan="3"><input type="submit" name="envia" value="OK">&nbsp; <a href="patrimonio.asp?tipo=cadastro">Novo</a> &nbsp; <a href="patrimonio.asp?tipo=descarte">Descarte</a><br>
	<font color="red"><%=mensagem%></font></td>
</tr>
<tr>
    <td colspan="4">&nbsp;</td>
</tr>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="330" height="1" border="0"></td>
</tr>
</table>
<br><!-- LISTAGEM DO PATRIMONIO -->
<table class="txt" border="0" width="600">
<tr bgcolor="#C0C0C0">
	<td>Nº</td>
	<td><b><a href="patrimonio.asp?tipo=cadastro&ordem=cd_patrimonio">PATRIMÔNIO</b></a></td>
	<td><b><a href="patrimonio.asp?tipo=cadastro&ordem=cd_pn">CÓD. PRODUTO</b></a></td>
	<td><b><a href="patrimonio.asp?tipo=cadastro&ordem=nm_equipamento">EQUIPAMENTO</b></a></td>
	<td><b><a href="patrimonio.asp?tipo=cadastro&ordem=nm_marca">MARCA</b></a></td>
	<td><b><a href="patrimonio.asp?tipo=cadastro&ordem=cd_ns">SÉRIE</b></a></td>
	<td><b><a href="patrimonio.asp?tipo=cadastro&ordem=nm_unidade">UNIDADE</b></a></td>
	<td><b><a href="patrimonio.asp?tipo=cadastro&ordem=dt_data">COMPRA</b></a></td>
	<td id="no_print">&nbsp;</td>
</tr>
<tr>
	<td colspan="100%"><img src="../imagens/px.gif" alt="" width="100%" height="1" border="0"></td>
</tr>
<%num_lista = 1
			condicao = " Where dt_descarte is null"
			ordem = " "&ordem&" "
			'ordem = " cd_codigo desc"
			strsql ="up_patrimonio_lista @ordem='"&ordem&"', @condicao = '"&condicao&"'"
		  	Set rs_patrimonio = dbconn.execute(strsql)
			
			Do while not rs_patrimonio.EOF
			cd_pat_codigo = rs_patrimonio("cd_codigo")
			cd_patrimonio = rs_patrimonio("cd_patrimonio")
			cd_equipamento = rs_patrimonio("cd_equipamento")
			nm_equipamento = rs_patrimonio("nm_equipamento")
			cd_marca = rs_patrimonio("cd_marca")
			nm_marca = rs_patrimonio("nm_marca")
			cd_ns = rs_patrimonio("cd_ns")
			cd_pn = rs_patrimonio("cd_pn")
			cd_unidade = rs_patrimonio("cd_unidade")
			nm_sigla = rs_patrimonio("nm_sigla")
			dt_data = rs_patrimonio("dt_data")
			%>
<!--tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio_cadastro.asp?cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"-->
<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');">
	<td align="center" valign="top"><b><%=zero(num_lista)%></b></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"><%=cd_patrimonio%></td>
		<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio_cadastro.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"><%=cd_pn%></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"><%=nm_equipamento%></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"><%=nm_marca%></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"><%=cd_ns%></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"><%=nm_sigla%></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"><%=zero(day(dt_data))%>/<%=zero(month(dt_data))%>/<%=year(dt_data)%></td>	
	<td id="no_print" valign="top"><img src="imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_pat_codigo%>')" type="button" value="A" title="Apagar"></td>
</tr>
<%num_lista = num_lista + 1
rs_patrimonio.movenext
loop
%>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>	
	<td><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="85" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="155" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="65" height="1" border="0"></td>
	<td id="no_print"><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
</tr><tr><td colspan="100%">&nbsp;</td></tr>

</form>
</table>

</td></tr></table>
<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1, cod2, cod3, cod4, cod5, cod6)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('acoes/patrimonio_acao.asp?cd_apaga='+cod1+'&cd_tipo=patrimonio');
	  }
}

</SCRIPT>
</body>
</html>
