<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_pat_codigo = request("cd_pat_codigo")
		if cd_pat_codigo <> "" Then
			
			condicao = " where cd_codigo="&cd_pat_codigo&" "
			ordem = " cd_codigo "
			
			strsql ="up_patrimonio_lista @condicao='"&condicao&"', @ordem='"&ordem&"'"
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
			dt_descarte = rs_patrimonio("dt_descarte")
			nm_motivo = rs_patrimonio("nm_motivo")
				
				if dt_descarte <> "" Then
				num_color = "red"
				end if
				
			num_lista = num_lista + 1
			num_color = ""
			rs_patrimonio.movenext
			loop
			
		End if
'cd_patrimonio = request("cd_patrimonio")
'cd_equipamento =  request("cd_equipamento")
'nm_equipamento = request("nm_equipamento")
'cd_marca =  request("cd_marca")
'nm_marca =  request("nm_marca")
'cd_ns =  request("cd_ns")
'cd_pn =  request("cd_pn")
'cd_unidade =  request("cd_unidade")
'nm_sigla =  request("nm_sigla")
'dt_data =  request("dt_data")
'dt_descarte = request("dt_descarte")
'nm_motivo = request("nm_motivo")

acao = "descartar"'request("acao")
	'if acao = "" Then
	'acao = "inserir"
	'Else
	'acao = "editar"
	'end if

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
<body><br>
<table align="center">
<tr><td>&nbsp;&nbsp;&nbsp;</td>
<td align="center">
<!--table cellspacing="0" cellpadding="0" border="1" rules="groups" class="txt" id="no_print"-->
<table class="txt" id="no_print">
<form action="../acoes/patrimonio_acao.asp" method="post">
<input type="hidden" name="janela" value="1">
<input type="hidden" name="cd_tipo" value="patrimonio">
<input type="hidden" name="cd_pat_codigo" value="<%=cd_pat_codigo%>">
<input type="hidden" name="acao" value="<%=acao%>">
<tr>
	<td colspan="100%">&nbsp;<b>Inventário - <font color="red">Descarte de equipamentos/instrumentos</font></b></td>
</tr>

<tr bgcolor=#cococo>
	<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
</tr>
<%if cd_pat_codigo <> "" Then %>
<tr><td align=center colspan="100%">&nbsp;</td></tr>
<tr>
    <td>&nbsp;<b>Nº Patrimônio:</b></td>
    <td colspan="3"><%=cd_patrimonio%></td>
</tr>
<tr>
    <td>&nbsp;<b>Cód. Fabricante:</b></td>
    <td colspan="3"><%=cd_pn%></td>
</tr>
<tr>
    <td>&nbsp;<b>Equip/Instr:</b></td>
		<%	selecao = " top 100 percent * "
			tabela = " TBL_equipamento"
			condicoes = ""
			criterios = " Order by nm_equipamento"
			
			strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
		  	Set rs_equip = dbconn.execute(strsql)%>
    <td colspan="3"><%Do while Not rs_equip.EOF
					if int(cd_equipamento) = rs_equip("cd_codigo") then%><%=rs_equip("nm_equipamento")%><%end if%>
					<%rs_equip.movenext
					eqp_check = ""
					Loop%>
					</td>
								
</tr>
<tr>
	<td>&nbsp;<b>Marca:</b></td>
	<%
	selecao = " top 100 percent * "
	tabela = " TBL_marca"
	condicoes = ""
	criterios = " Order by nm_marca"
	
	strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
  	Set rs_marca = dbconn.execute(strsql)%>
	<td><%Do while Not rs_marca.EOF
	if int(cd_marca) = rs_marca("cd_codigo") then%><%response.write(rs_marca("nm_marca"))%>
	<%end if
	rs_marca.movenext
	marca_check = ""
  	Loop%>
	</td>
</tr>
<tr>
    <td>&nbsp;<b>Num. Série:</b></td>
    <td colspan="3"><%=cd_ns%></td>
</tr>
<tr>
    <td>&nbsp;<b>Unidade:</b></td>
		<%strsql ="up_ADM_unidades_lista"
		Set rs_uni = dbconn.execute(strsql)%>
    <td colspan="3"><%Do While Not rs_uni.eof
					if int(cd_unidade) = rs_uni("cd_codigo") then%><%response.write(rs_uni("nm_unidade"))%><%end if%>
					
					<%rs_uni.movenext
					unidade_check = ""
					loop%>
					</select>
	</td>
</tr>


<tr>
    <td>&nbsp;<b>Data de compra:</b></td>
					<%if cd_pat_codigo <> "" Then%>
						<%dt_dia = day(dt_data)%>
						<%dt_mes = month(dt_data)%>
						<%dt_ano = year(dt_data)%>
					<%end if%>
    <td colspan="3"><%=dt_dia%>/<%=dt_mes%>/<%=dt_ano%></td>
</tr>
<tr>
    <td>&nbsp;<b>Data de descarte:</b></td>
					<%if dt_descarte <> "" Then%>
						<%dia_descarte = day(dt_descarte)%>
						<%mes_descarte = month(dt_descarte)%>
						<%ano_descarte = year(dt_descarte)%>
					<%end if%>
    <td colspan="3"><input type="text" name="cd_dia" size="2" maxlength="2" value="<%=dia_descarte%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
					<input type="text" name="cd_mes" size="2" maxlength="2" value="<%=mes_descarte%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
					<input type="text" name="cd_ano" size="4" maxlength="4" value="<%=ano_descarte%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
</tr>
<tr>
    <td valign="center">&nbsp;<b>Motivo do descarte:</b></td>
    <td colspan="3"><textarea cols="20" rows="4" name="nm_motivo"><%=nm_motivo%></textarea></td>
</tr>
<tr>
    <td colspan="4">&nbsp;</td>
</tr>
<tr>
    <td>&nbsp;</td>
    <td colspan="3"><input type="submit" name="ok" value="Descarta"> &nbsp; <a href="patrimonio_cadastro.asp">Lista/Novo</a><br>
	<font color="red"><%=mensagem%></font></td>
</tr>
<tr>
    <td colspan="4">&nbsp;</td>
</tr>
<%end if%>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="330" height="1" border="0"></td>
</tr>
</table>
<br><!-- LISTAGEM DO PATRIMONIO -->
<table class="txt" border="0" width="600">
<tr bgcolor="#C0C0C0">
	<td>Nº</td>
	<td><b>PATRIMÔNIO</b></td>
	<!--td><b>CÓD. PRODUTO</b></td-->
	<td><b>EQUIPAMENTO</b></td>
	<td><b>MARCA</b></td>
	<td><b>SÉRIE</b></td>
	<!--td><b>UNIDADE</b></td-->
	<td><b>COMPRA</b></td>
	<td><b>Descarte</b></td>
	<td><b>Motivo</b></td>
	<td id="no_print">&nbsp;</td>
</tr>
<tr>
	<td colspan="100%"><img src="../imagens/px.gif" alt="" width="100%" height="1" border="0"></td>
</tr>
<%num_lista = 1
			condicao = ""
			ordem = " nm_equipamento DESC "
			strsql ="up_patrimonio_lista @condicao='"&condicao&"', @ordem='"&ordem&"'"
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
			dt_descarte = rs_patrimonio("dt_descarte")
			nm_motivo = rs_patrimonio("nm_motivo")
			
			if dt_descarte <> "" Then
			num_color = "red"
			separador = "/"
			Else
			num_color = "black"
			end if
			%>
<!--tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio_cadastro.asp?cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"-->
<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');">
	<td align="center" valign="top"><font color="<%=num_color%>"><b><%=zero(num_lista)%></b></font></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=descarte&cd_pat_codigo=<%=cd_pat_codigo%>&acao=editar');" onmouseout="mOut(this,'');">
	<%=cd_patrimonio%></td>
	<!--td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=descarte&cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');">
	<%=cd_pn%></td-->
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=descarte&cd_pat_codigo=<%=cd_pat_codigo%>&acao=editar');" onmouseout="mOut(this,'');">
	<%=nm_equipamento%></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=descarte&cd_pat_codigo=<%=cd_pat_codigo%>&acao=editar');" onmouseout="mOut(this,'');">
	<%=nm_marca%></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=descarte&cd_pat_codigo=<%=cd_pat_codigo%>&acao=editar');" onmouseout="mOut(this,'');">
	<%=cd_ns%></td>
	<!--td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=descarte&cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&cd_patrimonio=<%=cd_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');">
	<%=nm_sigla%></td-->
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=descarte&cd_pat_codigo=<%=cd_pat_codigo%>&acao=editar');" onmouseout="mOut(this,'');">
	<%=zero(day(dt_data))%>/<%=zero(month(dt_data))%>/<%=year(dt_data)%></td>	
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=descarte&cd_pat_codigo=<%=cd_pat_codigo%>&acao=editar');" onmouseout="mOut(this,'');">
	<%=zero(day(dt_descarte))%><%=separador%><%=zero(month(dt_descarte))%><%=separador%><%=zero(year(dt_descarte))%></td>
	<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio.asp?tipo=descarte&cd_pat_codigo=<%=cd_pat_codigo%>&acao=editar');" onmouseout="mOut(this,'');">
	<%=left(nm_motivo,15)%>...</td>	
	<td id="no_print" valign="top"><img src="imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_pat_codigo%>')" type="button" value="A" title="Apagar">
	</td>
</tr>
<%num_lista = num_lista + 1
num_color = ""
rs_patrimonio.movenext
loop
%>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>	
	<td><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="155" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="85" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="65" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="65" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="95" height="1" border="0"></td>
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
