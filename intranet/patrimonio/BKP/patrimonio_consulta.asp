<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_pat_codigo = request("cd_pat_codigo")
str_patrimonio = request("no_patrimonio")
nm_tipo = request("nm_tipo")
cd_equipamento =  request("cd_equipamento")
nm_equipamento = request("nm_equipamento")
cd_especialidade = request("cd_especialidade")
	if cd_especialidade = "" Then 
		cd_especialidade = "0" 
	end if
cd_marca =  request("cd_marca")
nm_marca =  request("nm_marca")
cd_ns =  request("cd_ns")
cd_pn =  request("cd_pn")
cd_unidade =  request("cd_unidade")
nm_sigla =  request("nm_sigla")

dt_data =  request("dt_data")
ordem = request("ordem")
	if ordem = "" Then
	ordem = "nm_tipo, no_patrimonio"
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


function foco(){
document.form.no_patrimonio.focus(); }
// -->	
</script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "no_patrimonio"; // nome do primeiro campo do site
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

<body onload="foco()">

<table border="0" class="txt" id="no_print" style="font-size:12px; font-family:arial;" align="center">
<%
sql = "SELECT * FROM "

			if cd_pat_codigo <> "" Then
				condicao = "cd_codigo = 10"''"&cd_pat_codigo&"' "
				ordem = " "&ordem&" "
			
				strsql ="SELECT * FROM View_patrimonio_lista WHERE cd_codigo = "&cd_pat_codigo&""'@ordem=10" '"&ordem&"', @condicao = "&condicao&""
			  	Set rs_patrimonio = dbconn.execute(strsql)
				
				IF not rs_patrimonio.EOF THEN
	 			cd_pat_codigo = rs_patrimonio("cd_codigo")
				no_patrimonio = rs_patrimonio("no_patrimonio")
				cd_equipamento = rs_patrimonio("cd_equipamento")
				nm_equipamento = rs_patrimonio("nm_equipamento")
				cd_marca = rs_patrimonio("cd_marca")
				nm_marca = rs_patrimonio("nm_marca")
				cd_ns = rs_patrimonio("cd_ns")
				cd_pn = rs_patrimonio("cd_pn")
				nm_tipo = rs_patrimonio("nm_tipo")
				cd_especialidade = rs_patrimonio("cd_especialidade")
				cd_unidade = rs_patrimonio("cd_unidade")
				nm_sigla = rs_patrimonio("nm_sigla")
				cd_rack = rs_patrimonio("cd_rack")
				num_hospital =  rs_patrimonio("num_hospital")
				cd_rack = rs_patrimonio("cd_rack")
				num_hospital = rs_patrimonio("num_hospital")
				cd_categoria = rs_patrimonio("cd_categoria")
				dt_data = rs_patrimonio("dt_data")
				
				cd_preventiva = rs_patrimonio("cd_preventiva")
				dt_periodicidade_prev = rs_patrimonio("dt_periodicidade_prev")
				
				cd_calibracao = rs_patrimonio("cd_calibracao")
				dt_periodicidade_cal = rs_patrimonio("dt_periodicidade_cal")
				
				cd_seg_elet = rs_patrimonio("cd_seg_elet")
				dt_periodicidade_elet = rs_patrimonio("dt_periodicidade_elet")
				
				nao_constar = rs_patrimonio("nao_constar")
				
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				'ordem = " cd_codigo desc"
					'strsql_espec ="SELECT * FROM TBL_ESPEC Where cd_codigo='"&cd_especialidade&"'"
					strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"'"
				  	Set rs_espec = dbconn.execute(strsql_espec)
						if not rs_espec.EOF then
							nm_especialidade = rs_espec("nm_especialidade")
						end if
				'rs_espec.movenext
				End if
			End if

%>
<%if session("cd_codigo") = 46 then%>
	<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">patrimonio_consulta.asp</span></td></tr>
<%end if%>
<tr>
	<td colspan="3">&nbsp;<b>Inventário - <font color="red">Consulta equipamentos/instrumentos</font></b></td>
	<td align="right">&nbsp;</td>
</tr>
<tr bgcolor=#cococo>
	<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
</tr>
<!--tr><td align=center colspan="100%">&nbsp;</td></tr-->
<tr>
    <td>&nbsp;<b> Nº PATRIMÔNIO:</b></td>
    <td><%=no_patrimonio%></td>
    <td>&nbsp;<b>TIPO PATRIMÔNIO:</b></td>
    <td colspan="3">
		<%if nm_tipo = "E" Then %>
			Equipamento
		<%elseif nm_tipo = "I" Then%>
			Instrumento
		<%elseif nm_tipo = "O" Then%>
			Ótica
		<%end if%>		
	</td>
</tr>
<tr><td align=center colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<tr>
    <td>&nbsp; <b>ITEM:</b> </td>
		<%	selecao = " top 100 percent * "
			tabela = " TBL_equipamento"
			condicoes = " where cd_codigo="&cd_equipamento&" "
			criterios = " Order by nm_equipamento"
			
			strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
		  	Set rs_equip = dbconn.execute(strsql)%>
    <td><%=rs_equip("nm_equipamento_novo")%></td>
	<td>&nbsp; <b>CATEGORIA:</b> </td>
	<td>						<%if cd_categoria <> "" Then
									strsql = "SELECT * FROM TBL_patrimonio_categoria where cd_codigo = "&cd_categoria&" ORDER BY nm_categoria"
									Set rs_cat = dbconn.execute(strsql)
									while not rs_cat.EOF
										categoria = rs_cat("cd_codigo")
										nm_categoria = rs_cat("nm_categoria")
										%>
									<%=nm_categoria%>
									<%rs_cat.movenext
									wend
								end if%>
							
	</td>								
</tr>
<tr>
    <td>&nbsp; <b>MODELO:</b></td>
    <td><%=cd_pn%></td>
    <td>&nbsp; <b>NUM. SÉRIE:</b></td>
    <td><%=cd_ns%></td>
</tr>
<tr>
	<td>&nbsp; <b>MARCA:</b></td>
	<%
	selecao = " top 100 percent * "
	tabela = " TBL_marca"
	condicoes = " where cd_codigo = "&cd_marca&" "
	criterios = " Order by nm_marca"
	
	strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
  	Set rs_marca = dbconn.execute(strsql)%>
	<td>
	<%Do while Not rs_marca.EOF%>
	<%=rs_marca("nm_marca")%>
	<%rs_marca.movenext
	marca_check = ""
  	Loop%>
	</td>
    <td>&nbsp; <b>DATA DA AQUISIÇÃO:</b></td>
					<%if cd_pat_codigo <> "" Then%>
						<%dt_dia = zero(day(dt_data))%>
						<%dt_mes = zero(month(dt_data))%>
						<%dt_ano = year(dt_data)%>
					<%end if%>
    <td><%=dt_dia%>/<%=dt_mes%>/<%=dt_ano%></td>
</tr>
<tr><td align=center colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<tr>
	<td>&nbsp; <b>Rack:</b></td>
	<%strsql ="SELECT * FROM TBL_rack where A056_codrac = "&cd_rack&" order by A056_desrac"
		Set rs_rack = dbconn.execute(strsql)%>
	<td><%while not rs_rack.EOF%>
				<%=rs_rack("A056_desrac")%>
			<%rs_rack.movenext
		wend%>
	</td>
	<td>&nbsp; <b>Especialidade:</b> </td>
		<%	selecao = " top 100 percent * "
			'tabela = " TBL_Espec"
			tabela = " TBL_Especialidades"
			condicoes = " where cd_codigo = "&cd_especialidade&" "
			criterios = " Order by nm_especialidade"
			
			strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
		  	Set rs_espec = dbconn.execute(strsql)%>
	<td>				<%Do while Not rs_espec.EOF
							if int(cd_especialidade) = rs_espec("cd_codigo") then%><%espec_check = "selected"%><%end if%>							
								<%=rs_espec("nm_especialidade")%>
								<%rs_espec.movenext
								eqp_check = ""
								espec_check = ""
						Loop%>
	</td>
</tr>
<tr><td align=center colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>

<tr>
    <td>&nbsp; <b>Unidade:</b></td>
		<%strsql ="SELECT * FROM TBL_unidades WHERE cd_codigo = "&cd_unidade&""
		Set rs_uni = dbconn.execute(strsql)%>
    <td>	<%Do While Not rs_uni.eof%>
				<%=rs_uni("nm_unidade")%>
				<%rs_uni.movenext
			loop%>					
	</td>

	<td>&nbsp; <b>Nº Hospital (CC):</b></td>
	<td><%=num_hospital%></td>
</tr>
<tr><td align=center colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<tr>
	<td colspan="100%" align="left" style="border:1px solid black;"><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"> 
		<b>MANUTENÇÕES</b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
		<b>PERIODICIDADE</b> &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
		<b>&nbsp;</b></td>
	<!--td style="border:1px solid black;">Periodicidade</td-->
</tr>
<tr>
	<td>&nbsp; <b>Preventiva:</b></td>
	<td colspan="3"><%if cd_preventiva = 1 then%>[X]<%Else%>[ ]<%end if%> <b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
	<b></b>  &nbsp; &nbsp; &nbsp; &nbsp; <%=dt_periodicidade_prev%> <b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp;
	</td>
</tr>
<tr>
	<td>&nbsp; <b>Calibração:</b></td>
	<td colspan="3"><%if cd_calibracao = 1 then%>[X]<%Else%>[ ]<%end if%> <b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
	<b></b>  &nbsp; &nbsp; &nbsp; &nbsp; <%=dt_periodicidade_cal%> <b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	</td>
</tr>
<tr>
	<td>&nbsp; <b>Segurança Elet.:</b></td>
	<td colspan="3"><%if cd_seg_elet = 1 then%>[X]<%Else%>[ ]<%end if%> <b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
	<b></b>  &nbsp; &nbsp; &nbsp; &nbsp; <%=dt_periodicidade_elet%> <b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp;
	</td>

</tr>
<%if cd_pat_codigo <> "" Then%>
<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<tr><td colspan="100%">
	<%id_origem = cd_pat_codigo'*** Este é o Código do objeto ***
	cd_origem = 5
	pag_retorno = ".."&mem_posicao
	variaveis = "../.."&mem_posicao&"?tipo=cadastro" '*** Pagina de retorno, com variaveis%>
	<!--include file="../ocorrencia/ocorrencia_formulario.asp"-->
	<!--#include file="../ocorrencia/ocorrencia_visualizacao.asp"-->
</td></tr>
<%end if%>
</form>
<tr>
    <td colspan="4">&nbsp;</td>
</tr>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
</tr>
</table>
<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1)
	{
	  if (confirm ("Tem certeza que deseja excluir?"))
		  {
			document.location.href('acoes/patrimonio_acao.asp?cd_apaga='+cod1+'&cd_tipo=patrimonio&acao=apaga');
		  }
	}
</SCRIPT>
</body>
</html>
