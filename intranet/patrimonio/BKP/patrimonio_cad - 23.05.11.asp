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
<script language="javascript">
/*var a= "18";
var b= "12";
var c= eval(a/b);
var d= eval(a%b);
document.write("<b>Divisão:&nbsp;</b>"+c+"<br><b>Resto:&nbsp;</b>"+d);*/

function determina_manutencao_prev(){
	a = document.form.acao.value;
	b = document.form.cd_ano.value;
	
	if (a == "inserir"){	//verifica a ação
		if (cd_preventiva = 1){
			document.form.dt_plan_ano_prev.value = b
		}		
	}
	//document.write(a);
}
</script>

<body onload="foco()">
<table border="0" class="txt" id="no_print" style="font-size:12px; font-family:arial;" align="center">
<tr><td>:: <a href="patrimonio_cad.asp">Novo</a> ::</td></tr>
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
<form action="acoes/patrimonio_acao.asp" method="post" name="form" onsubmit="return validar_patrimonio(document.form);" onSubmit="return checa(this);">
<input type="hidden" name="janela" value="1">
<input type="hidden" name="cd_tipo" value="patrimonio">
<input type="hidden" name="cd_pat_codigo" value="<%=cd_pat_codigo%>">
<input type="hidden" name="acao" value="<%=acao%>">
<tr>
	<td colspan="3">&nbsp;<b>Inventário - <font color="red">Cadastro de equipamentos/instrumentos</font></b></td>
	<td align="right"><img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="javascript:JsDelete('<%=cd_pat_codigo%>')">&nbsp;</td>
</tr>
<tr bgcolor=#cococo>
	<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
</tr>
<!--tr><td align=center colspan="100%">&nbsp;</td></tr-->
<tr>
    <td>&nbsp;<b> Nº Patrimônio:</b></td>
    <td><input type="text" name="no_patrimonio" value="<%=no_patrimonio%>" class="inputs" onFocus="nextfield ='nm_tipo';"></td>
    <td>&nbsp;Tipo patrimoinio:</td>
    <td colspan="3">
		<%if nm_tipo = "E" Then 
			tipo_E = "selected"
		elseif nm_tipo = "I" Then
			tipo_I = "selected"
		elseif nm_tipo = "O" Then
			tipo_O = "selected"
		end if%>
		<select name="nm_tipo" onFocus="nextfield ='cd_equipamento';"  class="inputs">
			<option value=""></option>
			<option value="E" <%=tipo_E%>>Equipamento</option>
			<option value="I" <%=tipo_I%>>Instrumento</option>
			<option value="O" <%=tipo_O%>>Ótica</option>
		</select>
	</td>
</tr>
<tr><td align=center colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<tr>
    <td>&nbsp; Item: </td>
		<%	selecao = " top 100 percent * "
			tabela = " TBL_equipamento"
			condicoes = ""
			criterios = " Order by nm_equipamento"
			
			strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
		  	Set rs_equip = dbconn.execute(strsql)%>
    <td colspan="3"><select name="cd_equipamento" class="inputs" onFocus="nextfield ='cd_categoria';" class="inputs">
					<option value=""></option>
					<%Do while Not rs_equip.EOF
					if int(cd_equipamento) = rs_equip("cd_codigo") then%><%eqp_check = "selected"%><%end if%>
					<option value="<%=rs_equip("cd_codigo")%>" <%=eqp_check%>><%=rs_equip("nm_equipamento_novo")%></option>
					<%rs_equip.movenext
					eqp_check = ""
					Loop%>
					</select>
					<a href="javascript:;" onclick="window.open('janelas_cadastros/equipamento_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a>
	&nbsp; &nbsp; Categoria: <select name="cd_categoria" onFocus="nextfield ='cd_pn';" class="inputs">
								<option value="">&nbsp;</option>
							<%strsql = "SELECT * FROM TBL_patrimonio_categoria ORDER BY nm_categoria"
								Set rs_cat = dbconn.execute(strsql)
								while not rs_cat.EOF
									categoria = rs_cat("cd_codigo")
									nm_categoria = rs_cat("nm_categoria")
									
									if int(cd_categoria) = categoria then ck_cat = "SELECTED" else ck_cat = "" end if%>
								<option value="<%=categoria%>" <%=ck_cat%>><%=nm_categoria%></option>
								<%
								rs_cat.movenext
								wend%>
							</select>
	</td>								
</tr>
<tr>
    <td>&nbsp; Modelo:</td>
    <td><input type="text" name="cd_pn" value="<%=cd_pn%>" class="inputs" onFocus="nextfield ='cd_ns';"></td>
    <td>&nbsp; Num. Série:</td>
    <td><input type="text" name="cd_ns" size="30" value="<%=cd_ns%>" class="inputs" onFocus="nextfield ='cd_marca';">
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
	<td><select name="cd_marca" size="1" class="inputs" onFocus="nextfield ='cd_dia';">
	<option value=""></option>
	<%Do while Not rs_marca.EOF
	if int(cd_marca) = rs_marca("cd_codigo") then%><%marca_check = "selected"%><%end if%>
	<option value="<%=rs_marca("cd_codigo")%>" <%=marca_check%>><%=rs_marca("nm_marca")%></option>	
	<%rs_marca.movenext
	marca_check = ""
  	Loop%>
	</td>
    <td>&nbsp; Data da aquisição:</td>
					<%if cd_pat_codigo <> "" Then%>
						<%dt_dia = zero(day(dt_data))%>
						<%dt_mes = zero(month(dt_data))%>
						<%dt_ano = year(dt_data)%>
					<%end if%>
    <td><input type="text" name="cd_dia" size="2" maxlength="2" value="<%=dt_dia%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" onFocus="nextfield ='cd_mes';">/
					<input type="text" name="cd_mes" size="2" maxlength="2" value="<%=dt_mes%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" onFocus="nextfield ='cd_ano';">/
					<input type="text" name="cd_ano" size="4" maxlength="4" value="<%=dt_ano%>" class="inputs" onFocus="nextfield ='cd_rack';"></td>
</tr>
<tr><td align=center colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<tr>
	<td>&nbsp; Rack:</td>
	<%strsql ="SELECT * FROM TBL_rack order by A056_desrac"
		Set rs_rack = dbconn.execute(strsql)%>
	<td>
		<select name="cd_rack" class="inputs" onFocus="nextfield ='cd_especialidade';">
			<option value="" SELECTED>&nbsp;</option>
			<%while not rs_rack.EOF%>
				<option value="<%=rs_rack("A056_codrac")%>" <%if int(cd_rack) = rs_rack("A056_codrac") then response.write("SELECTED") End if%>><%=rs_rack("A056_desrac")%></option>
			<%rs_rack.movenext
			wend%>
		</select>
	
	</td>
	<td>&nbsp; Especialidade: </td>
		<%	selecao = " top 100 percent * "
			'tabela = " TBL_Espec"
			tabela = " TBL_Especialidades"
			condicoes = ""
			criterios = " Order by nm_especialidade"
			
			strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
		  	Set rs_espec = dbconn.execute(strsql)%>
	<td><select name="cd_especialidade" class="inputs" onFocus="nextfield ='cd_unidade';">
						<option value=""></option>
						<%Do while Not rs_espec.EOF
							if int(cd_especialidade) = rs_espec("cd_codigo") then%><%espec_check = "selected"%><%end if%>							
								<option value="<%=rs_espec("cd_codigo")%>" <%=espec_check%>><%=rs_espec("nm_especialidade")%></option>
								<%rs_espec.movenext
								eqp_check = ""
								espec_check = ""
						Loop%>
					</select>
					<a href="javascript:;" onclick="window.open('janelas_cadastros/equipamento_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a>
	</td>
</tr>
<tr><td align=center colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>

<tr>
    <td>&nbsp; Unidade:</td>
		<%strsql ="up_ADM_unidades_lista"
		Set rs_uni = dbconn.execute(strsql)%>
    <td><select name="cd_unidade" class="inputs" onFocus="nextfield ='num_hospital';">
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

	<td>&nbsp; Nº Hospital (CC):</td>
	<td><input type="text" name="num_hospital" size="20" value="<%=num_hospital%>" maxlength="20" class="inputs" onFocus="nextfield ='cd_preventiva';"><!--onFocus="nextfield ='dt_periodicidade_prev';"--></td>
</tr>
<tr><td align=center colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<tr>
	<td colspan="100%" align="left" style="border:1px solid black;">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		<b>MANUTENÇÕES</b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		<b>PERIODICIDADE</b> &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		<b>PLANEJAMENTO</b></td>
	<!--td style="border:1px solid black;">Periodicidade</td-->
</tr>
<tr>
	<td>&nbsp; <b>Preventiva:</b></td>
	<td colspan="3"><input type="checkbox" name="cd_preventiva" value="1" class="inputs" <%if cd_preventiva = 1 then response.write("CHECKED") end if%> onFocus="nextfield ='dt_periodicidade_prev';"><b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<b></b> <input type="text" name="dt_periodicidade_prev" size="3" maxlength="2" value="<%=dt_periodicidade_prev%>" class="inputs" onFocus="nextfield ='dt_plan_mes_prev';"  onKeyup="determina_manutencao_prev();"><b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<!--input type="text" name="dt_plan_mes_prev" size="2" maxlength="2" class="inputs" onFocus="nextfield ='dt_plan_ano_prev';"-->
	<select name="dt_plan_mes_prev" onFocus="nextfield ='dt_plan_ano_prev';">
		<option value=""></option>
	<%for i = 1 to 12%>	
		<option value="<%=i%>"><%=ucase(left(mes_selecionado(i),3))%></option>
	<%next%>
	</select>/<input type="text" name="dt_plan_ano_prev" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='cd_calibracao';">(mes/ano)
	(<b>Próx. Preventiva:</b>
	<%if cd_patrimonio <> "" Then
		strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = 1 order by dt_plan_prev desc"
			Set rs_prev = dbconn.execute(strsql)
		while not rs_prev.EOF
			dt_plan_prev = rs_prev("dt_plan_prev")%>
		<b style="color:#006600;"><%=Ucase(left(mes_selecionado(month(dt_plan_prev)),3))%>/<%=year(dt_plan_prev)%></b>
		<%rs_prev.movenext
		wend%>)
	<%end if%></td>
</tr>
<tr>
	<td>&nbsp; <b>Calibração:</b></td>
	<td colspan="3"><input type="checkbox" name="cd_calibracao" value="1" class="inputs" <%if cd_calibracao = 1 then response.write("CHECKED") end if%> onFocus="nextfield ='dt_periodicidade_cal';"><b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<b></b> <input type="text" name="dt_periodicidade_cal" size="3" maxlength="2" value="<%=dt_periodicidade_cal%>" class="inputs" onFocus="nextfield ='dt_plan_mes_cal';" onFocus="nextfield ='dt_plan_mes_cal';"><b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<!--input type="text" name="dt_plan_mes_cal" size="2" maxlength="2" class="inputs" onFocus="nextfield ='dt_plan_ano_cal';"-->
	<select name="dt_plan_mes_cal" onFocus="nextfield ='dt_plan_ano_cal';">
		<option value=""></option>
	<%for i = 1 to 12%>	
		<option value="<%=i%>"><%=ucase(left(mes_selecionado(i),3))%></option>
	<%next%>
	</select>/<input type="text" name="dt_plan_ano_cal" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='cd_seg_elet';">(mes/ano)
	(<b>Próx. Calibração:</b>
	<%if cd_patrimonio <> "" Then
		strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = 2 order by dt_plan_prev desc"
			Set rs_cal = dbconn.execute(strsql)
		while not rs_cal.EOF
			dt_plan_cal = rs_cal("dt_plan_prev")%>
		<b style="color:#00cc00;"><%=Ucase(left(mes_selecionado(month(dt_plan_cal)),3))%>/<%=year(dt_plan_cal)%></b>
		<%rs_cal.movenext
		wend%>)
		<%end if%></td>
</tr>
<tr>
	<td>&nbsp; <b>Segurança Elet.:</b></td>
	<td colspan="3"><input type="checkbox" name="cd_seg_elet" value="1" class="inputs" <%if cd_seg_elet = 1 then response.write("CHECKED") end if%> onFocus="nextfield ='dt_periodicidade_elet';"><b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<b></b> <input type="text" name="dt_periodicidade_elet" size="3" maxlength="2" value="<%=dt_periodicidade_elet%>" class="inputs" onFocus="nextfield ='dt_plan_mes_elet';" onFocus="nextfield ='dt_plan_mes_elet';"><b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<!--input type="text" name="dt_plan_mes_elet" size="2" maxlength="2" class="inputs" onFocus="nextfield ='dt_plan_ano_cal';"-->
	<select name="dt_plan_mes_elet" onFocus="nextfield ='dt_plan_ano_elet';">
			<option value=""></option>
	<%for i = 1 to 12%>	
		<option value="<%=i%>"><%=ucase(left(mes_selecionado(i),3))%></option>
	<%next%>
	</select>/<input type="text" name="dt_plan_ano_elet" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='envia';">(mes/ano)
	(<b>Próx Seg. Elétrica:</b>
	<%if cd_patrimonio <> "" Then
		strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = 3 order by dt_plan_prev desc"
			Set rs_elet = dbconn.execute(strsql)
		while not rs_elet.EOF
			dt_plan_elet = rs_elet("dt_plan_prev")%>
		<b style="color:#99cc00;"><%=Ucase(left(mes_selecionado(month(dt_plan_elet)),3))%>/<%=year(dt_plan_elet)%></b>
		<%rs_elet.movenext
		wend%>)
	<%end if%></td>

</tr>
<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<tr>
    <td>&nbsp;</td>
    <td><input type="submit" name="envia" value=" OK "> <br><font color="red"><%=mensagem%></font></td>
	<td align="right" valign="middle"><input type="checkbox" name="nao_constar" value="1"  onFocus="nextfield ='envia';" <%if int(nao_constar) = 1 then%>checked<%end if%>></td>
	<td align="left" valign="middle"><i style="font-size:8px;">Marque este campo caso desejar que este item não<br>conste no inventário de equipamento do hospital</i></td>
</tr>
<%if cd_pat_codigo <> "" Then%>
<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<tr><td colspan="100%">
	<%id_origem = cd_pat_codigo'*** Este é o Código do objeto ***
	cd_origem = 5
	pag_retorno = ".."&mem_posicao
	variaveis = "../.."&mem_posicao&"?tipo=cadastro" '*** Pagina de retorno, com variaveis%>
	<!--#include file="../ocorrencia/ocorrencia_formulario.asp"-->
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
