<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>

<%no_patrimonio = request("no_patrimonio")
nm_tipo = request("nm_tipo")
cd_pat_codigo = request("cd_pat_codigo")
existe = request("existe")
cd_equipamento =  request("cd_equipamento")
cd_categoria = request("cd_categoria")
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
cd_rack = request("cd_rack")
num_hospital = request("num_hospital")

dt_data =  request("dt_data")
	if dt_data <> "" then
		dt_dia = day(dt_data)
		dt_mes = month(dt_data)
		dt_ano = year(dt_data)
	else
		dt_dia = ""
		dt_mes = ""
		dt_ano = ""
		dt_data = dt_mes&"/"&dt_dia&"/"&dt_ano
		'dt_dia = day(now)
		'dt_mes = month(now)
		'dt_ano = year(now)
		'dt_data = month(now)&"/"&day(now)&"/"&year(now)
	end if
	
cd_preventiva = request("cd_preventiva")
dt_periodicidade_prev =request("dt_periodicidade_prev")

cd_calibracao = request("cd_calibracao")
dt_periodicidade_cal = request("dt_periodicidade_cal")

cd_seg_elet = request("cd_seg_elet")
dt_periodicidade_elet = request("dt_periodicidade_elet")

nao_constar = request("nao_constar")

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

jan = request("jan")
	if jan = 1 then
		ini_caminho = "../"
	end if
caminho = request("caminho")
	
	'if caminho = "" Then
	if jan = 1 Then
		str_novo = "/patrimonio/patrimonio_cad.asp?caminho=patrimonio"
	else
		str_novo = 	"/patrimonio.asp?tipo=cadastro&caminho=patrimonio&hahaha"
	end if

mensagem = request("mensagem")
cd_usuario = session("cd_codigo")
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
<body>
<table border="0" class="txt" id="no_print" style="font-size:12px; font-family:arial;" align="center">
<tr>
	<td>:: <a href="<%=str_novo%>&jan=<%=jan%>">Novo</a> :: <%if jan<> 1 Then%><a href="patrimonio.asp?tipo=busca&cat_patr=BTx05_kt">Busca</a><%end if%></td>
	<%if existe = 1 then%>
		<td align="center" style="background-color:#ff0033; color:#FFFFFF;" colspan="2">AVISO: ESTE NÚMERO DE PATRIMÔNIO JÁ EXISTE!</td>
	<%end if%>
</tr>
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
				nm_equipamento = rs_patrimonio("nm_equipamento_novo")
				cd_marca = rs_patrimonio("cd_marca")
				nm_marca = rs_patrimonio("nm_marca")
				cd_ns = rs_patrimonio("cd_ns")
				cd_pn = rs_patrimonio("cd_pn")
				nm_tipo = rs_patrimonio("nm_tipo")
				cd_especialidade = rs_patrimonio("cd_especialidade")
				cd_unidade = rs_patrimonio("cd_unidade")
				nm_unidade = rs_patrimonio("nm_unidade")
				nm_sigla = rs_patrimonio("nm_sigla")
				
				cd_rack = rs_patrimonio("cd_rack")
				nm_rack = rs_patrimonio("nm_rack")
				
				cd_unidade_comodato = rs_patrimonio("cd_unidade_comodato")
				cd_rack_comodato = rs_patrimonio("cd_rack_comodato")
				nm_unidade_comodato = rs_patrimonio("nm_unidade_comodato")
				nm_rack_comodato = rs_patrimonio("nm_rack_comodato")
				
				num_hospital =  rs_patrimonio("num_hospital")
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
				cd_status = rs_patrimonio("cd_status")
					if cd_status = 2 then
						cor_tabela = "#ffffcc"
						nm_status = "<span style='color:red;'> ( &nbsp; (&nbsp; ( D E S C A R T A D O ) &nbsp;) &nbsp; ) </span>"
					else
						cor_tabela = "#f0f0f0"
						nm_status = ""
					end if
				
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
<form action="<%=ini_caminho%><%=caminho%>/acoes/patrimonio_acao.asp" method="post" name="form" onsubmit="return validar_patrimonio(document.form);" onSubmit="return checa(this);">
<input type="hidden" name="jan" value="<%=jan%>">
<input type="hidden" name="cd_tipo" value="patrimonio">
<input type="hidden" name="cd_pat_codigo" value="<%=cd_pat_codigo%>">
<input type="hidden" name="acao" value="<%=acao%>">
<input type="hidden" name="caminho" value="<%=caminho%>">


<%if session("cd_codigo") = 46 then%>
	<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">patrimonio_cad.asp</span></td></tr>
<%end if%>
<tr style="background-color:silver;">
	<td  align="center" colspan="4>&nbsp;<b>INVENTÁRIO - <span style="color:white;">CADASTRO DE PATRIMÔNIO</span></b></td>	
</tr>
<tr bgcolor="#cococo" style="background-color:<%=cor_tabela%>;">
	<td align="center" colspan="4"><img src="imagens/px.gif"  height=1></td>
</tr>
<!--tr><td align=center colspan="100%">&nbsp;</td></tr-->
<tr style="background-color:<%=cor_tabela%>;">
    <td>&nbsp;<b> Nº Patrimônio:</b></td>
    <td><input type="text" name="no_patrimonio" value="<%=no_patrimonio%>" class="inputs" onFocus="nextfield ='nm_tipo';"></td>
    <td>&nbsp;Tipo patrimônio:</td>
    <td>
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
<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4">&nbsp;<b><%=nm_status%></b></td></tr>
<tr style="background-color:<%=cor_tabela%>;">
    <td>&nbsp; Item: </td>
		<%	selecao = " top 100 percent * "
			tabela = " TBL_equipamento"
			condicoes = " where nm_equipamento_novo <> Null and cd_status = 1 "
			criterios = " Order by nm_equipamento_novo"
			
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
									
									%>
								<option value="<%=categoria%>" <%if int(cd_categoria) = categoria then response.write("SELECTED") end if%>><%=nm_categoria%></option>
								<%
								rs_cat.movenext
								wend%>
							</select>
	</td>								
</tr>
<tr style="background-color:<%=cor_tabela%>;">
    <td>&nbsp; Modelo:</td>
    <td><input type="text" name="cd_pn" value="<%=cd_pn%>" class="inputs" onFocus="nextfield ='cd_ns';"></td>
    <td>&nbsp; Num. Série:</td>
    <td><input type="text" name="cd_ns" size="30" value="<%=cd_ns%>" class="inputs" onFocus="nextfield ='cd_marca';">
	</td>
</tr>
<tr style="background-color:<%=cor_tabela%>;">
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
						<%'if 
						dt_dia = zero(day(dt_data))%>
						<%dt_mes = zero(month(dt_data))%>
						<%dt_ano = year(dt_data)%>
					<%end if%>
    <td><input type="text" name="cd_dia" size="2" maxlength="2" value="<%=dt_dia%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
					<input type="text" name="cd_mes" size="2" maxlength="2" value="<%=dt_mes%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
					<input type="text" name="cd_ano" size="4" maxlength="4" value="<%=dt_ano%>" class="inputs" onFocus="nextfield ='cd_unidade';"></td>
</tr>
<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<%if cd_pat_codigo <> "" Then%>
<tr>
	<td>&nbsp; Empresa:</td><%'strsql = "SELECT * FROM TBL_empresa WHERE cd_codigo = "&cd_unidade_origem&""
	'Set rs_empresa = dbconn.execute(strsql)
	
	strsql = "SELECT * FROM VIEW_patrimonio_movimentacao_lista WHERE (cd_patrimonio="&cd_pat_codigo&") and id_movimento = 1 and cd_status = 1"
	Set rs_mov = dbconn.execute(strsql)
		cd_comprador = rs_mov("cd_unidade_origem")
		cd_compra = rs_mov("cd_unidade_destino")
		%>
	<td>
		<select name="cd_comprador" class="inputs" onFocus="nextfield ='cd_unidade_compra';">
			<option value=""></option>
			<%strsql = "SELECT * FROM TBL_empresa order by cd_codigo"' WHERE cd_codigo = "&rs_mov("cd_unidade_origem")&""
			Set rs_empresa = dbconn.execute(strsql)
			while not rs_empresa.EOF
				cd_empresa = rs_empresa("cd_codigo")
				nm_empresa = rs_empresa("nm_empresa")%>
				<option value="<%=cd_empresa%>" <%if cd_comprador = cd_empresa Then%>Selected<%end if%>><%=nm_empresa%></option>
			<%rs_empresa.movenext
			wend%>			
	</td>
	<td>&nbsp; Destino :</td>
		<%strsql ="up_ADM_unidades_lista"
		Set rs_uni = dbconn.execute(strsql)%>
    <td>
		<select name="cd_unidade_compra" class="inputs" onFocus="nextfield ='cd_unidade';">
			<option value=""></option>
			<%Do While Not rs_uni.eof%>
			<option value="<%=rs_uni("cd_codigo")%>" <%if cd_compra = rs_uni("cd_codigo") then%>Selected<%end if%>><%=rs_uni("nm_unidade")%></option>
			<%rs_uni.movenext
			unidade_check = ""
			loop%>
		</select>		
	</td>
</tr>

<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>

<%end if%>
<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; Unidade (NOTA):</td>
		<%strsql ="up_ADM_unidades_lista"
		Set rs_uni = dbconn.execute(strsql)%>
    <td>
		<select name="cd_unidade" class="inputs" onFocus="nextfield ='cd_rack';">
			<option value=""></option>
			<%Do While Not rs_uni.eof
			if int(cd_unidade) = rs_uni("cd_codigo") then%><%unidade_check = "selected"%><%end if%>
			<option value="<%=rs_uni("cd_codigo")%>" <%=unidade_check%>><%=rs_uni("nm_unidade")%></option>
			<%rs_uni.movenext
			unidade_check = ""
			loop%>
		</select>		
	</td>
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
		</select></td>
</tr>
<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; Unidade (MOV):</td>
    <td><%=nm_unidade_comodato%><input type="hidden" name="cd_unidade_comodato" value="<%=cd_unidade_comodato%>"></td>
	<td>&nbsp; Rack:</td>
	<td><%=nm_rack_comodato%><input type="hidden" name="cd_rack_comodato" value="<%=cd_rack_comodato%>"></td>
</tr>
<!--<tr>
	<td>&nbsp; Unidade:</td>
		<%'strsql ="up_ADM_unidades_lista"
		'Set rs_uni = dbconn.execute(strsql)%>
    <td>
		<%'if cd_pat_codigo = "" Then%>
			<select name="cd_unidade" class="inputs" onFocus="nextfield ='cd_rack';">
				<option value=""></option>
				<%'Do While Not rs_uni.eof
				'if int(cd_unidade) = rs_uni("cd_codigo") then%><%'unidade_check = "selected"%><%'end if%>
				<option value="<%'=rs_uni("cd_codigo")%>" <%'=unidade_check%>><%'=rs_uni("nm_unidade")%></option>
				<%'rs_uni.movenext
				'unidade_check = ""
				'loop%>
				</select>
		<%'else%>
			&nbsp;<%'=nm_unidade%>
			<input type="hidden" name="cd_unidade" value="<%'=cd_unidade%>">
		<%'end if%>
	<%''="("&nm_unidade_comodato&")"%>
	<input type="hidden" name="cd_unidade_comodato" value="<%'=cd_unidade_comodato%>">
	</td>
	<td>&nbsp; Rack:</td>
	<%'strsql ="SELECT * FROM TBL_rack order by A056_desrac"
		'Set rs_rack = dbconn.execute(strsql)%>
	<td>
		<%'if cd_pat_codigo = "" Then%>
			<select name="cd_rack" class="inputs" onFocus="nextfield ='cd_especialidade';">
				<option value="" SELECTED>&nbsp;</option>
				<%'while not rs_rack.EOF%>
					<option value="<%'=rs_rack("A056_codrac")%>" <%'if int(cd_rack) = rs_rack("A056_codrac") then response.write("SELECTED") End if%>><%'=rs_rack("A056_desrac")%></option>
				<%'rs_rack.movenext
				'wend%>
			</select>
		<%'else%>
			&nbsp;<%'=nm_rack%>
			<input type="hidden" name="cd_rack" value="<%'=cd_rack%>">
		<%'end if%><br>
		<%''="("&nm_rack_comodato&")"%>
		<input type="hidden" name="cd_rack_comodato" value="<%'=cd_rack_comodato%>">
	</td>
</tr>-->
<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>

<tr style="background-color:<%=cor_tabela%>;">
    <td>&nbsp; Especialidade: </td>
		<%	selecao = " top 100 percent * "
			'tabela = " TBL_Espec"
			tabela = " TBL_Especialidades"
			condicoes = ""
			criterios = " Order by nm_especialidade"
			
			strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
		  	Set rs_espec = dbconn.execute(strsql)%>
	<td><select name="cd_especialidade" class="inputs" onFocus="nextfield ='num_hospital';">
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
	

	<td>&nbsp; Nº Hospital (CC):</td>
	<td><input type="text" name="num_hospital" size="20" value="<%=num_hospital%>" maxlength="20" class="inputs" onFocus="nextfield ='cd_preventiva';"><!--onFocus="nextfield ='dt_periodicidade_prev';"--></td>
</tr>
<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<%'if cd_usuario = 46 then
if cd_pat_codigo <> "" Then
	if cd_status = "1" then%>
		<tr style="background-color:silver;"><td align="center" colspan="4"><b>Movimentação do Patrimônio</b></td></tr>
			<%
			
			strsql = "SELECT * FROM TBL_patrimonio_emprestimos where cd_patrimonio = '"&cd_pat_codigo&"' AND cd_cancela is null"'order by cd_codigo desc"
			Set rs_pat_emprestimo = dbconn.execute(strsql)
			while not rs_pat_emprestimo.EOF
				cd_emprestimo = rs_pat_emprestimo("cd_codigo")
				cd_situacao = rs_pat_emprestimo("cd_situacao")
				nm_solicitante = rs_pat_emprestimo("nm_solicitante")
				cd_unidade_des = rs_pat_emprestimo("cd_unidade_des")
				'cd_rack_des = rs_pat_emprestimo("cd_rack_des")
				dt_emprestimo = rs_pat_emprestimo("dt_emprestimo")
				dt_devolucao = rs_pat_emprestimo("dt_retorno")
				nm_empr_obs = rs_pat_emprestimo("nm_obs")
				cd_cancela = rs_pat_emprestimo("cd_cancela")
			rs_pat_emprestimo.movenext
			wend
				if dt_emprestimo = "" and dt_devolucao = "" and cd_situacao = "" then
					cor_emprest = "#dfdfdf"
				elseif dt_emprestimo <> "" and dt_devolucao <> "" and cd_situacao <> "" then
					cor_emprest = "#dfdfdf"
				
				elseif dt_emprestimo <> "" and dt_devolucao = "" and cd_situacao <> "" then
					cor_emprest = "#ffff99"
				end if
				
				
					if dt_emprestimo = "" OR dt_emprestimo = null Then
						dia_emprestimo = ""
						mes_emprestimo = ""
						ano_emprestimo = ""
					else
						dia_emprestimo = zero(day(dt_emprestimo))
						mes_emprestimo = zero(month(dt_emprestimo))
						ano_emprestimo = year(dt_emprestimo)
						
					end if
					
					if dt_devolucao = "" Then
						dia_devolucao = ""
						mes_devolucao = ""
						ano_devolucao = ""
					else
						dia_devolucao = zero(day(dt_devolucao))
						mes_devolucao = zero(month(dt_devolucao))
						ano_devolucao = year(dt_devolucao)
						
					end if
				
				if dt_devolucao <> "" then
					cd_emprestimo = ""
					cd_situacao = 1
					cd_unidade_des = 0
					nm_solicitante = ""
					nm_empr_obs = ""
					
					dia_emprestimo = ""
					mes_emprestimo = ""
					ano_emprestimo = ""
					
					dia_devolucao = ""
					mes_devolucao = ""
					ano_devolucao = ""
				end if%>
				<input type="hidden" name="cd_emprestimo" value="<%=cd_emprestimo%>">
		<%if cd_emprestimo = ""  OR cd_cancela = "1" Then%>		
			<tr style="background-color:<%=cor_emprest%>;">
					
				<td>&nbsp; Situação:</td>
				<td>
					<select name="cd_situacao" onFocus="nextfield ='dia_emprestimo';">
						<option value="3" <%'if cd_situacao = 1 then response.write("SELECTED") end if%>>OK</option>
						<option value="2" <%'if cd_situacao = 3 then response.write("SELECTED") end if%>>Manutenção</option>
						<option value="4" <%'if cd_situacao = 2 then response.write("SELECTED") end if%>>Empréstimo</option>
						<option value="5" <%'if cd_situacao = 4 then response.write("SELECTED") end if%>>Comodato</option>
						<option value="1" <%'if cd_situacao = 1 then response.write("SELECTED") end if%>>Descarte</option>
						
					</select>
				</td>
				<td>&nbsp; Data</td>
				<td><input type="text" name="dia_emprestimo" size="2" maxlength="2" value="<%'=dia_emprestimo%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="mes_emprestimo" size="2" maxlength="2" value="<%'=mes_emprestimo%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="ano_emprestimo" size="4" maxlength="4" value="<%'=ano_emprestimo%>"  onFocus="nextfield ='cd_unidade_des';"></td>
			</tr>
			<tr style="background-color:<%=cor_emprest%>;">
				<td>Unidade destino</td>
					<%strsql = "SELECT * FROM TBL_unidades where cd_hospital >= 1 and cd_status=1"
					Set rs_uni = dbconn.execute(strsql)%>
			    <td><select name="cd_unidade_des" class="inputs" onFocus="nextfield ='cd_rack_des';">
								<option value=""></option>
								<%Do While Not rs_uni.eof
								str_unidade_des = rs_uni("cd_codigo")%>
								<option value="<%=rs_uni("cd_codigo")%>" <%'if int(cd_unidade_des)=str_unidade_des then response.write("SELECTED") End if%>><%=rs_uni("nm_unidade")%></option>
								<%rs_uni.movenext
								unidade_check = ""
								loop%>
								</select>					
				</td>
				<td>&nbsp; Rack destino:</td>
				<%strsql ="SELECT * FROM TBL_rack order by A056_desrac"
				Set rs_rack = dbconn.execute(strsql)%>
			<td>
				<select name="cd_rack_des" class="inputs" onFocus="nextfield ='nm_solicitante';">
					<option value="" SELECTED>&nbsp;</option>
					<%while not rs_rack.EOF%>
						<option value="<%=rs_rack("A056_codrac")%>" <%'if int(cd_rack_des) = rs_rack("A056_codrac") then response.write("SELECTED") End if%>><%=rs_rack("A056_desrac")%></option>
					<%rs_rack.movenext
					wend%>
				</select></td>
			</tr>
			<tr style="background-color:<%=cor_emprest%>;">
				<td>&nbsp; Solicitante:</td>
				<td colspan="3"><input type="text" name="nm_solicitante" value="<%'=nm_solicitante%>" size="50" maxlength="200" onFocus="nextfield ='nm_empr_obs';"></td>
			</tr>
		<%else%>
			<tr style="background-color:<%=cor_emprest%>;">
					
				<td>&nbsp; Situação:</td>
				<input type="hidden" name="cd_situacao" value="<%=cd_situacao%>">
				<td>&nbsp;
					<%if cd_situacao = 2 then%>Manutenção
					<%elseif cd_situacao = 3 then%>OK
					<%elseif cd_situacao = 4 then%>Empréstimo
					<%elseif cd_situacao = 5 then%>Comodato
					<%end if%>
				</td>
				<td>&nbsp; Solicitante:</td>
				<td>&nbsp;<%=nm_solicitante%></td>
			</tr>
			<tr style="background-color:<%=cor_emprest%>;">
				<td>&nbsp; Unidade destino:</td>
					<%strsql = "SELECT * FROM TBL_unidades where cd_codigo='"&cd_unidade_des&"'"
					Set rs_uni = dbconn.execute(strsql)%>
			    <td>&nbsp;<%if Not rs_uni.eof then%>
						<%=rs_uni("nm_unidade")%>
					<%end if%>					
				</td>
				<td>&nbsp; <%if cd_status = "4" then%>Rack destino:<%end if%></td>
			    <td>&nbsp;<%if cd_status = "4" then%><%=nm_rack_comodato%><%end if%></td>
			</tr>
			<tr style="background-color:<%=cor_emprest%>;">
				<td>&nbsp;
				<%if cd_situacao = 4 then%>Data empréstimo:
				<%elseif cd_situacao = 2 then%>Data manutenção:
				<%elseif cd_situacao = 5 then%>Início comodato:<%end if%>
				</td>
				<td>&nbsp;<%=dia_emprestimo%>/<%=mes_emprestimo%>/<%=ano_emprestimo%></td>	
		<%end if
			
			'if dia_emprestimo <> "" Then
			if cd_emprestimo <> "" then%>
				<td>&nbsp;
				<%if cd_situacao = 2 then%>Data devolução:
				<%elseif cd_situacao = 4 then%>Data retorno:
				<%elseif cd_situacao = 5 then%>Data retorno:<%end if%></td>
				<td><input type="text" name="dia_empr_devolucao" size="2" maxlength="2" value="<%=dia_devolucao%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="mes_empr_devolucao" size="2" maxlength="2" value="<%=mes_devolucao%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="ano_empr_devolucao" size="4" maxlength="4" value="<%=ano_devolucao%>"></td>
				<td colspan="2">&nbsp;</td>
			</tr>
			<%end if%>
		
			<tr style="background-color:<%=cor_emprest%>;">
				<td>&nbsp; Observações:</td>
				<td colspan="2">
					<%if cd_emprestimo = "" OR cd_cancela = "1" Then%>
						<textarea cols="35" rows="2" name="nm_empr_obs"><%'=nm_empr_obs%></textarea>
					<%else%>
						<%=nm_empr_obs%>
					<%end if%>
				</td>
				<td align="center"><%if cd_emprestimo <> "" OR cd_cancela <> "1" then%>
					<img src="../imagens/ic_reprovado.gif" alt="Cancela o emprestimo/comodato" width="10" height="12" border="0" onClick="javascript:Jscancela('<%=cd_emprestimo%>')"><%end if%>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
					<img src="../imagens/ic_lupa.gif" alt="Ver listagem de movimentação" width="13" height="14" border="0" onClick="window.open('patrimonio_movimentacao_lista.asp?cd_pat_codigo=<%=cd_pat_codigo%>','nome','width=550,height=500,scrollbars=1')"></td>
			</tr>
		
		
		<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
	<%Else%>
		<tr><td colspan="4" align="right"><span onClick="window.open('patrimonio_movimentacao_lista.asp?cd_pat_codigo=<%=cd_pat_codigo%>','nome','width=550,height=500,scrollbars=1')"><img src="../imagens/ic_lupa.gif" alt="Ver listagem de movimentação" width="13" height="14" border="0"> &nbsp; Ver a listagem de movimentação </span>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</td></tr>	
	<%end if%>

<%end if%>
<tr>
	<td colspan="4" align="left" style="border:1px solid black;">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		<b>MANUTENÇÕES</b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		<b>PERIODICIDADE</b> &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		<%if cd_pat_codigo <> "" Then%><b>PLANEJAMENTO</b><%end if%></td>
	<!--td style="border:1px solid black;">Periodicidade</td-->
</tr>
<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; <b>Preventiva:</b></td>
	<td colspan="3"><input type="checkbox" name="cd_preventiva" value="1" class="inputs" <%if cd_preventiva = 1 then response.write("CHECKED") end if%> onFocus="nextfield ='dt_periodicidade_prev';"><b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<b></b> <input type="text" name="dt_periodicidade_prev" size="3" maxlength="2" value="<%=dt_periodicidade_prev%>" class="inputs" onFocus="nextfield ='cd_calibracao';" ><b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	
	<%If cd_pat_codigo <> "" Then%>
		<!--select name="dt_plan_mes_prev" onFocus="nextfield ='dt_plan_ano_prev';">
			<option value=""></option>
		<%'for i = 1 to 12%>	
			<option value="<%'=i%>"><%'=ucase(left(mes_selecionado(i),3))%></option>
		<%'next%>
		</select>/<input type="text" name="dt_plan_ano_prev" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='cd_calibracao';">(mes/ano)-->
		(<b>Próx. Preventiva:</b>
		<%if cd_pat_codigo <> "" Then
			strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio='"&cd_pat_codigo&"' and cd_suspenso <> 1 AND tipo_manut = 1 order by dt_plan_prev desc"
				Set rs_prev = dbconn.execute(strsql)
					while not rs_prev.EOF
						dt_plan_prev = rs_prev("dt_plan_prev")
						'cd_confirma = rs_prev("cd_confirma")
							'if cd_confirma = 1 then
							
						
						
						%>
						<b style="color:#006600;"><%=Ucase(left(mes_selecionado(month(dt_plan_prev)),3))%>/<%=year(dt_plan_prev)%></b>
					<%rs_prev.movenext
					wend%>
		<%end if%>)
	<%End if%>
	</td>
</tr>
<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; <b>Calibração:</b></td>
	<td colspan="3"><input type="checkbox" name="cd_calibracao" value="1" class="inputs" <%if cd_calibracao = 1 then response.write("CHECKED") end if%> onFocus="nextfield ='dt_periodicidade_cal';"><b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<b></b> <input type="text" name="dt_periodicidade_cal" size="3" maxlength="2" value="<%=dt_periodicidade_cal%>" class="inputs" onFocus="nextfield ='cd_seg_elet';"><b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<%If cd_pat_codigo <> "" Then%>
		<!--select name="dt_plan_mes_cal" onFocus="nextfield ='dt_plan_ano_cal';">
			<option value=""></option>
		<%'for i = 1 to 12%>	
			<option value="<%'=i%>"><%'=ucase(left(mes_selecionado(i),3))%></option>
		<%'next%>
		</select>/<input type="text" name="dt_plan_ano_cal" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='cd_seg_elet';">(mes/ano)-->
		(<b>Próx. Calibração:</b>
		<%if cd_pat_codigo <> "" Then
			strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = 2 order by dt_plan_prev desc"
				Set rs_cal = dbconn.execute(strsql)
					while not rs_cal.EOF
						dt_plan_cal = rs_cal("dt_plan_prev")%>
					<b style="color:#00cc00;"><%=Ucase(left(mes_selecionado(month(dt_plan_cal)),3))%>/<%=year(dt_plan_cal)%></b>
					<%rs_cal.movenext
					wend%>
		<%end if%>)
	<%End if%>
	</td>
</tr>
<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; <b>Segurança Elet.:</b></td>
	<td colspan="3"><input type="checkbox" name="cd_seg_elet" value="1" class="inputs" <%if cd_seg_elet = 1 then response.write("CHECKED") end if%> onFocus="nextfield ='dt_periodicidade_elet';"><b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<b></b> <input type="text" name="dt_periodicidade_elet" size="3" maxlength="2" value="<%=dt_periodicidade_elet%>" class="inputs" onFocus="nextfield ='envia';"><b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<%If cd_pat_codigo <> "" Then%>
		<!--select name="dt_plan_mes_elet" onFocus="nextfield ='dt_plan_ano_elet';">
				<option value=""></option>
		<%'for i = 1 to 12%>	
			<option value="<%'=i%>"><%'=ucase(left(mes_selecionado(i),3))%></option>
		<%'next%>
		</select>/<input type="text" name="dt_plan_ano_elet" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='envia';">(mes/ano)-->
		(<b>Próx Seg. Elétrica:</b>
		<%if cd_pat_codigo <> "" Then
			strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = 3 order by dt_plan_prev desc"
				Set rs_elet = dbconn.execute(strsql)
					while not rs_elet.EOF
						dt_plan_elet = rs_elet("dt_plan_prev")%>
					<b style="color:#99cc00;"><%=Ucase(left(mes_selecionado(month(dt_plan_elet)),3))%>/<%=year(dt_plan_elet)%></b>
					<%rs_elet.movenext
					wend%>
		<%end if%>)
	<%End if%>
	</td>
</tr>
<tr style="background-color:<%=cor_tabela%>;"><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<tr style="background-color:<%=cor_tabela%>;">
    <td>&nbsp;</td>
    <td><%if cd_status = "1" then%><input type="submit" name="envia" value=" OK "><%end if%> <br><font color="red"><%=mensagem%></font></td>
	<td align="right" valign="middle"><input type="checkbox" name="nao_constar" value="2"  onFocus="nextfield ='envia';" <%if int(nao_constar) = 2 then%>checked<%end if%>></td>
	<td align="left" valign="middle"><i style="font-size:8px;">Marque este campo caso desejar que este item não<br>conste no inventário de equipamento do hospital</i></td>
</tr>
<%if cd_pat_codigo <> "" Then%>
<tr style="background-color:<%=cor_tabela%>;"><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<tr style="background-color:<%=cor_tabela%>;">
	<td colspan="4">
	<%id_origem = cd_pat_codigo'*** Este é o Código do objeto ***
	cd_origem = 5
	pag_retorno = ".."&mem_posicao
	variaveis = "../.."&mem_posicao&"?tipo=cadastro" '*** Pagina de retorno, com variaveis%>
	<!--#include file="../ocorrencia/ocorrencia_formulario.asp"-->
	</td>
</tr>
<%end if%>
</form>
<tr style="background-color:<%=cor_tabela%>;">
    <td colspan="4">&nbsp;</td>
</tr>
<tr>
	<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
	<td><img src="../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
	<td colspan="2" align="right">
		<%if cd_status = "1" AND cd_usuario = "4" OR cd_status = "1" AND cd_usuario = "46" then%>
			<img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="javascript:JsDelete('<%=cd_pat_codigo%>')">&nbsp;
		<%end if%>
	</td>
</tr>
</table>
<%dia = day(now)
mes = month(now)
ano = year(now)
cd_hora = hour(now)
cd_minuto = minute(now)
cd_segundo = second(now)
dt_hoje = mes&"/"&dia&"/"&ano&" "&cd_hora&":"&cd_minuto&":"&cd_segundo

	'dt_atualizacao = cd_mes&"/"&cd_dia&"/"&cd_ano&" "&cd_hora&":"&cd_minuto&":"&cd_segundo
	dt_data = dt_mes&"/"&dt_dia&"/"&dt_ano
	'dt_descarte = dt_atualizacao
	
	

'strsql = "SELECT * FROM TBL_patrimonio_movimentacao where cd_patrimonio='"&cd_pat_codigo&"' AND id_movimento = 1"
'			set rs = dbconn.execute(strsql)
'				if not rs.EOF then
'					existe_mov = 1
'				end if
'					if existe_mov < 1 then '*** Grava o primeiro movimento = Empresa vinculada à unidade ***
'						'*** MOVIMENTAÇÃO ***
'						'*** Grava a empresa da Holding vinculada à unidade de negócio ***
'						cd_empresa = 1'XXX -> Seleciona a empresa da holding vinculada à unidade de negócio XXX - *** NECESSITA ALTERAÇÃO ***
'						id_movimento = 1 '*** indica o primeiro movimento do recém cadastrado patrimônio ***
'						xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_empresa&"', @cd_unidade_destino='"&cd_empresa&"', @dt_ida='"&dt_data&"', @dt_retorno='"&dt_data&"', @id_movimento="&id_movimento&", @cd_user="&cd_usuario&", @dt_cadastro='"&dt_hoje&"'"
'						dbconn.execute(xsql)
'					end if
%>

<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1)
	{
	  if (confirm ("Tem certeza que deseja excluir?"))
		  {
			document.location.href('acoes/patrimonio_acao.asp?cd_apaga='+cod1+'&jan=<%=jan%>&cd_tipo=patrimonio&acao=apaga');
		  }
	}
	
function Jscancela(cod1)
	{
	  if (confirm ("Tem certeza que deseja cancelar o emprestimo/comodato?"))
		  {
			document.location.href('acoes/patrimonio_acao.asp?cd_cancela='+cod1+'&cd_tipo=emprestimo&cd_pat_codigo=<%=cd_pat_codigo%>&acao=cancela&jan=<%=jan%>&caminho=<%=caminho%>');
			/*if (confirm ("O cancelamento foi efetuado, deseja excluir este emprestimo/comodato?"))
		  		{
		  		document.location.href('acoes/patrimonio_acao.asp?cd_apaga='+cod1x+'&cd_tipo=patrimonio&acao=apaga');
				}
			else{
				
				}*/
		  }
	}
</SCRIPT>
</body>
</html>
