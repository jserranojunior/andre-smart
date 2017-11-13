<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../../js/ajax2.js"></SCRIPT-->
<html>

<%cd_user = session("cd_codigo")
pasta_loc = "patrimonio\"
arquivo_loc = "patrimonio_cad.asp"

no_patrimonio = request("no_patrimonio")
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
cd_estado = request("cd_estado")

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
str_nf = request("cd_nf")
	
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
	elseif jan = 2 then
		ini_caminho = "../patrimonio"
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
natureza_mov = request("natureza_mov")
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

/*function fill_dia_emprestimo(){
	dia_emprestimo = document.getElementById("teste").value;
	document.form.dia_emprestimo_hide.value=dia_emprestimo;
	}*/
// -->	
</script>

<script language="JavaScript">

function mov_emprestimo()
	{
		document.form.nm_solicitante_obj.value = document.getElementById("nm_solicitante").value;
		document.form.nm_observacoes_obj.value = document.getElementById("nm_observacoes").value;
		document.form.dia_movimento_obj.value = document.getElementById("dia_movimento").value;
		document.form.mes_movimento_obj.value = document.getElementById("mes_movimento").value;
		document.form.ano_movimento_obj.value = document.getElementById("ano_movimento").value;
		document.form.und_destino_obj.value = document.getElementById("und_destino").value;
		document.form.rack_destino_obj.value = '11'; // Sem Rack		
	}
	
function mov_comodato()
	{
		document.form.dia_movimento_obj.value = document.getElementById("dia_movimento").value;
		document.form.mes_movimento_obj.value = document.getElementById("mes_movimento").value;
		document.form.ano_movimento_obj.value = document.getElementById("ano_movimento").value;
		document.form.nm_observacoes_obj.value = document.getElementById("nm_observacoes").value;
		document.form.und_destino_obj.value = document.getElementById("und_destino").value;
		document.form.rack_destino_obj.value = document.getElementById("rack_destino").value;
	}
	
function mov_rack()
	{
		document.form.nm_solicitante_obj.value = document.getElementById("nm_solicitante").value;
		document.form.nm_observacoes_obj.value = document.getElementById("nm_observacoes").value;
		document.form.dia_movimento_obj.value = document.getElementById("dia_movimento").value;
		document.form.mes_movimento_obj.value = document.getElementById("mes_movimento").value;
		document.form.ano_movimento_obj.value = document.getElementById("ano_movimento").value;
		document.form.und_destino_obj.value = document.getElementById("unidade_atual").value;
		document.form.rack_destino_obj.value = document.getElementById("rack_destino").value;		
	}
	
function mov_estoque()
	{
		document.form.dia_movimento_obj.value = document.getElementById("dia_movimento").value;
		document.form.mes_movimento_obj.value = document.getElementById("mes_movimento").value;
		document.form.ano_movimento_obj.value = document.getElementById("ano_movimento").value;
	}
	
function mov_descarte()
	{
		document.form.dia_movimento_obj.value = document.getElementById("dia_movimento").value;
		document.form.mes_movimento_obj.value = document.getElementById("mes_movimento").value;
		document.form.ano_movimento_obj.value = document.getElementById("ano_movimento").value;
		document.form.nm_observacoes.value = document.getElementById("nm_observacoes").value;
	}
	
function mov_devol()
	{
		document.form.dia_movimento_obj.value = document.getElementById("dia_movimento").value;
		document.form.mes_movimento_obj.value = document.getElementById("mes_movimento").value;
		document.form.ano_movimento_obj.value = document.getElementById("ano_movimento").value;
	}
	
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
<!--#include file="../js/ajax2.asp"-->
<!--#include file="js/patrimonio.js"-->
<body>
<!--#include file="../includes/arquivo_loc.asp"-->
<table border="0" class="txt" class="no_print" style="font-size:12px; font-family:arial;" align="center">
<tr>
	<td>:: <a href="<%=str_novo%>&jan=<%=jan%>">Novo</a> :: <%if jan<> 1 Then%><a href="patrimonio.asp?tipo=busca&cat_patr=BTx05_kt">Busca</a><%end if%></td>
</tr>
<%
'sql = "SELECT * FROM "

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
				cd_nf = rs_patrimonio("cd_nf")
				
				cd_rack = rs_patrimonio("cd_rack")
				nm_rack = rs_patrimonio("nm_rack")
				
				cd_unidade_comodato = rs_patrimonio("cd_unidade_comodato")
				cd_rack_comodato = rs_patrimonio("cd_rack_comodato")
				nm_unidade_comodato = rs_patrimonio("nm_unidade_comodato")
				nm_rack_comodato = rs_patrimonio("nm_rack_comodato")
				
				num_hospital =  rs_patrimonio("num_hospital")
				cd_estado = rs_patrimonio("cd_estado")
				num_hospital = rs_patrimonio("num_hospital")
				cd_categoria = rs_patrimonio("cd_categoria")
				dt_data = rs_patrimonio("dt_data")
				
				cd_preventiva = rs_patrimonio("cd_preventiva")
				dt_periodicidade_prev = rs_patrimonio("dt_periodicidade_prev")
				cd_fornecedor_prev = rs_patrimonio("cd_fornecedor_prev")
				
				cd_calibracao = rs_patrimonio("cd_calibracao")
				dt_periodicidade_cal = rs_patrimonio("dt_periodicidade_cal")
				cd_fornecedor_cal = rs_patrimonio("cd_fornecedor_cal")
				
				cd_seg_elet = rs_patrimonio("cd_seg_elet")
				dt_periodicidade_elet = rs_patrimonio("dt_periodicidade_elet")
				cd_fornecedor_elet = rs_patrimonio("cd_fornecedor_elet")
				
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

Admin = 46%>
<form action="<%=ini_caminho%><%=caminho%>/acoes/patrimonio_acao.asp" method="post" name="form" id="form" onsubmit="return validar_patrimonio(document.form);" onSubmit="return checa(this);">
<input type="hidden<%if cd_user=Admin Then response.write("s")%>" name="jan" value="<%=jan%>">
<input type="hidden<%if cd_user=Admin Then response.write("s")%>" name="cd_tipo" value="patrimonio">
<input type="hidden<%if cd_user=Admin Then response.write("s")%>" name="cd_pat_codigo" value="<%=cd_pat_codigo%>">
<input type="hidden<%if cd_user=Admin Then response.write("s")%>" name="acao" value="<%=acao%>">
<input type="hidden<%if cd_user=Admin Then response.write("s")%>" name="caminho" value="<%=caminho%>">
<br>

<tr style="background-color:silver;">
	<td  align="center" colspan="4>&nbsp;<b>INVENTÁRIO - <span style="color:white;">CADASTRO DE PATRIMÔNIO <%=jan%></span></b></td>
</tr>
<tr style="background-color:<%=cor_emprest%>;">
				<td colspan="4"><span id='divPat_new'>&nbsp;</span></td>
			</tr>
<tr bgcolor="#cococo" style="background-color:<%=cor_tabela%>;">
	
	<%if existe = 1 then%>
		<td align="center"><img src="../imagens/px.gif"  height=1></td>
		<td align="center" style="background-color:#ff0033; color:#FFFFFF;" colspan="2">AVISO: ESTE NÚMERO DE PATRIMÔNIO JÁ EXISTE!</td>
		<td align="center"><input type="checkbox" name="cd_ciente" value="1">Ok, Estou ciente.</td>
	<%else%>
		<td align="center" colspan="4"><img src="../imagens/px.gif"  height=1></td>
	<%end if%>
	
</tr>
<!--tr><td align=center colspan="100%">&nbsp;</td></tr-->
<%'if %>
<tr style="background-color:<%=cor_tabela%>;">
    <td>&nbsp;<b> Nº Patrimônio:</b></td>
    <td><input type="text" name="no_patrimonio" value="<%=no_patrimonio%>" onblur="num_patrimonio();" class="inputs"></td>
    <td>&nbsp;Tipo:</td>
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
	&nbsp; &nbsp; Categoria: <select name="cd_categoria" class="inputs">
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
    <td><input type="text" name="cd_pn" value="<%=cd_pn%>" class="inputs"></td>
    <td>&nbsp; Num. Série:</td>
    <td><input type="text" name="cd_ns" size="30" value="<%=cd_ns%>" class="inputs">
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
	<td><select name="cd_marca" size="1" class="inputs">
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
    <td><input type="text" name="cd_dia" size="2" maxlength="2" value="<%=dt_dia%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)">/
					<input type="text" name="cd_mes" size="2" maxlength="2" value="<%=dt_mes%>" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)">/
					<input type="text" name="cd_ano" size="4" maxlength="4" value="<%=dt_ano%>" class="inputs"></td>
</tr>
<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; Nº Nota Fiscal:</td>
	<td><%if cd_nf <> "" AND str_nf = "" Then%>
			<%=cd_nf%>
		<%else
			cd_nf = replace(str_nf,"'","")%>
			<input type="text" name="cd_nf" value="<%=cd_nf%>" size="30" maxlength="50" class="inputs">
		<%end if%>
	</td>
	<td colspan="2">&nbsp;</td>
</tr>
<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<%if cd_pat_codigo <> "" Then%>
<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; Empresa:</td>
		
	<%strsql = "SELECT top (100) * FROM VIEW_patrimonio_movimentacao_lista WHERE (cd_patrimonio="&cd_pat_codigo&") and id_movimento = 1 and cd_status = 1"
	Set rs_mov = dbconn.execute(strsql)
		if not rs_mov.EOF Then
		cd_comprador = rs_mov("cd_unidade_origem")
		cd_compra = rs_mov("cd_unidade_destino")
		end if%>
	<td>
		<select name="cd_comprador" class="inputs">
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
		<select name="cd_unidade_compra" class="inputs">
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

<%else%>

<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; Empresa:</td>
	<td>
		<select name="cd_comprador" class="inputs">
			<option value=""></option>
			<%strsql = "SELECT * FROM TBL_empresa order by cd_codigo"'
			Set rs_empresa = dbconn.execute(strsql)
			while not rs_empresa.EOF
				cd_empresa = rs_empresa("cd_codigo")
				nm_empresa = rs_empresa("nm_empresa")%>
				<option value="<%=cd_empresa%>"><%=nm_empresa%></option>
			<%rs_empresa.movenext
			wend%>			
	</td>
	<td>&nbsp; Destino :</td>
		<%strsql ="up_ADM_unidades_lista"
		Set rs_uni = dbconn.execute(strsql)%>
    <td>
		<select name="cd_unidade_compra" class="inputs">
			<option value=""></option>
			<%Do While Not rs_uni.eof%>
			<option value="<%=rs_uni("cd_codigo")%>"><%=rs_uni("nm_unidade")%></option>
			<%rs_uni.movenext
			unidade_check = ""
			loop%>
		</select>		
	</td>
</tr>

<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>

<%end if%>

<tr style="background-color:<%=cor_tabela%>;">
    <td>&nbsp; Especialidade: </td>
		<%	selecao = " top 100 percent * "
			'tabela = " TBL_Espec"
			tabela = " TBL_Especialidades"
			condicoes = " where cd_status > 0"
			criterios = " Order by nm_especialidade"
			
			strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
		  	Set rs_espec = dbconn.execute(strsql)%>
	<td><select name="cd_especialidade" class="inputs">
						<option value=""></option>
						<%Do while Not rs_espec.EOF
							if int(cd_especialidade) = rs_espec("cd_codigo") then%><%espec_check = "selected"%><%end if%>							
								<option value="<%=rs_espec("cd_codigo")%>" <%=espec_check%>><%=rs_espec("nm_especialidade")%></option>
								<%rs_espec.movenext
								eqp_check = ""
								espec_check = ""
						Loop%>
					</select>
					<!--a href="javascript:;" onclick="window.open('janelas_cadastros/equipamento_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a-->
	</td>
	

	<td>&nbsp; Situação:</td>

	<%strsql ="SELECT  cd_codigo, nm_estado  FROM TBL_patrimonio_estado order by cd_codigo"
		  	Set rs_estado = dbconn.execute(strsql)%>
	<td><select name="cd_estado" class="inputs">						
						<%Do while Not rs_estado.EOF
							if int(cd_estado) = rs_estado("cd_codigo") then%><%estado_check = "selected"%><%end if%>							
								<option value="<%=rs_estado("cd_codigo")%>" <%=estado_check%>><%=rs_estado("nm_estado")%></option>
								<%rs_estado.movenext							
								estado_check = ""
						Loop%>
					</select>	
	<!--input type="text" name="num_hospital" size="20" value="<%=num_hospital%>" maxlength="20" class="inputs"---><!--onFocus="nextfield ='dt_periodicidade_prev';"--></td>
</tr>
<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>

<tr style="background-color:#b9c5b4<%'=cor_tabela%>;">
	<td>&nbsp; Unidade (NOTA):</td>
		<%'strsql ="up_ADM_unidades_lista"
	if cd_pat_codigo <> "" Then
	strsql ="SELECT * FROM TBL_unidades where cd_codigo="&cd_unidade&""' order by nm_unidade"
		Set rs_uni = dbconn.execute(strsql) %>
    <td>
		<%if not rs_uni.EOF Then
			response.write(rs_uni("nm_unidade"))
		end if%>
		<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
	</td>	
		<%else%>	
	<td><select name="cd_unidade" class="inputs">
			<option value=""></option>
			<%strsql ="SELECT * FROM TBL_unidades WHERE cd_status = 1"' OR  cd_hospital = 3 AND cd_status = 1"
			Set rs_uni = dbconn.execute(strsql) 
			Do While Not rs_uni.eof
				if int(cd_unidade) = rs_uni("cd_codigo") then%>
					<%unidade_check = "selected"%>
				<%end if%>
			<option value="<%=rs_uni("cd_codigo")%>" <%=unidade_check%>><%=rs_uni("nm_unidade")%></option>
		<%rs_uni.movenext
		unidade_check = ""
		loop%>
		</select>
	</td>
	<%end if%>
	<td>&nbsp; Rack:<%'="***"&cd_pat_codigo%></td>
	<%if cd_pat_codigo = "" Then%>
	
	<td>
		<%strsql ="SELECT * FROM TBL_rack"' where A056_codrac='"&cd_rack&"' order by A056_desrac"
		Set rs_rack = dbconn.execute(strsql)%>
		<select name="cd_rack" class="inputs">
			<option value="" SELECTED>&nbsp;</option>
			<%while not rs_rack.EOF
			cod_rack = rs_rack("A056_codrac")
			nom_rack = rs_rack("A056_desrac")%>
				<option value="<%=cod_rack%>" <%if int(cd_rack) = cod_rack then response.write("SELECTED") End if%>><%=nom_rack%></option>
			<%rs_rack.movenext
			wend%>
		</select>
		</td>
	<%else%>
	<td><%=nm_rack%><input type="hidden" name="cd_rack" value="<%=cd_rack%>"></td>
	<%end if%>
</tr>
<%if cd_pat_codigo <> "" Then%>
<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; Unidade (MOV):</td>
    <td><%=nm_unidade_comodato%><input type="hidden" name="cd_unidade_comodato" value="<%=cd_unidade_comodato%>"></td>
	<td>&nbsp; Rack:</td>
	<td><%=nm_rack_comodato%><input type="hidden" name="cd_rack_comodato" value="<%=cd_rack_comodato%>"></td>
</tr>
<%end if%>


<tr style="background-color:<%=cor_tabela%>;"><td align=center colspan="4"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0">&nbsp;</td></tr>
<%'if cd_usuario = 46 then
if cd_pat_codigo <> "" Then
	'if cd_status = "1" then%>
		<tr style="background-color:<%=cor_tabela%>;"><td align="center" colspan="4"><b>Ultimas 5 Movimentações do Patrimônio</b></td></tr>
		<tr style="background-color:<%=cor_tabela%>;">
			<td colspan="4">
				<table border="1" style="border:1px solid black; border-collapse:collapse;font-size:12px; font-family:arial;" align="center">
					<tr style="background-color:gray; color:white;">
							<td>&nbsp;<br><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
							<td>Origem<br><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>				
							<td>Destino<br><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
							<td>Rack<br><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
							<td>Ida<br><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
							<td>Retorno<br><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
							<td>Tipo<br><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>				
						</tr>
						
						<%'****************************
						'*** Lista as movimentações ***
						'******************************
						n_linha = 1
						pri_devol_aberto = 0
						nat_mov = 0
						ultimo_movimento = 0
						
						strsql = "SELECT top 5 * FROM VIEW_patrimonio_movimentacao_lista WHERE (cd_patrimonio="&cd_pat_codigo&")and cd_status = 1 order by id_movimento desc"
						Set rs = dbconn.execute(strsql)
						
						while not rs.EOF
							cd_movimentacao = rs("cd_codigo")
							if n_linha = 1 then mov_atual = cd_movimentacao
							cd_unidade_origem = rs("cd_unidade_origem")
							nm_unidade_origem = rs("nm_sigla_origem")
							cd_unidade_destino = rs("cd_unidade_destino")
							nm_unidade_destino = rs("nm_sigla_destino")
							cd_rack_origem = rs("cd_rack_origem")
							cd_rack_destino = rs("cd_rack_destino")
							nm_rack_destino = left(rs("nm_rack_destino"),17)
							dt_ida = rs("dt_ida")
								dt_ida = zero(day(dt_ida))&"/"&zero(month(dt_ida))&"/"&year(dt_ida)
							dt_retorno = rs("dt_retorno")
								if not isnull(dt_retorno)then 
									dt_retorno = zero(day(dt_retorno))&"/"&zero(month(dt_retorno))&"/"&year(dt_retorno)
								else
									dt_retorno = "-"
								end if
									
							cd_tipo_movimentacao = rs("cd_tipo_movimentacao")
							nm_tipo_movimentacao = rs("nm_situacao")
							natureza_mov = rs("cd_natureza")
								
							id_movimento = rs("id_movimento")
								if ultimo_movimento = 0 then
									ultimo_movimento = id_movimento
									ultima_unidade_destino = cd_unidade_destino
								end if
								
								if id_movimento = 1 then
									strsql = "SELECT * FROM TBL_empresa WHERE cd_codigo = "&cd_unidade_origem&""
									Set rs_empresa = dbconn.execute(strsql)
									if not rs_empresa.EOF Then nm_unidade_origem = rs_empresa("nm_empresa_abrv")
								end if
								
									if n_linha = 1 Then
										registro_atual = cd_movimentacao
										situacao_atual = cd_tipo_movimentacao
										unidade_atual = cd_unidade_destino
										devolucao_atual = rs("dt_retorno")
										'natureza_mov = rs("cd_natureza")
									end if
									
									if nat_mov < "1" AND dt_retorno = "-" then
										natureza_mov_anterior = "x"&cd_tipo_movimentacao&" - "&dt_retorno
										nat_mov = 1
									elseif nat_mov < "1" AND dt_retorno <> "-" then
										natureza_mov_anterior = "y"&natureza_mov&" - "&dt_retorno
										nat_mov = 1
									else
										natureza_mov_anterior = "z"&natureza_mov&" - "&dt_retorno
										nat_mov = 0
									end if
									
									if  pri_devol_aberto = 0 AND  dt_retorno = "-" Then
										cd_unidade_devolve = cd_unidade_origem
										cd_rack_devolve = cd_rack_origem
											pri_devol_aberto = 1
									end if
								
							n_linha = n_linha + 1%>
							<tr bgcolor="<%if id_movimento=1 then%>#dddddd<%else%>white<%end if%>">
								<td>&nbsp;<%=id_movimento%></td>
								<td><%=nm_unidade_origem%></td>				
								<td><%=nm_unidade_destino%></td>
								<td><%=nm_rack_destino%></td>
								<td><%=dt_ida%></td>
								<td><%=dt_retorno%></td>
								<td><%=nm_tipo_movimentacao%></td>
								<!--td><%=cd_movimentacao%></td-->
							</tr>
						<%if num_linha = 40 then
							cabeca = 0
						end if
						rs.movenext
						wend%>
				</table>
			</td>
		</tr>
		
	<%'if cd_status = "2" then
	if cd_status <> "" then%>
		<tr style="background-color:<%=cor_tabela%>;"><td colspan="4">&nbsp;</td></tr>
		
		
			<%'**********************************
			strsql = "SELECT top 1 * FROM TBL_patrimonio_movimentacao where cd_patrimonio = '"&cd_pat_codigo&"' AND dt_retorno IS NULL AND cd_status = 1 order by dt_ida desc"
			Set rs_pat_mov = dbconn.execute(strsql)
			while not rs_pat_mov.EOF
				cd_movimentacao = rs_pat_mov("cd_codigo")
				cd_emprestimo = rs_pat_mov("cd_codigo")
				cd_situacao = rs_pat_mov("cd_tipo_movimentacao")
				cd_unidade_des = rs_pat_mov("cd_unidade_destino")
				cd_rack_des = rs_pat_mov("cd_rack_destino")
				'response.write(cd_pat_codigo&"*"&cd_emprestimo&":"&cd_situacao&"-"&dt_devolucao&"lll<br>")
				
			rs_pat_mov.movenext
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
					
					'dia_emprestimo = ""
					'mes_emprestimo = ""
					'ano_emprestimo = ""
					
					dia_devolucao = ""
					mes_devolucao = ""
					ano_devolucao = ""
				end if%>
				<input type="hidden" name="cd_emprestimo" value="<%=cd_emprestimo%>">
				<input type="hidden" name="cd_mov_atual" value="<%=mov_atual%>">
			<tr style="background-color:<%=cor_tabela%>;">
				<td colspan="4"><%'=registro_atual&":"&situacao_atual&"-"&devolucao_atual&"-"&cd_unidade_des%>
				</td>
			</tr>	
		<%cor_emprest = "#ffffce"%>
		
		
		<tr>
			<td colspan="4">
		
				<table>
		
					<tr style="background-color:silver;">
						<td colspan="5">MOVIMENTAÇÃO</td>
					</tr>
					<tr>
						<td>Destino<br><img src="../imagens/px.gif" width="70" height="1"></td>
						<td>Rack<br><img src="../imagens/px.gif" width="125" height="1"></td>
						<td>Data<br><img src="../imagens/px.gif" width="180" height="1"></td>
						<td>Tipo<br><img src="../imagens/px.gif" width="220" height="1"></td>
					</tr>
					<tr>
					<!-- dados da origem-->
					<input type="hidden" name="id_movimento" value="<%=ultimo_movimento%>">
					<input type="hidden" name="unid_ori" value="<%=ultima_unidade_destino%>">
					
						<td>
						<select name="unid_des" size="1">
							<option value="">&nbsp;</option>
							<%strsql = "SELECT * FROM TBL_Unidades where cd_status='1' AND cd_hospital between 1 AND 2"
							Set rs_und = dbconn.execute(strsql)
							while not rs_und.EOF
								cd_unidade = rs_und("cd_codigo")
								nm_unidade = rs_und("nm_unidade")
								nm_sigla = rs_und("nm_sigla")
								'response.write(cd_pat_codigo&"*"&cd_emprestimo&":"&cd_situacao&"-"&dt_devolucao&"lll<br>")
									%>
									<option value="<%=cd_unidade%>"><%=nm_sigla%></option>
								
						
						
									<%
							rs_und.movenext
							wend
							%>
						</select>
						</td>
						<td>
						<select name="rack_des" size="1">
							<option value="">&nbsp;</option>
							<%strsql = "SELECT * FROM TBL_rack where A056_status='1'"
							Set rs_rack = dbconn.execute(strsql)
							while not rs_rack.EOF
								cd_rack = rs_rack("A056_codrac")
								nm_rack = rs_rack("A056_desrac")%>
									<option value="<%=cd_rack%>"><%=left(nm_rack,14)%></option>
								
						
						
							<%rs_rack.movenext
							wend%>
						</select>
						</td>
						<td>
							<input type="text" name="dia_mov" size="2" maxlength="2">/
							<input type="text" name="mes_mov" size="2" maxlength="2">/
							<input type="text" name="ano_mov" size="4" maxlength="4">
						</td>
						<td>
							<select name="tipo_mov" size="1">
								<option value="">&nbsp;</option>
								<option value="1">Descarte</option>
								<option value="2">Manutenção</option>
								<!--option value="3">&OK;</option-->
								<option value="4">Empréstimo</option>
								<option value="5">Comodato</option>
								<option value="6">Estoque</option>
								<option value="7">Rack</option>
								<option value="8">Cancela Comodato / Devolução</option>
							</select>
						</td>
					</tr>
				</table>	
		

<%
end if
end if%>
<tr>
	<td colspan="4" align="left" style="border:1px solid black;">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		<b>MANUTENÇÕES</b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		<b>PERIODICIDADE</b> &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		<b>ASSISTÊNCIA</b> &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
		<%if cd_pat_codigo <> "" Then%><b>PLANEJAMENTO</b><%end if%>
		<img src="../imagens/blackdot.gif" alt="" width="680" height="2" border="0"></td>
	<!--td style="border:1px solid black;">Periodicidade</td-->
</tr>
<tr style="background-color:<%=cor_tabela%>;">
	<td>&nbsp; <b>Preventiva:</b><%'=cd_pat_codigo%></td>
	<td colspan="3"><input type="checkbox" name="cd_preventiva" value="1" class="inputs" <%if cd_preventiva = 1 then response.write("CHECKED") end if%> ><b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<b></b> <input type="text" name="dt_periodicidade_prev" size="3" maxlength="2" value="<%=dt_periodicidade_prev%>" class="inputs"><b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<select name="cd_fornecedor_prev" style="width:180px;">
		<option value="0">&nbsp;</option>
		<%strsql ="SELECT * FROM TBL_fornecedor WHERE cd_status = 1 order by nm_fornecedor"
		Set rs_forn = dbconn.execute(strsql)
			while not rs_forn.EOF
			cod_fornecedor_prev = rs_forn("cd_codigo")
			fornecedor_prev = rs_forn("nm_fornecedor")
			%>
			<option value="<%=cod_fornecedor_prev%>" <%if cd_fornecedor_prev = cod_fornecedor_prev then response.write("Selected")%>><%=fornecedor_prev%></option>
			<%rs_forn.movenext
			wend%>
	</select>
	<%If cd_pat_codigo <> "" Then%>
		<!--select name="dt_plan_mes_prev" onFocus="nextfield ='dt_plan_ano_prev';">
			<option value=""></option>
		<%'for i = 1 to 12%>	
			<option value="<%'=i%>"><%'=ucase(left(mes_selecionado(i),3))%></option>
		<%'next%>
		</select>/<input type="text" name="dt_plan_ano_prev" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='cd_calibracao';">(mes/ano)-->
		
		(<b>Próx. Prev.:</b>
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
	<td colspan="3"><input type="checkbox" name="cd_calibracao" value="1" class="inputs" <%if cd_calibracao = 1 then response.write("CHECKED") end if%> ><b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<b></b> <input type="text" name="dt_periodicidade_cal" size="3" maxlength="2" value="<%=dt_periodicidade_cal%>" class="inputs"><b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<select name="cd_fornecedor_cal" style="width:180px;">
		<option value="0">&nbsp;</option>
		<%strsql ="SELECT * FROM TBL_fornecedor WHERE cd_status = 1 order by nm_fornecedor"
		Set rs_forn = dbconn.execute(strsql)
			while not rs_forn.EOF
			cod_fornecedor_cal = rs_forn("cd_codigo")
			fornecedor_cal = rs_forn("nm_fornecedor")
			%>
			<option value="<%=cod_fornecedor_cal%>" <%if cd_fornecedor_cal = cod_fornecedor_cal then response.write("Selected")%>><%=fornecedor_cal%></option>
			<%rs_forn.movenext
			wend%>
	</select>
<%If cd_pat_codigo <> "" Then%>
		<!--select name="dt_plan_mes_cal" onFocus="nextfield ='dt_plan_ano_cal';">
			<option value=""></option>
		<%'for i = 1 to 12%>	
			<option value="<%'=i%>"><%'=ucase(left(mes_selecionado(i),3))%></option>
		<%'next%>
		</select>/<input type="text" name="dt_plan_ano_cal" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='cd_seg_elet';">(mes/ano)-->
		(<b>Próx. Cal.:</b>
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
	<td colspan="3"><input type="checkbox" name="cd_seg_elet" value="1" class="inputs" <%if cd_seg_elet = 1 then response.write("CHECKED") end if%>><b>Aplicavel?</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<b></b> <input type="text" name="dt_periodicidade_elet" size="3" maxlength="2" value="<%=dt_periodicidade_elet%>" class="inputs"><b>Meses</b> &nbsp; &nbsp; &nbsp; &nbsp; 
	<select name="cd_fornecedor_elet" style="width:180px;">
		<option value="0">&nbsp;</option>
		<%strsql ="SELECT * FROM TBL_fornecedor WHERE cd_status = 1 order by nm_fornecedor"
		Set rs_forn = dbconn.execute(strsql)
			while not rs_forn.EOF
			cod_fornecedor_elet = rs_forn("cd_codigo")
			fornecedor_elet = rs_forn("nm_fornecedor")
			%>
			<option value="<%=cod_fornecedor_elet%>" <%if cd_fornecedor_elet = cod_fornecedor_elet then response.write("Selected")%>><%=fornecedor_elet%></option>
			<%rs_forn.movenext
			wend%>
	</select>
	<%If cd_pat_codigo <> "" Then%>
		<!--select name="dt_plan_mes_elet" onFocus="nextfield ='dt_plan_ano_elet';">
				<option value=""></option>
		<%'for i = 1 to 12%>	
			<option value="<%'=i%>"><%'=ucase(left(mes_selecionado(i),3))%></option>
		<%'next%>
		</select>/<input type="text" name="dt_plan_ano_elet" size="4" maxlength="4"  class="inputs" onFocus="nextfield ='envia';">(mes/ano)-->
		(<b>Próx Seg.:</b>
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
    <td><%'if cd_status <> "" then%><input type="submit" name="envia" value=" OK "><%'end if%> <br><font color="red"><%=mensagem%></font></td>
	<td align="right" valign="middle"><input type="checkbox" name="nao_constar" value="2" <%if int(nao_constar) = 2 then%>checked<%end if%>></td>
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
	<%if cd_status < 2 then%>
		<!--#include file="../ocorrencia/ocorrencia_formulario.asp"-->
	<%elseif cd_status = 2 then%>
		<!--#include file="../ocorrencia/ocorrencia_visualizacao.asp"-->
	<%end if%>
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
	
function Jscancela(cod1,cod2)
	{
	  if (confirm ("Tem certeza que deseja cancelar o emprestimo/comodato?"))
		  {


			document.location.href('acoes/patrimonio_acao.asp?cd_cancela='+cod1+'&cd_situacao='+cod2+'&cd_tipo=emprestimo&cd_pat_codigo=<%=cd_pat_codigo%>&acao=cancela&jan=<%=jan%>&caminho=<%=caminho%>');
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
