<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<head>
	<title>Inventário</title>
</head>
<script language="javascript">
<!--

function ordem(a,b,ordem_total){
	a=a;
	b=b;
	ordem_total=ordem_total
	
		if (ordem_total != ''){
			virg =  ',';
			}
		else{
		virg= ''
		}		
	ordem_total = ordem_total + virg
		if (a != ""){
			ordem_total = ordem_total + a;
			}
		if (b != ""){
			ordem_total = ordem_total.replace(b,'');
			}	
		
	document.form.ordem_res.value=a;
	document.form.ordem_total.value=(ordem_total.replace(",,", ","));
}
function limpa_ordem (){
	document.form.ordem_total.value='';
}
// -->	
</script>
<%
'***** 1.Verifica quais campos foram selecionados *****
cat_patr = request("cat_patr")
campos = request("campos")
	
	campo_0 = Instr(1,campos,"nm_tipo",0)
	campo_1 = Instr(1,campos,"cd_equipamento,nm_equipamento_novo",0)
	campo_2 = Instr(1,campos,"cd_pn",0)
	campo_3 = Instr(1,campos,"cd_ns",0)
	campo_4 = Instr(1,campos,"cd_marca,nm_marca",0)
	campo_5 = Instr(1,campos,"cd_especialidade,nm_especialidade",0)
	campo_6 = Instr(1,campos,"num_hospital",0)
	campo_7 = Instr(1,campos,"cd_rack,nm_rack",0)
	campo_8 = Instr(1,campos,"cd_unidade,nm_sigla,nm_unidade",0)
	campo_9 = Instr(1,campos,"dt_data",0)
	campo_10 = Instr(1,campos,"nao_constar",0)

'***** 2.Detalhamento da Busca **************************
if cat_patr = "BTx05_kt" Then
	tipo_patr = request("tipo_patr")
else
	if cat_patr = "BTx05_kt" then
		tipo_patr = "0"
	elseif cat_patr = "pRkct_9f" then
		tipo_patr = "E"
	elseif cat_patr = "Zg76$mv" then
		tipo_patr = "I"
	elseif cat_patr = "93_BctAA" then
		tipo_patr = "O"
	end if
	
	'tipo_patr = cat_patr
end if
		if tipo_patr <> "" then
			tipo_cond = "nm_tipo = '"&tipo_patr&"'"
			prox_cond = " AND "
		end if


patrimonio = request("patrimonio")
	if patrimonio <> "" then
	 	patr_cond = prox_cond&"no_patrimonio = '"&patrimonio&"'"
		prox_cond = " AND "
	end if
equipamento = request("equipamento")
	if equipamento <> "" then
	 	equip_cond = prox_cond&"nm_equipamento_novo LIKE '"&equipamento&"%'"
		prox_cond = " AND "
	end if
modelo = request("modelo")
	if modelo <> "" then
	 	model_cond = prox_cond&"cd_pn LIKE '"&modelo&"%'"
		prox_cond = " AND "
	end if
nserie = request("nserie")
	if nserie <> "" then
	 	nserie_cond = prox_cond&"cd_ns LIKE '"&nserie&"%'"
		prox_cond = " AND "
	end if
marca = request("marca")
	if marca > 0 then
	 	marca_cond = prox_cond&"cd_marca = '"&marca&"'"
		prox_cond = " AND "
	end if
especialidade = request("especialidade")
	if especialidade > 0 then
	 	espec_cond = prox_cond&" cd_especialidade = '"&especialidade&"' "
		prox_cond = " AND "
	end if
codhosp = request("codhosp")
	if codhosp <> "" then
	 	codhosp_cond = prox_cond&"num_hospital LIKE '"&codhosp&"%'"
		prox_cond = " AND "
	end if
rack = request("rack")
	if rack > 0 then
	 	rack_cond = prox_cond&" cd_rack = "&rack&" "
		prox_cond = " AND "
	end if
unidade = request("unidade")
	if unidade > "0" then
	 	unidade_cond = prox_cond&"cd_unidade = '"&unidade&"'"
		prox_cond = " AND "
	end if
dia_compra_i = request("dia_compra_i")
mes_compra_i = request("mes_compra_i")
ano_compra_i = request("ano_compra_i")
dia_compra_f = request("dia_compra_f")
mes_compra_f = request("mes_compra_f")
ano_compra_f = request("ano_compra_f")
	
	if dia_compra_i <> "" OR mes_compra_i <> "" OR dia_compra_f <> "" OR mes_compra_f <> "" then
		if dia_compra_f > ultimodiames(mes_compra_f,ano_compra_f) AND dia_compra_f <> "" then
			dia_compra_f = ultimodiames(mes_compra_f,ano_compra_f)
		end if
		'compra_cond = prox_cond&"dt_data between '"&mes_compra_i&"/"&dia_compra_i&"/"&ano_compra_i&"' AND '"&mes_compra_f&"/"&dia_compra_f&"/"&ano_compra_f&"'"
		compra_cond = prox_cond&"dt_data between '"&mes_compra_i&"/"&dia_compra_i&"/"&ano_compra_i&"' AND '"&mes_compra_f&"/"&dia_compra_f&"/"&ano_compra_f&" 23:59'"
		prox_cond = " AND "
	end if

nao_constar = request("nao_constar")
	if nao_constar < 2 Then
		constar_cond = prox_cond&"nao_constar = "&nao_constar&""
		prox_cond = " AND "
	end if


'***** 3.Ordem dos registros da Busca *******************
ordem_total = request("ordem_total")
	ver_string_ordem = instr(ordem_total,",")
		if ver_string_ordem = 1 then
			ordem_total = mid(ordem_total,2,len(ordem_toral))
		end if
	
			if ordem_total <> "" Then
				ordem = "ORDER BY "&ordem_total
			end if
%>
<body>
<br>
<br><!-- LISTAGEM DO PATRIMONIO -->
<table align="center" id="no_print">
	<tr><td colspan="100%"><a href="javascript:void(0);" onclick="window.open('patrimonio/patrimonio_cad.asp','nome','width=700','height=50'); return false;" >:: Novo ::</a></td></tr>
	<form action="patrimonio.asp" name="form" id="Form" method="post">
	<input type="hidden" name="tipo" value="busca">
	<input type="hidden" name="cat_patr" value="<%=cat_patr%>">
	<tr>
		<td colspan="100%" align="center">BUSCA AVANÇADA
		<%if tipo_patr = "E" Then%> - EQUIPAMENTOS
		<%elseif tipo_patr = "I" Then%> - INSTRUMENOS
		<%elseif tipo_patr = "O" Then%> - ÓTICAS
		<%end if%>
		</td>
	</tr>
	<tr style="background-color:gray; color:white;">
		<td>MOSTRAR</td>
		<td>DETALHE</td>
		<td>ORDEM</td>
	</tr>
	<tr>
		<td>:: <a href="javascript:selecionar_tudo()">Todos</a> :: <a href="javascript:deselecionar_tudo()">Nada</a> ::</td>
		<td><input type="hidden" name="ordem_res"></td>
		<td><input type="hidden" name="ordem_total" value="<%=ordem_total%>">
		<img src="../imagens/blackdot.gif" alt="Limpa orndenação" width="33" height="15" border="0" onClick="limpa_ordem()"></td>
	</tr>
	
	
<!--*** MOSTRA OS CAMPOS ***-->
<%if cat_patr = "BTx05_kt" Then%>
	<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'#e2e2e2');">
		<td><input type="checkbox" name="campos" value="nm_tipo" disabled checked readonly="1">Tipo:
		<td><select name="tipo_patr">
				<option value= "" <%if tipo_patr = ""  Then%>SELECTED<%end if%>>Todos</option>
				<option value="E" <%if tipo_patr = "E" Then%>SELECTED<%end if%>>Equipamento</option>
				<option value="I" <%if tipo_patr = "I" Then%>SELECTED<%end if%>>Instrumento</option>
				<option value="O" <%if tipo_patr = "O" Then%>SELECTED<%end if%>>Ótica</option>
			</select>
		</td>
		<td>&nbsp;</td>
	</tr>
<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="nm_tipo,no_patrimonio" <%if campo_0 <> 0 Then response.write("CHECKED") end if%>>Patrimonio:</td>		
		<td><input type="text" name="patrimonio" size="20" maxlength="20" value="<%=patrimonio%>" style="background-color:white;"></td>
		<td>
			<img src="../imagens/blackdot.gif" alt="Incluí" width="15" height="15" border="0" onClick="ordem('nm_tipo,no_patrimonio','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="exclui da ordem" width="15" height="15" border="0" onClick="ordem('','nm_tipo,no_patrimonio',document.form.ordem_total.value)">
		<!--input type="text" name="ordem_patrimonio" size="2" maxlength="2" readonly--></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_equipamento,nm_equipamento_novo" <%if campo_1 <> 0 Then response.write("CHECKED") end if%>>Nome</td>		
		<td><input type="text" name="equipamento" size="20" maxlength="20" value="<%=equipamento%>" style="background-color:white;"></td>
		<td>
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('nm_equipamento_novo','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('','nm_equipamento_novo',document.form.ordem_total.value)"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_pn" <%if campo_2 <> 0 Then response.write("CHECKED") end if%>>Modelo</td>
		<td><input type="text" name="modelo" size="20" maxlength="20" value="<%=modelo%>" style="background-color:white;"></td>
		<td>
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('cd_pn','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('','cd_pn',document.form.ordem_total.value)"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_ns" <%if campo_3 <> 0 Then response.write("CHECKED") end if%>>Nº Série</td>
		<td><input type="text" name="nserie" size="20" maxlength="20" value="<%=nserie%>" style="background-color:white;"></td>
		<td>
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('cd_ns','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('','cd_ns',document.form.ordem_total.value)"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_marca,nm_marca" <%if campo_4 <> 0 Then response.write("CHECKED") end if%>>Marca</td>
		<td><select name="marca" style="background-color:white;">
				<option value=0></option>
			<%strsql = "SELECT * FROM  TBL_marca "
		  	Set rs_marca = dbconn.execute(strsql)%>
			<%while not rs_marca.EOF
				cd_marca = rs_marca("cd_codigo")
				nm_marca = rs_marca("nm_marca")%>
				<option value="<%=cd_marca%>" <%if int(marca) = cd_marca then%>SELECTED<%end if%>><%=nm_marca%></option>
			<%rs_marca.movenext
			wend%>
			</select></td>
		<td>
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('nm_marca','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('','nm_marca',document.form.ordem_total.value)"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_especialidade,nm_especialidade" <%if campo_5 <> 0 Then response.write("CHECKED") end if%>>Especialidade</td>		
		<td><select name="especialidade" style="background-color:white;">
				<option value=0></option>
			<%strsql = "SELECT * FROM  TBL_especialidades where cd_status = 1 "
		  	Set rs_espec = dbconn.execute(strsql)%>
			<%while not rs_espec.EOF
				cd_espec = rs_espec("cd_codigo")
				nm_especialidade = rs_espec("nm_especialidade")%>
				<option value="<%=cd_espec%>" <%if int(especialidade) = cd_espec then%>SELECTED<%end if%>><%=nm_especialidade%></option>
			<%rs_espec.movenext
			wend%>
			</select></td>
		<td>
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('nm_especialidade','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('','nm_especialidade',document.form.ordem_total.value)"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="num_hospital" <%if campo_6 <> 0 Then response.write("CHECKED") end if%>>Nº Hospital</td>
		<td><input type="text" name="codhosp" size="20" maxlength="20" value="<%=codhosp%>" style="background-color:white;"></td>
		<td>
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('num_hospital','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('','num_hospital',document.form.ordem_total.value)"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_rack,nm_rack" <%if campo_7 <> 0 Then response.write("CHECKED") end if%>>Rack</td>
		<td><select name="rack" style="background-color:white;">
				<option value=0></option>
			<%strsql = "SELECT * FROM  TBL_rack order by a056_desrac"' where a056_status = 1 "
		  	Set rs_rack = dbconn.execute(strsql)%>
			<%while not rs_rack.EOF
				cd_rack = rs_rack("a056_codrac")
				nm_rack = rs_rack("a056_desrac")%>
				<option value="<%=cd_rack%>" <%if int(rack) = cd_rack then%>SELECTED<%end if%>><%=nm_rack%></option>
			<%rs_rack.movenext
			wend%>	
			</select></td>
		<td>
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('nm_rack','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('','nm_rack',document.form.ordem_total.value)"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_unidade,nm_sigla,nm_unidade" <%if campo_8 <> 0 Then response.write("CHECKED") end if%>>Unidade</td>
		<td><select name="unidade" style="background-color:white;">
				<option value=0></option>
			<%strsql = "SELECT * FROM  TBL_unidades where cd_status = 1 order by nm_unidade"
		  	Set rs_unidade = dbconn.execute(strsql)%>
			<%while not rs_unidade.EOF
				cd_unidade = rs_unidade("cd_codigo")
				nm_unidade = rs_unidade("nm_unidade")%>
				<option value="<%=cd_unidade%>" <%if int(unidade) = cd_unidade then%>SELECTED<%end if%>><%=nm_unidade%></option>
			<%rs_unidade.movenext
			wend%>
			</select></td>
		<td>
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('nm_sigla','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('','nm_sigla',document.form.ordem_total.value)"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="dt_data" <%if campo_9 <> 0 Then response.write("CHECKED") end if%>>Compra</td>
		<td>
			<input type="text" name="dia_compra_i" size="2" maxlength="2" value="<%=dia_compra_i%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" style="background-color:white;">/<input type="text" name="mes_compra_i" size="2" maxlength="2" value="<%=mes_compra_i%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" style="background-color:white;">/<input type="text" name="ano_compra_i" size="4" maxlength="4" value="<%=ano_compra_i%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" style="background-color:white;"> Até
			<input type="text" name="dia_compra_f" size="2" maxlength="2" value="<%=dia_compra_f%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" style="background-color:white;">/<input type="text" name="mes_compra_f" size="2" maxlength="2" value="<%=mes_compra_f%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" style="background-color:white;">/<input type="text" name="ano_compra_f" size="4" maxlength="4" value="<%=ano_compra_f%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" style="background-color:white;">
		</td>
		<td>
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('dt_data','',document.form.ordem_total.value)">
			<img src="../imagens/blackdot.gif" alt="" width="15" height="15" border="0" onClick="ordem('','dt_data',document.form.ordem_total.value)"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="nao_constar" <%if campo_10 <> 0 Then response.write("CHECKED") end if%>>Não constar</td>
		<td>
			<%if nao_constar = 0 Then
				ncons_ck0 = "CHECKED"
			elseif nao_constar = 1 Then
				ncons_ck1 = "CHECKED"
			else
				ncons_ckx = "CHECKED"
			end if%>
			<input type="radio" name="nao_constar" value="0" <%=ncons_ck0%>>constar
			<input type="radio" name="nao_constar" value="1" <%=ncons_ck1%>>não constar
			<input type="radio" name="nao_constar" value="2" <%=ncons_ckx%>>ambos
		</td>
	</tr>
	<tr><td><input type="submit" name="ok" value="Listar"></td></tr>
	<!--tr><td colspan="100%"><script language="jscript">window.document.write(document.form.ordem_total.value);</script>
</td></tr-->	
	</form>
</table>
<br>
<%if campos <> "" then%>
<table class="txt" border="0" width="10" align="center">
<tr bgcolor="#C0C0C0">
		<td>&nbsp;Nº</td>
	<%if campo_0 <> 0 Then%>
		<td><b>&nbsp;PAT</b></td>
	<%end if%>
	<%if campo_1 <> 0 Then%>
		<td><b>&nbsp;NOME</b></td>
	<%end if%>
	<%if campo_2 <> 0 Then%>
		<td><b>&nbsp;MODELO</b></td>
	<%end if%>
	<%if campo_3 <> 0 Then%>
		<td><b>&nbsp;Nº SÉRIE</b></td>
	<%end if%>
	<%if campo_4 <> 0 Then%>
		<td><b>&nbsp;MARCA</b></td>
	<%end if%>
	<%if campo_5 <> 0 Then%>
		<td><b>&nbsp;ESPECIALIDADE</b></td>
	<%end if%>
	<%if campo_6 <> 0 Then%>
		<td><b>&nbsp;Nº HOSPITAL</b></td>
	<%end if%>
	<%if campo_7 <> 0 Then%>
		<td><b>&nbsp;RACK</b></td>
	<%end if%>
	<%if campo_8 <> 0 Then%>
		<td><b>&nbsp;UNIDADE</b></td>
	<%end if%>
	<%if campo_9 <> 0 Then%>
		<td><b>&nbsp;COMPRA</b></td>
	<%end if%>
	
	
</tr>
<tr>
	<td colspan="100%"><img src="../imagens/px.gif" alt="" width="100%" height="1" border="0"></td>
</tr>
<%num_lista = 1
		
	'*** CONDIÇÕES ***
		if tipo_patr <> "" OR  patrimonio <> "" OR equipamento <> "" OR modelo <> "" OR nserie <> "" OR marca > 0  OR especialidade > 0 OR codhosp <> "" OR rack > 0 OR unidade > 0 OR dia_compra_i <> "" OR dia_compra_f <> ""  OR mes_compra_i <> ""  OR mes_compra_f <> "" OR nao_constar < 2 Then
			condicao = "WHERE "&tipo_cond&" "& patr_cond&" "&equip_cond&" "&model_cond&" "&nserie_cond&" "&marca_cond&" "&espec_cond&" "&codhosp_cond&" "&rack_cond&" "&unidade_cond&" "&compra_cond&" "&constar_cond&""'dt_descarte is null"
		'if patrimonio <> "" OR equipamento <> "" OR modelo <> "" OR nserie <> "" OR marca > 0  OR especialidade > 0 OR codhosp <> "" OR rack > 0 OR unidade > 0 Then
		'	condicao = "WHERE "& patr_cond&" "&equip_cond&" "&model_cond&" "&nserie_cond&" "&marca_cond&" "&espec_cond&" "&codhosp_cond&" "&rack_cond&" "&unidade_cond&" "'dt_descarte is null"
		end if		
		
	'*** SQL - BUSCA - *** 
		strsql = "SELECT cd_codigo,"&campos&" FROM View_patrimonio_lista "&condicao&" "&ordem&""
		Set rs_patrimonio = dbconn.execute(strsql)
			
		'*** Caso não haja nada no BD ***
		if rs_patrimonio.EOF Then%>
		<tr><td colspan="100%">Nada encontrado</td></tr>
		<%else
			Do while not rs_patrimonio.EOF
			'*** Código do Patrimonio - PERMANENTE ***
			cd_pat_codigo = rs_patrimonio("cd_codigo")
			'*****************************************
			
			'*** Patrimonio ***
			if campo_0 <> 0 Then
				nm_tipo = rs_patrimonio("nm_tipo")
				no_patrimonio = rs_patrimonio("no_patrimonio")
			end if
			
			'*** Equipamento ***
			if campo_1 <> 0 Then
				cd_equipamento = rs_patrimonio("cd_equipamento")
				nm_equipamento = rs_patrimonio("nm_equipamento_novo")
			end if
			
			'*** Modelo ***
			if campo_2 <> 0 Then
				cd_pn = rs_patrimonio("cd_pn")
			end if
			
			'*** Nº Série ***
			if campo_3 <> 0 Then
				cd_ns = rs_patrimonio("cd_ns")
			end if
			
			'*** Marca ***
			if campo_4 <> 0 Then
				cd_marca = rs_patrimonio("cd_marca")
				nm_marca = rs_patrimonio("nm_marca")
			end if
			
			'*** Especialidade ***
			if campo_5 <> 0 Then
				cd_especialidade = rs_patrimonio("cd_especialidade")
				nm_especialidade = rs_patrimonio("nm_especialidade")
			end if
			
			'*** Nº Hospital ***
			if campo_6 <> 0 Then
				num_hospital = rs_patrimonio("num_hospital")
			end if
			
			'*** Rack ***
			if campo_7 <> 0 Then
				cd_rack = rs_patrimonio("cd_rack")
				nm_rack = rs_patrimonio("nm_rack")
			end if
			
			'*** Unidade ***
			if campo_8 <> 0 Then
				cd_unidade = rs_patrimonio("cd_unidade")
				nm_sigla = rs_patrimonio("nm_sigla")
				nm_unidade = rs_patrimonio("nm_unidade")
			end if
			
			'*** Compra ***
			if campo_9 <> 0 Then
				dt_data = rs_patrimonio("dt_data")
			end if
			
			'*** Não constar ***
			if campo_10 <> 0 then
				'nao_constar = 
			end if
			%>
<!--tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio_cadastro.asp?cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&no_patrimonio=<%=no_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"-->
<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar','Alterações','width=700','height=600'); return false;">
	<td align="center" valign="top" bgcolor="#C0C0C0">&nbsp;<b><%=zero(num_lista)%></b></td>
		<%if campo_0 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_tipo%><%=no_patrimonio%></td>
   		<%end if%>
		<%if campo_1 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_equipamento%></td>
   		<%end if%>
		<%if campo_2 <> 0 Then%>
			<td valign="top">&nbsp;<%=cd_pn%></td>
   		<%end if%>
		<%if campo_3 <> 0 Then%>
			<td valign="top">&nbsp;<%=cd_ns%></td>
   		<%end if%>
		<%if campo_4 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_marca%></td>
   		<%end if%>
		<%if campo_5 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_especialidade%></td>
   		<%end if%>
		<%if campo_6 <> 0 Then%>
			<td valign="top">&nbsp;<%=num_hospital%></td>
   		<%end if%>
		<%if campo_7 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_rack%></td>
   		<%end if%>
		<%if campo_8 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_sigla%></td>
   		<%end if%>
		<%if campo_9 <> 0 Then%>
			<td valign="top">&nbsp;<%=zero(day(dt_data))&"/"&zero(month(dt_data))&"/"&year(dt_data)%></td>
   		<%end if%>
</tr>
			<%nm_especialidade = ""
			num_lista = num_lista + 1
			rs_patrimonio.movenext
			loop
		end if%>
<tr><td colspan="100%"><%'=campos%></td></tr>
<!--tr><td colspan="100%">SELECT cd_codigo,<%=campos%> <br>FROM View_patrimonio_lista <br><%=condicao%> &nbsp; <br> <%=ordem%></td></tr-->
<tr>
		<td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>	
	<%if campo_0 <> 0 Then%>
		<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
	<%end if%>
	<%if campo_1 <> 0 Then%>
	<td><img src="../imagens/px.gif" alt="" width="250" height="1" border="0"></td>
	<%end if%>
	<%if campo_2 <> 0 Then%>
		<td><img src="../imagens/px.gif" alt="" width="85" height="1" border="0"></td>	
	<%end if%>
	<%if campo_3 <> 0 Then%>
		<td><img src="../imagens/px.gif" alt="" width="115" height="1" border="0"></td>
	<%end if%>
	<%if campo_4 <> 0 Then%>
		<td><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
	<%end if%>
	<%if campo_5 <> 0 Then%>
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
	<%end if%>
	<%if campo_6 <> 0 Then%>
		<td><img src="../imagens/px.gif" alt="" width="85" height="1" border="0"></td>
	<%end if%>
	<%if campo_7 <> 0 Then%>
		<td><img src="../imagens/px.gif" alt="" width="120" height="1" border="0"></td>
	<%end if%>
	<%if campo_8 <> 0 Then%>
		<td><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
	<%end if%>
	<%if campo_9 <> 0 Then%>
		<td id="no_print"><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
	<%end if%>	
</tr>
<tr><td colspan="100%">&nbsp;</td></tr>


</table><br>

<%end if%>

</body>
</html>
