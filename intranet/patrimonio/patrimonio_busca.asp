<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<head>
	<title>Inventário</title>
</head>
<script language="javascript">
<!--
function ordem(a,ordem_total,obj){
	a=a;
	objeto=obj;
	ordem_total=ordem_total;
		if (ordem_total != ''){
			virg =  ',';
			}
		else{
		virg= ''
		}		
	ordem_total = ordem_total + virg
		if (a != ""){
			ordem_total = ordem_total + a;}
		
	document.form.ordem_res.value=a;
	document.form.ordem_total.value=(ordem_total.replace(",,", ","));		
	
	var el=document.getElementById(objeto);
	el.style.display=(el.style.display!='none'?'none':'');
	
	document.getElementsByName(objeto)[1].value= document.form.ordem_inicial.value+'º';
	document.form.ordem_inicial.value = document.form.ordem_inicial.value*1+1;
}
function limpa_ordem(obj){
	document.form.ordem_total.value='';
	
	document.form.ordem_1.value='';
	document.form.ordem_2.value='';
	document.form.ordem_3.value='';
	document.form.ordem_4.value='';
	document.form.ordem_5.value='';
	document.form.ordem_6.value='';
	document.form.ordem_7.value='';
	document.form.ordem_8.value='';
	document.form.ordem_9.value='';
	document.form.ordem_10.value='';
	document.form.ordem_11.value='';
	
	document.form.ordem_inicial.value='1';
	
	document.ordem_1.style.display='inline';
	document.ordem_2.style.display='inline';
	document.ordem_3.style.display='inline';
	document.ordem_4.style.display='inline';
	document.ordem_5.style.display='inline';
	document.ordem_6.style.display='inline';
	document.ordem_7.style.display='inline';
	document.ordem_8.style.display='inline';
	document.ordem_9.style.display='inline';
	document.ordem_10.style.display='inline';
	document.ordem_11.style.display='inline';
}
// -->	
</script>
<%
usuario = session("cd_codigo")
jan = request("jan")

'***** 1.Verifica quais campos foram selecionados *****
cat_patr = request("cat_patr")
campos = request("campos")
	
	campo_0 = Instr(1,campos,"nm_tipo",0)
	campo_1 = Instr(1,campos,"cd_equipamento,nm_equipamento_novo",0)
	campo_2 = Instr(1,campos,"cd_pn",0)
	campo_3 = Instr(1,campos,"cd_ns",0)
	campo_4 = Instr(1,campos,"cd_marca,nm_marca",0)
	campo_5 = Instr(1,campos,"cd_especialidade,nm_especialidade",0)
	campo_6 = Instr(1,campos,"cd_estado,nm_estado",0)
	campo_7 = Instr(1,campos,"cd_rack_comodato,nm_rack_comodato",0)
	campo_8 = Instr(1,campos,"cd_unidade_comodato,nm_sigla_comodato,nm_unidade_comodato",0)
	campo_9 = Instr(1,campos,"dt_data",0)
	campo_10 = Instr(1,campos,"nao_constar",0)
	campo_11 = Instr(1,campos,"cd_status",0)

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
estado = request("estado")
	if trim(estado) <> "" then
	 	estado_cond = prox_cond&" cd_estado = '"&estado&"'"
		prox_cond = " AND "
	end if
rack = request("rack")
	if rack > 0 then
	 	rack_cond = prox_cond&" cd_rack_comodato = "&rack&" "
		prox_cond = " AND "
	end if
unidade = request("unidade")
	if unidade > "0" then
	 	unidade_cond = prox_cond&"cd_unidade_comodato = '"&unidade&"'"
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
	if nao_constar = 0 Then ' Mostra ambos
		constar_cond = prox_cond&"nao_constar between 1 and 2"
		prox_cond = " AND "
	elseif nao_constar = 1 Then 'Somente os que devem constar nas listas
		constar_cond = prox_cond&"nao_constar = "&nao_constar&""
		prox_cond = " AND "
	elseif nao_constar = 2 Then 'Somente os que não devem constar nas listas
		constar_cond = prox_cond&"nao_constar = "&nao_constar&""
		prox_cond = " AND "
	end if

cd_status = request("cd_status")
	if cd_status = "0" Then ' Mostra ambos
		status_cond = prox_cond&"cd_status between 1 and 2"
		prox_cond = " AND "
	elseif cd_status = "1" Then 'Somente os patrimônios ativos
		status_cond = prox_cond&"cd_status = "&cd_status&""
		prox_cond = " AND "
	elseif cd_status = "2" Then 'Somente os que não devem constar nas listas
		status_cond = prox_cond&"cd_status = "&cd_status&""
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
		
		ordem_inicial = request("ordem_inicial")
		ordem_1 = request("ordem_1")
		ordem_2 = request("ordem_2")
		ordem_3 = request("ordem_3")
		ordem_4 = request("ordem_4")
		ordem_5 = request("ordem_5")
		ordem_6 = request("ordem_6")
		ordem_7 = request("ordem_7")
		ordem_8 = request("ordem_8")
		ordem_9 = request("ordem_9")
		ordem_10 = request("ordem_10")
		ordem_11 = request("ordem_11")
		
%>
<body>
<br>
<br><!-- LISTAGEM DO PATRIMONIO -->
<table align="center" id="no_print">
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">patrimonio_busca.asp</span></td></tr>
	<%end if%>
	<tr><td colspan="100%"><%if usuario = 4 OR usuario = 46 then%><a href="javascript:void(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?jan=1&caminho=patrimonio','nome','width=730,height=600,scrollbars=1'); return false;">:: Novo ::</a><%else%>&nbsp;<%end if%></td></tr>
	<form action="patrimonio.asp" name="form" id="Form" method="post">
	<input type="hidden" name="tipo" value="busca">
	<input type="hidden" name="cat_patr" value="<%=cat_patr%>">
	<input type="hidden" name="ordem_res">
	<input type="hidden" name="ordem_total" value="<%=ordem_total%>">
	<input type="hidden" name="ordem_inicial" value="<%if ordem_inicial = "" Then%>1<%else response.write(ordem_inicial) end if%>">
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
		<td>&nbsp;</td>
		<td><a href="javascript:void();" onClick="limpa_ordem()">Limpa Ordem</a></td>
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
		<td><%ordem_var = ordem_1
		ordem_num="ordem_1"
		ordem_campo="nm_tipo"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="nm_tipo,no_patrimonio" <%if campo_0 <> 0 Then response.write("CHECKED") end if%>>Patrimonio:</td>		
		<td><input type="text" name="patrimonio" size="20" maxlength="20" value="<%=patrimonio%>" style="background-color:white;"></td>
		<td>
		<%ordem_var = ordem_2
		ordem_num="ordem_2"
		ordem_campo="no_patrimonio"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_equipamento,nm_equipamento_novo" <%if campo_1 <> 0 Then response.write("CHECKED") end if%>>Nome</td>		
		<td><input type="text" name="equipamento" size="20" maxlength="20" value="<%=equipamento%>" style="background-color:white;"></td>
		<td>
		<%ordem_var = ordem_3
		ordem_num="ordem_3"
		ordem_campo="nm_equipamento_novo"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_pn" <%if campo_2 <> 0 Then response.write("CHECKED") end if%>>Modelo</td>
		<td><input type="text" name="modelo" size="20" maxlength="20" value="<%=modelo%>" style="background-color:white;"></td>
		<td>
		<%ordem_var = ordem_4
		ordem_num="ordem_4"
		ordem_campo="cd_pn"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_ns" <%if campo_3 <> 0 Then response.write("CHECKED") end if%>>Nº Série</td>
		<td><input type="text" name="nserie" size="20" maxlength="20" value="<%=nserie%>" style="background-color:white;"></td>
		<td>
		<%ordem_var = ordem_5
		ordem_num="ordem_5"
		ordem_campo="cd_ns"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_marca,nm_marca" <%if campo_4 <> 0 Then response.write("CHECKED") end if%>>Marca</td>
		<td><select name="marca" style="background-color:white;">
				<option value=0></option>
			<%strsql = "SELECT * FROM  TBL_marca order by nm_marca"
		  	Set rs_marca = dbconn.execute(strsql)%>
			<%while not rs_marca.EOF
				cd_marca = rs_marca("cd_codigo")
				nm_marca = rs_marca("nm_marca")%>
				<option value="<%=cd_marca%>" <%if int(marca) = cd_marca then%>SELECTED<%end if%>><%=nm_marca%></option>
			<%rs_marca.movenext
			wend%>
			</select></td>
		<td>
		<%ordem_var = ordem_6
		ordem_num="ordem_6"
		ordem_campo="nm_marca"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
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
		<%ordem_var = ordem_7
		ordem_num="ordem_7"
		ordem_campo="nm_especialidade"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_estado,nm_estado" <%if campo_6 <> 0 Then response.write("CHECKED") end if%>>Situação</td>		
		<td><select name="estado" style="background-color:white;">
				<option value=""</option>
			<%strsql = "SELECT * FROM  TBL_patrimonio_estado "
		  	Set rs_estado = dbconn.execute(strsql)%>
			<%while not rs_estado.EOF
				cd_estado = rs_estado("cd_codigo")
				nm_estado = rs_estado("nm_estado")%>
				<option value="<%=cd_estado%>" <%if int( 0 & estado) = cd_estado then%>SELECTED<%end if%>><%=nm_estado%></option>
			<%rs_estado.movenext
			wend%>
			</select></td>
		<td>
		<%ordem_var = ordem_8
		ordem_num="ordem_8"
		ordem_campo="nm_estado"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_rack_comodato,nm_rack_comodato" <%if campo_7 <> 0 Then response.write("CHECKED") end if%>>Rack</td>
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
		<%ordem_var = ordem_9
		ordem_num="ordem_9"
		ordem_campo="nm_rack_comodato"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_unidade_comodato,nm_sigla_comodato,nm_unidade_comodato" <%if campo_8 <> 0 Then response.write("CHECKED") end if%>>Unidade</td>
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
		<%ordem_var = ordem_10
		ordem_num="ordem_10"
		ordem_campo="nm_sigla_comodato"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="dt_data" <%if campo_9 <> 0 Then response.write("CHECKED") end if%>>Compra</td>
		<td>
			<input type="text" name="dia_compra_i" size="2" maxlength="2" value="<%=dia_compra_i%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" style="background-color:white;">/<input type="text" name="mes_compra_i" size="2" maxlength="2" value="<%=mes_compra_i%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" style="background-color:white;">/<input type="text" name="ano_compra_i" size="4" maxlength="4" value="<%=ano_compra_i%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" style="background-color:white;"> Até
			<input type="text" name="dia_compra_f" size="2" maxlength="2" value="<%=dia_compra_f%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" style="background-color:white;">/<input type="text" name="mes_compra_f" size="2" maxlength="2" value="<%=mes_compra_f%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" style="background-color:white;">/<input type="text" name="ano_compra_f" size="4" maxlength="4" value="<%=ano_compra_f%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" style="background-color:white;">
		</td>
		<td>
		<%ordem_var = ordem_11
		ordem_num="ordem_11"
		ordem_campo="dt_data"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="nao_constar" <%if campo_10 <> 0 Then response.write("CHECKED") end if%>>Não constar</td>
		<td>
			<%if nao_constar = 0 Then
				ncons_ck0 = "CHECKED"
			elseif nao_constar = 1 Then
				ncons_ck1 = "CHECKED"
			elseif nao_constar = 2 Then
				ncons_ck2 = "CHECKED"
			end if%>
			<input type="radio" name="nao_constar" value="0" <%=ncons_ck0%>>ambos
			<input type="radio" name="nao_constar" value="1" <%=ncons_ck1%>>constar
			<input type="radio" name="nao_constar" value="2" <%=ncons_ck2%>>não constar
			
		</td>
	</tr>
	<!--tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td><input type="checkbox" name="campos" value="cd_status" <%if campo_11 <> 0 Then response.write("CHECKED") end if%>>Situação</td>
		<td>
			<%if cd_status = 0 Then
				status_ck0 = "CHECKED"
			elseif cd_status = 1 Then
				status_ck1 = "CHECKED"
			elseif cd_status = 2 Then
				status_ck2 = "CHECKED"
			end if%>
			<input type="radio" name="cd_status" value="0" <%=status_ck0%>>Ambos
			<input type="radio" name="cd_status" value="1" <%=status_ck1%>>Ativos
			<input type="radio" name="cd_status" value="2" <%=status_ck2%>>Descartado
			
		</td>
	</tr-->
	<tr><td><input type="submit" name="ok" value="Listar"></td></tr>
	<!--tr><td colspan="100%"><script language="jscript">window.document.write(document.form.ordem_total.value);</script>
</td></tr-->	
	</form>
</table>
<br>
<%if campos <> "" then%>
<table class="txt" border="0" width="10" align="center">
<%'if usuario = "460" then%><!--tr><td colspan="10"><%="<b>SELECT</b> cd_codigo,"&campos&" <br><b>FROM</b> View_patrimonio_lista "&condicao&" <br>"&ordem&""%></td></tr--><%'end if%>
	
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
		<td><b>&nbsp;SITUAÇÂO</b></td>
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
		if tipo_patr <> "" OR  patrimonio <> "" OR equipamento <> "" OR modelo <> "" OR nserie <> "" OR marca > 0  OR especialidade > 0 OR codhosp <> "" OR rack > 0 OR unidade > 0 OR dia_compra_i <> "" OR dia_compra_f <> ""  OR mes_compra_i <> ""  OR mes_compra_f <> "" OR nao_constar >= 1 OR cd_status >= 1 Then
			condicao = "WHERE "&tipo_cond&" "& patr_cond&" "&equip_cond&" "&model_cond&" "&nserie_cond&" "&marca_cond&" "&espec_cond&" "&estado_cond&" "&rack_cond&" "&unidade_cond&" "&compra_cond&" "&constar_cond&" "&status_cond&""'dt_descarte is null"
		end if		
		
	'*** SQL - BUSCA - *** 
		strsql = "SELECT cd_codigo,cd_status,"&campos&" FROM View_patrimonio_lista "&condicao&" "&ordem&""
		'response.write strsql
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
				cd_estado = rs_patrimonio("cd_estado")
				nm_estado = rs_patrimonio("nm_estado")
			end if
			
			'*** Rack ***
			if campo_7 <> 0 Then
				'cd_rack = rs_patrimonio("cd_rack")
				'nm_rack = rs_patrimonio("nm_rack")
				
				cd_rack = rs_patrimonio("cd_rack_comodato")
				nm_rack = rs_patrimonio("nm_rack_comodato")
			end if
			
			'*** Unidade ***
			if campo_8 <> 0 Then
				'cd_unidade = rs_patrimonio("cd_unidade")
				'nm_sigla = rs_patrimonio("nm_sigla")
				'nm_unidade = rs_patrimonio("nm_unidade")
				
				cd_unidade = rs_patrimonio("cd_unidade_comodato")
				nm_sigla = rs_patrimonio("nm_sigla_comodato")
				nm_unidade = rs_patrimonio("nm_unidade_comodato")
			end if
			
			'*** Compra ***
			if campo_9 <> 0 Then
				dt_data = rs_patrimonio("dt_data")
			end if
			
			'*** Não constar ***
			if campo_10 <> 0 then
				'nao_constar = 
			end if
			
			'*** Descarte ***
			'if campo_110 <> 0 then
				cd_status = rs_patrimonio("cd_status")
				if cd_status = "2" then
					stilo_linha = "color:silver;"
				else
					stilo_linha = "color:black;"
				end if
			'end if
			
			if usuario = 4 OR  usuario = 46 then
				abre_pag = "patrimonio_cad"
				tamanho_janela = "600"
			else
				abre_pag = "patrimonio_consulta"
				tamanho_janela = "500"
			end if
			%>
<!--tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio_cadastro.asp?cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&no_patrimonio=<%=no_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"-->

<tr onclick="javascript:window.open('patrimonio/<%=abre_pag%>.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar&jan=1&caminho=patrimonio','','width=705,height=<%=tamanho_janela%>,scrollbars=1');">
	<td align="center" valign="top" bgcolor="#C0C0C0">&nbsp;<b><%=zero(num_lista)%></b></td>
		<%if campo_0 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>"   >&nbsp;<%=nm_tipo%><%=prefx_pat(no_patrimonio)%></td>
   		<%end if%>
		<%if campo_1 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>">&nbsp;<%=nm_equipamento%></td>
   		<%end if%>
		<%if campo_2 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>">&nbsp;<%=cd_pn%></td>
   		<%end if%>
		<%if campo_3 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>">&nbsp;<%=cd_ns%></td>
   		<%end if%>
		<%if campo_4 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>">&nbsp;<%=nm_marca%></td>
   		<%end if%>
		<%if campo_5 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>">&nbsp;<%=nm_especialidade%></td>
   		<%end if%>
		<%if campo_6 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>">&nbsp;<%=nm_estado%></td>
   		<%end if%>
		<%if campo_7 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>">&nbsp;<%=nm_rack%></td>
   		<%end if%>
		<%if campo_8 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>">&nbsp;<%=nm_sigla%></td>
   		<%end if%>
		<%if campo_9 <> 0 Then%>
			<td valign="top" style="<%=stilo_linha%>">&nbsp;<%=zero(day(dt_data))&"/"&zero(month(dt_data))&"/"&year(dt_data)%></td>
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
