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
	
	document.form.ordem_inicial.value='1';
	
	document.ordem_1.style.display='inline';
	document.ordem_2.style.display='inline';
	document.ordem_3.style.display='inline';
	document.ordem_4.style.display='inline';
	document.ordem_5.style.display='inline';
	document.ordem_6.style.display='inline';
	document.ordem_7.style.display='inline';
	
	
}
// -->	
</script>
<%
'***** 1.Verifica quais campos foram selecionados *****
cat_patr = request("cat_patr")

campos = request("campos")
	'campo_0 = Instr(1,campos,"no_patrimonio",0)
	campo_1 = Instr(1,campos,"nm_tipo,no_patrimonio",0)
	campo_2 = Instr(1,campos,"nm_equipamento",0)
	campo_3 = Instr(1,campos,"unidade_origem,sigla_origem",0)
	campo_4 = Instr(1,campos,"unidade_uso,sigla_uso",0)
	campo_5 = Instr(1,campos,"a002_datpro",0)
	campo_6 = Instr(1,campos,"a002_numpro",0)


'***** 2.Detalhamento da Busca **************************
'if cat_patr = "BTx05_kt" Then
'	tipo_patr = request("tipo_patr")
'else
'	if cat_patr = "BTx05_kt" then
'		tipo_patr = "0"
'	elseif cat_patr = "pRkct_9f" then
'		tipo_patr = "E"
'	elseif cat_patr = "Zg76$mv" then
'		tipo_patr = "I"
'	elseif cat_patr = "93_BctAA" then
'		tipo_patr = "O"
'	end if
	
	'tipo_patr = cat_patr
'end if
'		if tipo_patr <> "" then
'			tipo_cond = "nm_tipo = '"&tipo_patr&"'"
'			prox_cond = " AND "
'		end if

patrimonio = request("patrimonio")
	if patrimonio <> "" then
	 	patr_cond = prox_cond&"no_patrimonio = '"&patrimonio&"'"
		prox_cond = " AND "
	end if
equipamento = request("equipamento")
eqp_busca = request("eqp_busca")
	if equipamento <> "" then
	 	equip_cond = prox_cond&"nm_equipamento_novo LIKE '"&eqp_busca&""&equipamento&"%'"
		prox_cond = " AND "
	end if
unidade_o = request("unidade_o")
	if unidade_o <> "0" then
	 	unidade_o_cond = prox_cond&"cd_unidade_o LIKE '"&unidade_o&"'"
		prox_cond = " AND "
	end if
unidade_u = request("unidade_u")
	if unidade_u <> "0" then
	 	unidade_u_cond = prox_cond&"cd_unidade_u LIKE '"&unidade_u&"'"
		prox_cond = " AND "
	end if

sel_dia_i = request("sel_dia_i")
sel_mes_i = request("sel_mes_i")
sel_ano_i = request("sel_ano_i")
sel_dia_f = request("sel_dia_f")
sel_mes_f = request("sel_mes_f")
sel_ano_f = request("sel_ano_f")
	cor_aviso = "black"
	if sel_dia_i <> "" OR sel_mes_i <> "" OR sel_dia_f <> "" OR sel_mes_f <> "" then
		'*** Evita datas inválidas ***
		'*** ultimo dia do mes ***
		if int(sel_dia_f) > ultimodiames(sel_mes_f,sel_ano_f) then
			sel_dia_f = ultimodiames(sel_mes_f,sel_ano_f)
		end if
		'*** Data inicial maior que final ***
		if int(sel_dia_i) > int(sel_dia_f) OR int(sel_mes_i) > int(sel_mes_f) OR int(sel_ano_i) > int(sel_ano_f) then
			sel_dia_f = sel_dia_i
			sel_mes_f = sel_mes_i
			sel_ano_f = sel_ano_i
			cor_aviso = "red"
		end if
		
		
		procdata_cond = prox_cond&" a002_datpro between '"&sel_mes_i&"/"&sel_dia_i&"/"&sel_ano_i&"' AND '"&sel_mes_f&"/"&sel_dia_f&"/"&sel_ano_f&" 23:59' "
		'procdata_cond = prox_cond&" a002_datpro between '"&sel_mes_i&"/"&sel_dia_i&"/"&sel_ano_i&"' AND '"&sel_mes_f&"/"&sel_dia_f&"/"&sel_ano_f&"' "
		prox_cond = " AND "
	end if
protocolo = request("protocolo")
	if protocolo <> "" then
		len_proto = len(protocolo)
		
		if len_proto = 11 then
			str_un_proto = mid(protocolo,1,2)
			str_proto_proto = mid(protocolo,4,6)
			proto_cond = prox_cond&" a002_numpro = '"&str_proto_proto&"' AND cd_unidade_u = '"&str_un_proto&"'"
			prox_cond = " AND "
		elseif len_proto <= 6 then
	 		proto_cond = prox_cond&" a002_numpro = '"&protocolo&"' "
			prox_cond = " AND "
		end if 
	end if

'nao_constar = request("nao_constar")
'	if nao_constar < 2 Then
'		constar_cond = prox_cond&"nao_constar = "&nao_constar&""
'		prox_cond = " AND "
'	end if


	'*** CONDIÇÕES ***
	
	if tipo_patr <> "" OR  patrimonio <> "" OR equipamento <> "" OR unidade_o <> "0" OR unidade_u <> "0" OR sel_dia_i <> "" OR sel_mes_i <> "" OR sel_ano_i <> ""  OR sel_dia_f <> ""  OR sel_mes_f <> "" OR sel_ano_f <> "" AND protocolo <> "" Then
		condicao = " WHERE "&patr_cond&" "& equip_cond&" "&unidade_o_cond&" "&unidade_u_cond&" "&procdata_cond&" "&proto_cond&" " 
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
			
'**********************************************
'**************** Ordem ***********************
sentido = request("sentido")

ordem_inicial = request("ordem_inicial")
		ordem_1 = request("ordem_1")
		ordem_2 = request("ordem_2")
		ordem_3 = request("ordem_3")
		ordem_4 = request("ordem_4")
		ordem_5 = request("ordem_5")
		ordem_6 = request("ordem_6")
		ordem_7 = request("ordem_7")
		
ordem_total = request("ordem_total")
	ver_string_ordem = instr(ordem_total,",")
	
		if ver_string_ordem = 1 then
			ordem_total = mid(ordem_total,2,len(ordem_toral))
		end if
	
			if ordem_total <> "" Then
				ordem = "ORDER BY "&ordem_total&" "&sentido 
			end if
			
				ordem_primaria = instr(ordem_total,",")
					if ordem_primaria = 0 then
						ordem_primaria = len(ordem_total)
					end if
						primeiro_campo = trim(mid(ordem_total,1,ordem_primaria))
%>

<body>
<br>
<br><!-- LISTAGEM DO PATRIMONIO -->
<table align="center" class="no_print">
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">protocolo_relatorio_utilizacao_oticas.asp</span></td></tr>
	<%end if%>
	<tr>
		<td><a href="javascript:void(0);" onclick="window.open('patrimonio/patrimonio_cad.asp','nome','width=700','height=50'); return false;" >:: Novo ::</a></td>
		<td colspan="2"><b>Listagem de utilização patrimônial</b></td>
	</tr>
		
	<form action="protocolo.asp" name="form" id="Form" method="post">
	<input type="hidden" name="tipo" value="patrimonio">
	<input type="hidden" name="cat_patr" value="<%=cat_patr%>">
	<!--tr>
		<td colspan="100%" align="center">BUSCA AVANÇADA
		<%if tipo_patr = "E" Then%> - EQUIPAMENTOS
		<%elseif tipo_patr = "I" Then%> - INSTRUMENOS
		<%elseif tipo_patr = "O" Then%> - ÓTICAS
		<%end if%>
		</td>
	</tr-->
	<tr style="background-color:gray; color:white;">
		<td>MOSTRAR</td>
		<td>DETALHE</td>
		<td>ORDEM</td>
	</tr>
	<tr>
		<td>:: <a href="javascript:selecionar_tudo()">Todos</a> :: <a href="javascript:deselecionar_tudo()">Nada</a> ::</td>
		<td>&nbsp;</td>
		<td>
		<input type="hidden" name="ordem_res">
		<input type="hidden" name="ordem_total" value="<%=ordem_total%>">
		<input type="hidden" name="ordem_inicial" value="<%if ordem_inicial = "" Then%>1<%else response.write(ordem_inicial) end if%>">
		&nbsp;<a href="javascript:void();" onClick="limpa_ordem()">Limpa Ordem</a></td>
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
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#c0c0c0');" bgcolor="#c0c0c0">
		<td><input type="checkbox" name="campos" value="" disabled>Conta:</td>		
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_1
			ordem_num="ordem_1"
			ordem_campo="conta"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#c0c0c0;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#c0c0c0');" bgcolor="#c0c0c0">
		<td><input type="checkbox" name="campos" value="nm_tipo,no_patrimonio,cd_patrimonio" <%if campo_1 <> 0 Then response.write("CHECKED") end if%>>Patrimonio:</td>		
		<td><input type="text" name="patrimonio" size="20" maxlength="20" value="<%=patrimonio%>" style="background-color:white;"></td>
		<td><%ordem_var = ordem_2
			ordem_num="ordem_2"
			ordem_campo="nm_tipo,no_patrimonio"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#c0c0c0;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#c0c0c0');" bgcolor="#c0c0c0">
		<td><input type="checkbox" name="campos" value="nm_equipamento_novo" <%if campo_2 <> 0 Then response.write("CHECKED") end if%>>Nome</td>		
		<td><input type="text" name="equipamento" size="20" maxlength="20" value="<%=equipamento%>" style="background-color:white;">
			<input type="radio" name="eqp_busca" value="" <%if eqp_busca = "" Then%>checked<%end if%>>Exato &nbsp; &nbsp; <input type="radio" name="eqp_busca" value="%" <%if eqp_busca = "%" Then%>checked<%end if%>>Qualquer</td>
		<td><%ordem_var = ordem_3
			ordem_num="ordem_3"
			ordem_campo="nm_equipamento_novo"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#c0c0c0;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#c0c0c0');" bgcolor="#c0c0c0">
		<td><input type="checkbox" name="campos" value="unidade_origem,sigla_origem" <%if campo_3 <> 0 Then response.write("CHECKED") end if%>>Unidade origem</td>
		<td><select name="unidade_o" style="background-color:white;">
				<option value="0"></option>
			<%strsql = "SELECT * FROM  TBL_unidades where cd_status = 1 order by nm_unidade"
		  	Set rs_unidade = dbconn.execute(strsql)%>
			<%while not rs_unidade.EOF
				cd_unidade = rs_unidade("cd_codigo")
				nm_unidade = rs_unidade("nm_unidade")%>
				<option value="<%=cd_unidade%>" <%if int(unidade_o) = cd_unidade then%>SELECTED<%end if%>><%=nm_unidade%></option>
			<%rs_unidade.movenext
			wend%>
			</select></td>
		<td><%ordem_var = ordem_4
			ordem_num="ordem_4"
			ordem_campo="sigla_origem"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#c0c0c0;" value="<%=ordem_var%>"></td>
	</tr>
	<tr><td colspan="4"><img src="../../imagens/blackdot.gif" width="500" height="1"></td></tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#e8e8e8');" bgcolor="#e8e8e8">
		<td><input type="checkbox" name="campos" value="unidade_uso,sigla_uso" <%if campo_4 <> 0 Then response.write("CHECKED") end if%>>Unidade uso</td>
		<td><select name="unidade_u" style="background-color:white;">
				<option value=0></option>
			<%strsql = "SELECT * FROM  TBL_unidades where cd_status = 1 and cd_hospital = 1 order by nm_unidade"
		  	Set rs_unidade = dbconn.execute(strsql)%>
			<%while not rs_unidade.EOF
				cd_unidade = rs_unidade("cd_codigo")
				nm_unidade = rs_unidade("nm_unidade")%>
				<option value="<%=cd_unidade%>" <%if int(unidade_u) = cd_unidade then%>SELECTED<%end if%>><%=nm_unidade%></option>
			<%rs_unidade.movenext
			wend%>
			</select></td>
		<td><%ordem_var = ordem_5
			ordem_num="ordem_5"
			ordem_campo="sigla_uso"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#e8e8e8;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#e8e8e8');" bgcolor="#e8e8e8">
		<td><input type="checkbox" name="campos" value="a002_datpro" <%if campo_5 <> 0 Then response.write("CHECKED") end if%>>Período</td>
		<td><input type="text" name="sel_dia_i" size="2" maxlength="2" value="<%=sel_dia_i%>" style="background-color:white;" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
			<input type="text" name="sel_mes_i" size="2" maxlength="2" value="<%=sel_mes_i%>" style="background-color:white;" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" >/
			<input type="text" name="sel_ano_i" size="4" maxlength="4" value="<%=sel_ano_i%>" style="background-color:white;" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" >
			&nbsp;Até&nbsp;
			
			<input type="text" name="sel_dia_f" size="2" maxlength="2" value="<%=sel_dia_f%>" style="background-color:white;color=<%=cor_aviso%>;" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" >/
			<input type="text" name="sel_mes_f" size="2" maxlength="2" value="<%=sel_mes_f%>" style="background-color:white;color=<%=cor_aviso%>;" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" >/
			<input type="text" name="sel_ano_f" size="4" maxlength="4" value="<%=sel_ano_f%>" style="background-color:white;color=<%=cor_aviso%>;" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" ></td>
		<td><%ordem_var = ordem_6
			ordem_num="ordem_6"
			ordem_campo="a002_datpro"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#e8e8e8;" value="<%=ordem_var%>"></td>
	</tr>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#e8e8e8');" bgcolor="#e8e8e8">
		<td><input type="checkbox" name="campos" value="a002_numpro,cd_unidade_u" <%if campo_6 <> 0 Then response.write("CHECKED") end if%>>Nº Protocolo</td>
		<td><input type="text" name="protocolo" size="20" maxlength="20" value="<%=protocolo%>" style="background-color:white;"></td>
		<td><%ordem_var = ordem_7
			ordem_num="ordem_7"
			ordem_campo="cd_unidade_u,a002_numpro"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#e8e8e8;" value="<%=ordem_var%>"></td>
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
		<td>&nbsp;</td>
		<td>&nbsp;QTD</td>
	<%if campo_1 <> 0 Then%>
		<td><b>&nbsp;PAT</b></td>
	<%end if%>
	<%if campo_2 <> 0 Then%>
		<td><b>&nbsp;ITEM</b></td>
	<%end if%>
	<%if campo_3 <> 0 Then%>
		<td><b>&nbsp;ORIGEM</b></td>
	<%end if%>
	<%if campo_4 <> 0 Then%>
		<td><b>&nbsp;USO</b></td>
	<%end if%>
	<%if campo_5 <> 0 Then%>
		<td><b>&nbsp;DATA</b></td>
	<%end if%>
	<%'if campo_5 <> 0 Then%>
		<!--td><b>&nbsp;ESPECIALIDADE</b></td-->
	<%'end if%>
	<%if campo_6 <> 0 Then%>
		<td><b>&nbsp;PROTOCOLO</b></td>
	<%end if%>
		<td><b>O.S.</b></td>
</tr>
<tr>
	<td colspan="100%"><img src="../imagens/px.gif" alt="" width="100%" height="1" border="0"></td>
</tr>
<%num_lista = 1
	
	'*****************************************************************************************************************
	'*** Modelo de busca que funciona ***
	'SELECT     COUNT(no_patrimonio) AS conta, no_patrimonio, nm_equipamento, unidade_origem, unidade_uso, A002_numpro
	'FROM         VIEW_protocolo_patrimonio_utilizacao
	'WHERE     (A002_datpro BETWEEN '07/01/2010' AND '07/31/2010')
	'GROUP BY no_patrimonio, nm_equipamento, unidade_origem, unidade_uso, A002_numpro
	'ORDER BY A002_numpro, nm_equipamento, no_patrimonio, conta
	%>
	<!--tr><td colspan="100%">
	<span style="color:red;">SELECT COUNT</span>(no_patrimonio) AS conta,<%=campos%> <br>
	<span style="color:red;">FROM VIEW_protocolo_patrimonio_utilizacao</span> <%=condicao%><br>
	<span style="color:red;">GROUP BY </span><%=campos%> <br>
	<span style="color:red;">ORDER BY </span>no_patrimonio">
	
	</td></tr-->	
<%
'if kkkk = 2001 then
if kkkk <> 2001 then
'************************************************************************************************************
	'*** SQL - BUSCA - *** 
		'strsql = "SELECT COUNT(no_patrimonio) AS conta, "&campos&" FROM VIEW_protocolo_patrimonio_utilizacao "&condicao&" "&ordem&""
		'condicao
		'strsql = "SELECT COUNT(no_patrimonio) AS conta,"&campos&" FROM VIEW_protocolo_patrimonio_utilizacao "&patr_cond&" "&equip_cond&" "&unidade_o_cond&" "&unidade_u_cond&" "&procdata_cond&" GROUP BY "&campos&" order by no_patrimonio"
		
		
		'strsql = "up_protocolo_patrimonio_utilizacao_lista @campos="&campos&", @condicao="&condicao&", @ordem="&ordem&""
		'strsql = "up_protocolo_patrimonio_utilizacao_lista @campos="&campos&",  @ordem="&ordem&""
		'strsql = "up_protocolo_patrimonio_utilizacao_lista @campos='"&campos&"'"
		'Set rs_patrimonio = dbconn.execute(strsql)
		
		strsql = "SELECT COUNT(no_patrimonio) AS conta,"&campos&" FROM VIEW_protocolo_patrimonio_utilizacao "&condicao&" GROUP BY "&campos&" "&ordem&""
		'strsql = "up_protocolo_patrimonio_utilizacao_lista @campos="&campos&", @condicao="&condicao&", @ordem="&ordem&""
		
		Set rs_patrimonio = dbconn.execute(strsql)
		
		'*** Caso não haja nada no BD ***
		if rs_patrimonio.EOF Then%>
		<tr><td colspan="100%">Nada encontrado</td></tr>
		<%else
			Do while not rs_patrimonio.EOF
			'*** Código do Patrimonio - PERMANENTE ***
			'cd_pat_codigo = rs_patrimonio("cd_")
			'*****************************************
			conta = rs_patrimonio("conta")
			'*** Patrimonio ***
			if campo_1 <> 0 Then
				nm_tipo = rs_patrimonio("nm_tipo")
				no_patrimonio = rs_patrimonio("no_patrimonio")
				cd_patrimonio = rs_patrimonio("cd_patrimonio")
			end if
			
			'*** Equipamento ***
			if campo_2 <> 0 Then
				'cd_equipamento = rs_patrimonio("cd_equipamento")
				nm_equipamento = rs_patrimonio("nm_equipamento_novo")
			end if
			
			'*** Unidade de origem ***
			if campo_3 <> 0 Then
				nm_sigla_origem = rs_patrimonio("sigla_origem")
				nm_unidade_origem = rs_patrimonio("unidade_origem")
			end if
			
			'*** Unidade de uso ***
			if campo_4 <> 0 Then
				nm_sigla_uso = rs_patrimonio("sigla_uso")
				nm_unidade_uso = rs_patrimonio("unidade_uso")
			end if
			
			'*** Data do procedimento ***
			if campo_5 <> 0 Then
				a002_datpro = rs_patrimonio("a002_datpro")
					'a002_datpro = zero(day(a002_datpro))&"/"&zero(month(a002_datpro))&"/"&year("a002_datpro")
					a002_datpro = zero(day(a002_datpro))&"/"&zero(month(a002_datpro))&"/"&year(a002_datpro)
			end if
			
			'*** Protocolo ***
			if campo_6 <> 0 Then
				cd_unidade_u = rs_patrimonio("cd_unidade_u")
				a002_numpro = rs_patrimonio("a002_numpro")
					no_protocolo = digitov(zero(cd_unidade_u)&"."&proto(a002_numpro))
				'a002_datpro = digitov("27."&proto(a002_numpro))
			end if
			
			
			%>
<!--tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('patrimonio_cadastro.asp?cd_pat_codigo=<%=cd_pat_codigo%>&cd_equipamento=<%=cd_equipamento%>&cd_marca=<%=cd_marca%>&cd_unidade=<%=cd_unidade%>&cd_ns=<%=cd_ns%>&cd_pn=<%=cd_pn%>&no_patrimonio=<%=no_patrimonio%>&dt_data=<%=dt_data%>&acao=editar');" onmouseout="mOut(this,'');"-->
<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');">
	<td align="center" valign="top" bgcolor="#C0C0C0">&nbsp;<b><%=zero(num_lista)%></b></td>
	<td valign="top" align="right">&nbsp;<%=zero(conta)%></td>
		<%if campo_0 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_tipo%><%=no_patrimonio%></td>
   		<%end if%>
		<%if campo_1 <> 0 Then%>
			<td valign="top">&nbsp;<span style="color:gray;"><%=nm_tipo%></span><%=no_patrimonio%><%'=nm_equipamento%></td>
   		<%end if%>
		<%if campo_2 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_equipamento%></td>
   		<%end if%>
		<%if campo_3 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_sigla_origem%><%'=nm_unidade_origem%></td>
   		<%end if%>
		<%if campo_4 <> 0 Then%>
			<td valign="top">&nbsp;<%=nm_sigla_uso%></td>
   		<%end if%>
		<%if campo_5 <> 0 Then%>
			<td valign="top">&nbsp;<%=a002_datpro%></td>
   		<%end if%>
		<%if campo_6 <> 0 Then%>
			<td valign="top" onclick="window.open('protocolo/protocolo_folha.asp?tipo=ver&cd_unidade=<%=cd_unidade_u%>&cd_protocolo=<%=a002_numpro%>&cd_digito=<%'=%>','visualizar','width=460,height=425,scrollbars=yes'); return false;">&nbsp;<%=no_protocolo%></td>
   		<%end if%>
		<%'*** Abre a janela de verificação ***
		titulo = "Patrimonio"
		campo = "no_patrimonio"
		variavel = no_patrimonio
		'condicao = "WHERE (no_patrimonio = |"&no_patrimonio&"|)"
		condicao = "WHERE (cd_patrimonio = |"&cd_patrimonio&"|)"%>
		<td onclick="window.open('manutencao_2/manutencao_lista_janela2.asp?condicao=<%=condicao%>&variavel=<%=variavel%>&titulo=<%=titulo%>','visualizar','width=460,height=350,scrollbars=yes'); return false;">&nbsp;<%=no_protocolo%>
		<%'*** Verifica a quantidade de manutenção ***
		'no_patrimonio
		strsql = "SELECT COUNT(no_patrimonio) AS conta_manut FROM VIEW_os_lista2 WHERE (cd_patrimonio = '"&cd_patrimonio&"')"
		Set rs_conta_manut = dbconn.execute(strsql)
		
			while not rs_conta_manut.EOF
				response.write(rs_conta_manut("conta_manut"))
			rs_conta_manut.movenext
			wend%>
		</td>		
</tr>
			<%no_protocolo = ""
			num_lista = num_lista + 1
			rs_patrimonio.movenext
			loop
		end if
	end if%>
<tr><td colspan="100%"><%'=campos%></td></tr>
<!--tr><td colspan="100%">SELECT cd_codigo,<%=campos%> <br>FROM View_patrimonio_lista <br><%=condicao%> &nbsp; <br> <%=ordem%></td></tr-->
<tr>
		<td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>	
	<%'if campo_0 <> 0 Then%>
		<!--td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td-->
	<%'end if%>
	<%if campo_1 <> 0 Then%>
	<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
	<%end if%>
	<%if campo_2 <> 0 Then%>
		<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>	
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
		<td class="no_print"><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
	<%end if%>	
</tr>

<%if session("cd_codigo") = 46 then%>
		<tr class="no_print">
			<td colspan="100" align="center">
				Campos:<%=campos%><br>
				Condicao: <%=condicao%><br>
				Ordem: <%=ordem%><br><br>
				
				SELECT COUNT(no_patrimonio) AS conta,<%=campos%><br>
				FROM VIEW_protocolo_patrimonio_utilizacao <%=condicao%> <br>
				GROUP BY <%=campos%> <br>
				<%=ordem%>
			</td>
		</tr>
	<%end if%>
	

</table><br>
<%end if%>




</body>
</html>
