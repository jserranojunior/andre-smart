<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<script language="javascript">
<!-- 1ª etapa
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
	document.form.ordem_6_1.value='';
	document.form.ordem_7.value='';
	document.form.ordem_8.value='';
	document.form.ordem_9.value='';
	document.form.ordem_10.value='';
	document.form.ordem_11.value='';
	document.form.ordem_12.value='';
	document.form.ordem_13.value='';
	document.form.ordem_14.value='';
	document.form.ordem_15.value='';
	document.form.ordem_16.value='';
	document.form.ordem_17.value='';
	
	document.form.ordem_inicial.value='1';
	
	document.ordem_1.style.display='inline';
	document.ordem_2.style.display='inline';
	document.ordem_3.style.display='inline';
	document.ordem_4.style.display='inline';
	document.ordem_5.style.display='inline';
	document.ordem_6.style.display='inline';
	document.ordem_6_1.style.display='inline';
	document.ordem_7.style.display='inline';
	document.ordem_8.style.display='inline';
	document.ordem_9.style.display='inline';
	document.ordem_10.style.display='inline';
	document.ordem_11.style.display='inline';
	document.ordem_12.style.display='inline';
	document.ordem_13.style.display='inline';
	document.ordem_14.style.display='inline';
	document.ordem_15.style.display='inline';
	document.ordem_16.style.display='inline';
	document.ordem_17.style.display='inline';
	
}
// -->	
</script>
<%x = request("x")
ano_atual = Year(now)

cd_user = session("cd_codigo")

'*** DETALHES ***************
questao = request("questao")
campo_periodo = request("campo_periodo")
campos = request("campos")
cd_ns = request("cd_ns")
no_patrimonio = request("no_patrimonio")
no_patrimonio_vazio = request("no_patrimonio_vazio")
nm_patrimonio = request("nm_patrimonio")

cd_equipamento_novo = request("cd_equipamento_novo")
cd_equipamento = request("cd_equipamento")

nm_equipamento_novo = request("nm_equipamento_novo")
nm_equipamento = request("nm_equipamento")

nm_unidade = request("nm_unidade")
cd_unidade_os = request("cd_unidade_os")

cd_fornecedor = request("cd_fornecedor")
cd_especialidade = request("cd_especialidade")
cd_avaliacao = request("cd_avaliacao")
cd_marca = request("cd_marca")
os_pendente = request("os_pendente")
num_os_assist = request("num_os_assist")
cd_motivo = request("cd_motivo")
cd_situacao = request("cd_situacao")
cd_gestao = request("cd_gestao")



'***** Pendentes ******
if os_pendente = "0" Then 'O.S. Pendente
cond_fecha = "AND (fecha BETWEEN 0 AND 0)"
ck_pend1 = "checked"
elseif os_pendente = "1" then ' O.S. Encerrada
cond_fecha = "AND (fecha BETWEEN 1 AND 1)"
ck_pend2 = "checked"
Elseif os_pendente = "" Then ' O.S. Todas
cond_fecha = " AND (fecha BETWEEN '0' AND '1')"
ck_pend3 = "checked"
End if

'***** Verifica quais campos foram selecionados *****
	campo_ = Instr(1,campos,"fecha",0)
	campo_0 = Instr(1,campos,"num_os",0)
	campo_0_1 = Instr(1,campos,"cd_codigo",0)
	campo_2_1 = Instr(1,campos,"cd_ns",0)
	campo_2_2 = Instr(1,campos,"no_patrimonio",0)
	campo_2_3 = Instr(1,campos,"nm_patrimonio",0)
	campo_2_4 = Instr(1,campos,"no_patrimonio_vazio",0)
	campo_1_1 = Instr(1,campos,"Nm_equipamento",0)
	campo_1 = Instr(1,campos,"nm_equipamento_novo",0)
	'campo_2 = Instr(1,campos,"nm_unidade",0)
	campo_2 = Instr(1,campos,"nm_sigla",0)
	campo_3 = Instr(1,campos,"cd_fornecedor",0)
	campo_4 = Instr(1,campos,"cd_especialidade",0)
	campo_5 = Instr(1,campos,"nm_avaliacao",0)
	campo_6 = Instr(1,campos,"nm_marca",0)
	'if os_pendente = "" then
	'	campo_7 = Instr(1,campos,"dt_os",0)
	'elseif os_pendente = 1 then
		campo_7 = campo_periodo 'Instr(1,campos,"dt_retorno_un",0)
		if campo_periodo = 1 then
			if os_pendente = "" then
				campos = campos &",dt_entrada"
			elseif os_pendente = 0 then
				campos = campos &",dt_entrada"
			elseif os_pendente = 1 then
				campos = campos &",dt_retorno_un"
			end if
		end if
	'end if
	campo_8 = Instr(1,campos,"num_os_assist",0)
	campo_12 = Instr(1,campos,"cd_valor_orc",0)
	campo_13 = Instr(1,campos,"nm_natureza_defeito",0)
	campo_14 = Instr(1,campos,"nm_situacao",0)
	'campo_15 = Instr(1,campos,"",0)
	campo_16 = Instr(1,campos,"cd_gestao,nm_gestao,cd_gestao_ordem",0)

'***** Tratamento da ordem de listagem ******
'*** 2ª etapa ***
ordem_total = request("ordem_total")
	ver_string_ordem = instr(ordem_total,",")
		if ver_string_ordem = 1 then
			ordem_total = mid(ordem_total,2,len(ordem_toral))
		end if
	
			if ordem_total <> "" Then
				ordem = "ORDER BY "&ordem_total&" "&sentido 
			end if
	
ordem_inicial = request("ordem_inicial")
		ordem_1 = request("ordem_1")
		ordem_2 = request("ordem_2")
		ordem_3 = request("ordem_3")
		ordem_4 = request("ordem_4")
		ordem_5 = request("ordem_5")
		ordem_6 = request("ordem_6")
		ordem_6_1 = request("ordem_6_1")
		ordem_7 = request("ordem_7")
		ordem_8 = request("ordem_8")
		ordem_9 = request("ordem_9")
		ordem_10 = request("ordem_10")
		ordem_11 = request("ordem_11")
		ordem_12 = request("ordem_12")
		ordem_13 = request("ordem_13")
		ordem_14 = request("ordem_14")
		ordem_15 = request("ordem_15")
		ordem_16 = request("ordem_16")
		ordem_17 = request("ordem_17")
		
sentido = request("sentido")
'if sentido = "DESC" Then
'ordem_mes = 0
'ordem_ano = 0
'Else
'end if

if campo_0 = 1 Then 'OR campo_2_2 = 1 Then
	codigo = "cd_codigo"
else
	codigo = "sequencia"
end if%>

<%mes_hoje = month(now)
ano_hoje = year(now)

mes_i = request("mes_i")
ano_i = request("ano_i")
dia_i = request("dia_i")
	if dia_i="" OR mes_i = "" OR ano_i = "" Then
		dia_i="1"
		mes_i = mes_hoje
		ano_i = ano_hoje
	end if
data_i = mes_i&"/"&dia_i&"/"&ano_i

mes_f = request("mes_f")
ano_f = request("ano_f")
dia_f = request("dia_f")

	ver_ultimo_dia = UltimoDiaMes(mes_f, ano_f) 
	if int(ver_ultimo_dia) < int(dia_f) then '*** verifica se o ultimo dia é válido.
	dia_f = UltimoDiaMes(mes_f, ano_f)
		alerta = ("O dia da data final selecionado foi substituído automaticamente pelo último dia válido do período.<br>")
	end if

	if dia_f = "" Then '*** Seleciona o ultimo dia do mes atual ao acessar o formulário vazio.
		dia_f = UltimoDiaMes(mes_hoje, ano_hoje)
		alerta = ("ultimo dia automático")
	end if	
	
data_f = mes_f&"/"&dia_f&"/"&ano_f


select_option = "num_qtd"

if no_patrimonio_vazio = "1" Then
'cond_pat_vazio = " AND (no_patrimonio IS not null)"
''cond_pat_vazio = " AND (no_patrimonio <> '')"
cond_pat_vazio = " AND (no_pat_old IS not null)"
end if

if questao = "quantidade" then

	selecao = "SELECT COUNT('"&select_option&"') as conta, "&codigo&", "&campos&""
	cond = " WHERE sequencia = '1' "
	cond_pat = " AND (no_patrimonio = '"&no_patrimonio&"')"
	cond_patrimonio = " AND (no_patrimonio = '"&nm_patrimonio&"')"
	cond_pat_vazio = cond_pat_vazio
	
	if cd_equipamento = "0" AND cd_equipamento_novo <> "0" Then
		'cond_eqp = " AND (cd_equipamento = '"&cd_equipamento_novo&"')"
		cond_eqp = ""
	elseif cd_equipamento <> "0" AND cd_equipamento_novo = "0" Then
		cond_eqp = " AND (cd_equipamento = '"&cd_equipamento&"')"
	elseif cd_equipamento <> "0" AND cd_equipamento_novo <> "0" Then
		cond_eqp = " AND (cd_equipamento = '"&cd_equipamento&"')"
	end if
	'cond_eqp = " AND (cd_equipamento = '"&cd_equipamento_novo&"')"
	'cond_eqp_1 = " AND (cd_equipamento = '"&cd_equipamento&"')"
	
	
	'cond_und = " AND (nm_unidade = '"&nm_unidade&"')"
	cond_und = " AND (cd_unidade_os = '"&cd_unidade_os&"')"
	
	if	os_pendente = "" Then
		'cond_periodo = " AND (dt_os BETWEEN '"&data_i&"' AND '"&data_f&"')"
		cond_periodo = " AND (dt_entrada BETWEEN '"&data_i&"' AND '"&data_f&"')"
		tipo_data = "abert."
	elseif	os_pendente = 0 Then
		'cond_periodo = " AND (dt_os BETWEEN '"&data_i&"' AND '"&data_f&"')"
		cond_periodo = " AND (dt_entrada BETWEEN '"&data_i&"' AND '"&data_f&"')"
		tipo_data = "abert."
	elseif os_pendente = 1 then
		cond_periodo = " AND (dt_retorno_un BETWEEN '"&data_i&"' AND '"&data_f&"')"
		tipo_data = "fecha."
	end if
	
	cond_forn = " AND (cd_fornecedor = '"&cd_fornecedor&"')"
	cond_esp = " AND (cd_especialidade = '"&cd_especialidade&"')"
	cond_aval = " AND (cd_avaliacao = '"&cd_avaliacao&"')"
	cond_marca = " AND (cd_marca = '"&cd_marca&"')"
	cond_motivo = " AND (nm_natureza_defeito = '"&cd_motivo&"')"
	cond_situacao = " AND (cd_situacao = '"&cd_situacao&"')"
	cond_gestao = " AND (cd_gestao = '"&cd_gestao&"')"
	cond_os_assist = " AND num_os_assist like '"&num_os_assist&"'"
	
	
	'*** Agrupamento ***
	'agrupamento = " GROUP BY sequencia,cd_unidade, fecha, "&campos&", "&codigo&""', cd_codigo"
	agrupamento = " GROUP BY sequencia, "&campos&", "&codigo&""', cd_codigo"

Else
	selecao = "SELECT * FROM VIEW_os_lista2"
End if

'***** Substitui a vírgula da sentença de busca *****
	busca_equip = Instr(1,nm_equipamento_novo,",",0)
	if busca_equip <> 0 Then
		nm_equipamento_novo = replace(nm_equipamento_novo,",","%') OR (nm_equipamento_novo Like '%")
	End if
	
	busca_equip_1 = Instr(1,Nm_equipamento,",",0)
	if busca_equip_1 <> 0 Then
		Nm_equipamento = replace(Nm_equipamento,",","%') OR (Nm_equipamento Like '%")
	End if
'****************************************************

'If cd_ns = "0" Then
'	cond_ns = ""
'End if

'If cd_equipamento = "0" AND nm_equipamento_novo = "" Then
'	cond_eqp = ""
'ElseIf cd_equipamento = "0" AND nm_equipamento_novo <> "" Then
'	cond_eqp = "AND (nm_equipamento_novo Like '%"&nm_equipamento_novo&"%')"
'End if

'	If cd_equipamento = "0" AND Nm_equipamento = "" Then
'		cond_eqp_1 = ""
'	ElseIf cd_equipamento = "0" AND Nm_equipamento <> "" Then
'		cond_eqp_1 = "AND (Nm_equipamento Like '%"&Nm_equipamento&"%')"
'	End if

		If cd_equipamento = "0" AND Nm_equipamento = "" Then
			
			cond_eqp_1 = ""
		ElseIf cd_equipamento = "0" AND Nm_equipamento <> "" Then
			cond_eqp_1 = "AND (Nm_equipamento Like '%"&Nm_equipamento&"%')"
		End if

If campo_2_1 = "0" Then
	cond_cd = ""
End if

if nm_patrimonio = "" then 'if campo_2_3 = "0" Then
	cond_patrimonio = ""
Elseif nm_patrimonio <> "" Then
	'cond_patrimonio = "AND (no_patrimonio Like '%"&nm_patrimonio&"%')"
	cond_patrimonio = "AND (no_patrimonio = '"&nm_patrimonio&"')"
End if

if num_os_assist = "" Then
	cond_os_assist = ""
end if


'If nm_unidade = "0" Then
If cd_unidade_os = "0" Then
	cond_und = ""
End if

If cd_fornecedor = "0" then
	cond_forn = "" 
End if
If cd_especialidade = "0" then
	cond_esp = ""
End if
If cd_avaliacao = "0" then
	cond_aval = ""
End if
If cd_marca = "0" then
	cond_marca = ""
End if
If cd_motivo = "0" then
	cond_motivo = ""
End if
If cd_situacao = "0" then
	cond_situacao = ""
End if
If cd_gestao = "0" then
	cond_gestao = ""
End if
%>
<body>
<br>
<table align="center" border="1">
	<tr>
		<td align="center" bgcolor="#C0C0C0"><b>Buscas avançadas <%'=no_patrimonio_vazio%><span style="color:red;"> Manutenção / Compras</span></b>
		<!--<br>M-<%=Instr(1,"cd_equipamento,Nm_equipamento_novo","nm_equipamento",0)%><br>
			P-<%=Instr(1,"cd_equipamento,nm_equipamento_novo","nm_equipamento_novo",0)%><br>
			C-<%=campos%>--></td>
	</tr>
	<tr><td colspan="2"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manutencao2.asp" name="busca" id="busca">
					<input type="hidden" name="tipo" value="nova2">
				<tr><td colspan="2"><b>NOVA O.S.: </b><input type="text" name="cd_unidade2" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"><input type="text" name="num_os2" size="10" maxlength="10" onBlur="submit();"><input type="submit" value="ok"></td></tr>
					</form>
		<tr><td align="center">		
			<table border="1" bordercolor="#C0C0C0" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" align="center">
				<form action="manutencao2.asp" method="post" name="form">
				<input type="hidden" name="tipo" value="relatorio">
				<!-- 3ª Etapa -->
				<input type="hidden" name="ordem_res">
				<input type="hidden" name="ordem_total" value="<%=ordem_total%>">
				<input type="hidden" name="ordem_inicial" value="<%if ordem_inicial = "" Then%>1<%else response.write(ordem_inicial) end if%>">
				<tr id="no_print">
					<td colspan="3"><img src="imagens/px.gif" width="140" height="1"></td>
					<td colspan="4"><img src="imagens/px.gif" width="485" height="1"></td>
					<td><img src="imagens/px.gif" width="115" height="1"></td>
				</tr>
				<tr id="no_print">
					<td colspan="3">&nbsp;<b>MOSTRAR</b></td>
					<td colspan="4">&nbsp;<b>DETALHES</b></td>
					<td>&nbsp;<b>ORDEM</b></td>
				</tr>
				<tr>
					<td colspan="7">&nbsp;</td>
					<td><a href="javascript:void();" onClick="limpa_ordem()">Limpa Ordem</a></td>
				</tr>
				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3">&nbsp;<a href="javascript:selecionar_tudo()">tudo</a> :: <a href="javascript:deselecionar_tudo()">nada</a>
					</td>
					<td colspan="4">&nbsp;<b>QUANTIDADE</b>
						<input type="hidden" name="questao" value="quantidade">
						<!--select name="questao">
						<%'if questao="quantidade" Then%><%qtd = "Selected"%>
						<%'Elseif questao = "listagem" Then%><%list = "selected"%>
						<%'end if%>
							<option value="quantidade" <%=qtd%>>Quantidade</option>
							<option value="listagem" <%=list%>>Listagem</option>
						</select-->
					</td>
					<td>
					<%ordem_var = ordem_1
					ordem_num="ordem_1"
					ordem_campo="conta"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_ <> 0 then%><%ck_campo_ ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="fecha" <%=ck_campo_%>>Status</td>
					<td colspan="4"><input type="radio" name="os_pendente" value="0" <%=ck_pend1%>>&nbsp;Pendentes <b>(Abertas)</b> &nbsp; &nbsp; &nbsp;
									<input type="radio" name="os_pendente" value="1" <%=ck_pend2%>>&nbsp;Encerradas <b>(Fechadas)</b> &nbsp; &nbsp; &nbsp;
									<input type="radio" name="os_pendente" value="" <%=ck_pend3%>>&nbsp;Todas
					</td>
					<td>
					<%ordem_var = ordem_2
					ordem_num="ordem_2"
					ordem_campo="fecha"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>e2e2e2">
					<td colspan="3"><%if campo_0 <> 0 then%><%ck_campo_0 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_unidade,num_os,cd_codigo" <%=ck_campo_0%>>Ordem de serviço
					</td>
					<td colspan="4"></td>
					<td>
					<%ordem_var = ordem_3
					ordem_num="ordem_3"
					ordem_campo="num_os"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>				
				<%bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_2_1 <> 0 then%><%ck_campo_2_1 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_ns" <%=ck_campo_2_1%>>Número de série
					</td>
					<td colspan="4">&nbsp;</td>		
					<td>
					<%ordem_var = ordem_4
					ordem_num="ordem_4"
					ordem_campo="cd_ns"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_2_2 <> 0 then%><%ck_campo_2_2 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="no_pat_old,no_patrimonio" <%=ck_campo_2_2%>>Patrimônio
					</td>
					<%if no_patrimonio_vazio = 1 then%><%patrimonio_vazio = "checked"%><%else%><%patrimonio_vazio2 = "checked"%><%end if%>
					<td colspan="4"><input type="text" name="nm_patrimonio" value="" size="20">
					<input type="radio" name="no_patrimonio_vazio" value="1" <%=patrimonio_vazio%>>	(Não Vazios)  &nbsp;
					<input type="radio" name="no_patrimonio_vazio" value="0" <%=patrimonio_vazio2%>> (vazios)</td>		
					<td>
					<%ordem_var = ordem_5
					ordem_num="ordem_5"
					'ordem_campo="no_pat_old,no_patrimonio"
					ordem_campo="no_pat_old"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
<!-- 1 ------------------------------------------------------------------------------------------------------------------------------------------- -->				
				
				
				<%bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_1 <> 0 then%><%ck_campo_1 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_equipamento,nm_equipamento_novo" <%=ck_campo_1%>>Equipamentos (lista nova)
					</td>
					<td colspan="4"><%=cd_equipamento_novo%>/<%=cd_equipamento%><br>
						<select name="cd_equipamento_novo" style="background-color:#ff9966;">
							<option value="0">Todos (P)
						<%xsql_eqp = "Select * From TBL_equipamento where cd_status = 1 AND nm_equipamento_novo is not null order by nm_equipamento_novo"
						Set rs_eqp = dbconn.execute(xsql_eqp)
						Do while not rs_eqp.EOF
							cd_equip = rs_eqp("cd_codigo")
							nm_equip = rs_eqp("nm_equipamento_novo")
								if campo_1 <> 0 Then
									if int(cd_equipamento_novo) = cd_equip then
										selected = "selected" 
									end if
								end if%>
							<option value="<%=cd_equip%>" <%=selected%>><%=nm_equip%>
							</option>
						<%rs_eqp.movenext
						selected = ""
						loop%>
						</select>
		<%'***** Substitui o código de busca por vírgula *****
		busca_equip = Instr(1,nm_equipamento_novo,"%') OR (nm_equipamento_novo Like '%",0)
		if busca_equip <> 0 Then
			nm_equipamento_novo = replace(nm_equipamento_novo,"%') OR (nm_equipamento_novo Like '%",",")
		End if
		'****************************************************%>
						 ou <input type="text" name="nm_equipamento_novo" value="<%=nm_equipamento_novo%>" size="20" style="background-color:white;">
					</td>
					<td>
					<%ordem_var = ordem_6
					ordem_num="ordem_6"
					ordem_campo="nm_equipamento_novo"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
<!-- 2------------------------------------------------------------------------------------------------------------------------------------------- -->
			
				<%bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_1_1 <> 0 then%><%ck_campo_1_1 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_equipamento,Nm_equipamento" <%=ck_campo_1_1%>>Equipamentos (lista antiga)
					</td>
					<td colspan="4">
						<select name="cd_equipamento" style="background-color:#99ffff;">
							<option value="0">Todos (M)
						<%xsql_eqp = "Select * From TBL_equipamento where nm_equipamento is not Null order by Nm_equipamento"
						Set rs_eqp = dbconn.execute(xsql_eqp)
						Do while not rs_eqp.EOF
							cd_equip = rs_eqp("cd_codigo")
							nm_equip = rs_eqp("Nm_equipamento")
								if campo_1_1 <> 0 Then
									if int(cd_equipamento) = cd_equip then
										selected = "selected" 
									end if
								end if%>
							<option value="<%=cd_equip%>" <%=selected%>><%=nm_equip%>
							</option>
						<%rs_eqp.movenext
						selected = ""
						loop%>
						</select>
		<%'***** Substitui o código de busca por vírgula *****
		busca_equip_1 = Instr(1,Nm_equipamento,"%') OR (Nm_equipamento Like '%",0)
		if busca_equip_1 <> 0 Then
			Nm_equipamento = replace(Nm_equipamento,"%') OR (Nm_equipamento Like '%",",")
		End if
		'****************************************************%>
						 ou <input type="text" name="nm_equipamento" value="<%=Nm_equipamento%>" size="20" style="background-color:white;">
					</td>
					<td>
					<%ordem_var = ordem_6_1
					ordem_num="ordem_6_1"
					ordem_campo="Nm_equipamento"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
<!-- 3 ------------------------------------------------------------------------------------------------------------------------------------------- -->

				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_2 <> 0 then%><%ck_campo_2 ="checked"%><%end if%>
						<!--input type="checkbox" name="campos" value="nm_unidade" <%=ck_campo_2%>>Unidades-->
						<input type="checkbox" name="campos" value="cd_unidade_os,nm_sigla" <%=ck_campo_2%>>Unidades</td>
					<td colspan="4">
						<!--select name="nm_unidade"-->
						<select name="cd_unidade_os">
							<option value="0">Todos
						<%xsql_und = "Select * From TBL_unidades where cd_status = 1 order by nm_unidade"
						Set rs_und = dbconn.execute(xsql_und)
						Do while not rs_und.EOF
							cd_unid = rs_und("cd_codigo")
							nm_sigla = rs_und("nm_sigla")
							nm_unid = rs_und("nm_unidade")
								if int(cd_unidade_os) = cd_unid then
									selected = "selected" 
								End if%>
							<option value="<%=cd_unid%>" <%=selected%>><%=nm_unid%> - <%=nm_sigla%>-<%=cd_unid%>
							</option>
						<%rs_und.movenext
						selected = ""
						loop
						
						nm_unidade = ""%>
						</select>
					</td>
					<td>
					<%ordem_var = ordem_7
					ordem_num="ordem_7"
					ordem_campo="nm_sigla"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_3 <> 0 then%><%ck_campo_3 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_fornecedor,nm_fornecedor" <%=ck_campo_3%>>A. Técnica
					</td>
					<td colspan="4">
						<select name="cd_fornecedor">
							<option value="0">Todos
						<%xsql_forn = "Select * From TBL_fornecedor order by nm_fornecedor"
						Set rs_forn = dbconn.execute(xsql_forn)
						Do while not rs_forn.EOF
							cd_forn = rs_forn("cd_codigo")
							nm_forn = rs_forn("nm_fornecedor")
								if int(cd_fornecedor) = cd_forn then
									selected = "selected" 
								End if%>
							<option value="<%=cd_forn%>" <%=selected%>><%=nm_forn%>
							</option>
						<%rs_forn.movenext
						selected = ""
						loop%>
						</select>
					</td>
					<td>
					<%ordem_var = ordem_8
					ordem_num="ordem_8"
					ordem_campo="nm_fornecedor"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_4 <> 0 then%><%ck_campo_4 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_especialidade,nm_especialidade" <%=ck_campo_4%>>Especialidades
					</td>
					<td colspan="4">
						<select name="cd_especialidade">
							<option value="0">Todos
						<%xsql_esp = "Select * From TBL_especialidades order by nm_especialidade"
						Set rs_esp = dbconn.execute(xsql_esp)
						Do while not rs_esp.EOF
							cd_esp = rs_esp("cd_codigo")
							nm_esp = rs_esp("nm_especialidade")
								if int(cd_especialidade) = cd_esp then
									selected = "selected" 
								End if
							%>
							<option value="<%=cd_esp%>" <%=selected%>><%=nm_esp%> <%'=cd_esp%>
							</option>
						<%rs_esp.movenext
						selected = ""
						loop%>
						</select>
						</td>
						<td>
						<%ordem_var = ordem_9
					ordem_num="ordem_9"
					ordem_campo="nm_especialidade"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_5 <> 0 then%><%ck_campo_5 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_avaliacao,nm_avaliacao" <%=ck_campo_5%>>Avaliação
					</td>
					<td colspan="4">
						<select name="cd_avaliacao">
							<option value="0">Todos
						<%xsql_aval = "Select * From TBL_avaliacao Order by nm_avaliacao"
						Set rs_aval = dbconn.execute(xsql_aval)
						Do while not rs_aval.EOF
							cd_aval = rs_aval("cd_codigo")
							nm_aval = rs_aval("nm_avaliacao")
								if int(cd_avaliacao) = cd_aval then
									selected = "selected" 
								End if
								%>
							<option value="<%=cd_aval%>" <%=selected%>><%=nm_aval%>
							</option>
						<%rs_aval.movenext
						selected = ""
						loop%>
						</select>
					</td>
					<td>
					<%ordem_var = ordem_10
					ordem_num="ordem_10"
					ordem_campo="nm_avaliacao"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_6 <> 0 then%><%ck_campo_6 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_marca,nm_marca" <%=ck_campo_6%>>Marcas
					</td>
					<td colspan="4">
						<select name="cd_marca">
							<option value="0">Todos
						<%xsql_marca = "Select * From TBL_marca Order by nm_marca"
						Set rs_marca = dbconn.execute(xsql_marca)
						Do while not rs_marca.EOF
							cdmarca = rs_marca("cd_codigo")
							marca = rs_marca("nm_marca")
								if int(cd_marca) = cdmarca then
									selected = "selected" 
								End if%>
							<option value="<%=cdmarca%>" <%=selected%>><%=marca%>
							</option>
						<%rs_marca.movenext
						selected = ""
						loop%>
						</select>
					</td>
					<td>
					<%ordem_var = ordem_11
					ordem_num="ordem_11"
					ordem_campo="nm_marca"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "FFFFFF"%>			
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_8 <> 0 then%><%ck_campo_8 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="num_os_assist" <%=ck_campo_8%>>O.S. Assistência
					</td>
					<td colspan="4"><input type="text" name="num_os_assist" value="<%=num_os_assist%>" style="background-color:white;"></td>
					<td>
					<%ordem_var = ordem_12
					ordem_num="ordem_12"
					ordem_campo="num_os_assist"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
<!-- --------------------------------------------------------------------------------- -->				
				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_12 <> 0 then%><%ck_campo_12 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_valor_orc" <%=ck_campo_12%>>Valores
					</td>
					<td colspan="4">&nbsp;</td>
					<td>
					<%ordem_var = ordem_13
					ordem_num="ordem_13"
					ordem_campo="cd_calor_orc"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_13 <> 0 then%><%ck_campo_13 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="nm_natureza_defeito" <%=ck_campo_13%>>Motivo
					</td>
					<td colspan="4">
					<select name="cd_motivo">
							<option value="0">Todos</option>
						<%xsql_motivo = "Select * From TBL_ordemservico_motivo Order by nm_motivo"
						Set rs_motivo = dbconn.execute(xsql_motivo)
						Do while not rs_motivo.EOF
							cdmotivo = rs_motivo("cd_codigo")
							motivo = rs_motivo("nm_motivo")
								if int(cd_motivo) = cdmotivo then
									selected = "selected"
								End if%>
							<option value="<%=cdmotivo%>" <%=selected%>><%=motivo%>
							</option>
						<%rs_motivo.movenext
						selected = ""
						loop%></td>
					<td>
					<%ordem_var = ordem_14
					ordem_num="ordem_14"
					ordem_campo="nm_natureza_defeito"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_14 <> 0 then%><%ck_campo_14 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="nm_situacao" <%=ck_campo_14%>>Situação
					</td>
					<td colspan="4">
					<select name="cd_situacao">
							<option value="0">Todos</option>
						<%xsql_situacao = "Select * From TBL_situacao Order by nm_situacao"
						Set rs_situacao = dbconn.execute(xsql_situacao)
						Do while not rs_situacao.EOF
							cdsituacao = rs_situacao("cd_codigo")
							situacao = rs_situacao("nm_situacao")
								if int(cd_situacao) = cdsituacao then
									selected = "selected"
								End if%>
							<option value="<%=cdsituacao%>" <%=selected%>><%=situacao%>
							</option>
						<%rs_situacao.movenext
						selected = ""
						loop%></td>
					<td>
					<%ordem_var = ordem_15
					ordem_num="ordem_15"
					ordem_campo="nm_situacao"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_16 <> 0 then%><%ck_campo_16 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_gestao,nm_gestao,cd_gestao_ordem" <%=ck_campo_16%>>Gestão
					</td>
					<td colspan="4">
					<select name="cd_gestao">
							<option value="0">Todos</option>
							<!--option value="">Null</option-->
						<%xsql_gestao = "Select * From TBL_gestao Order by cd_ordem"
						Set rs_gestao = dbconn.execute(xsql_gestao)
						Do while not rs_gestao.EOF
							cdgestao = rs_gestao("cd_codigo")
							gestao = rs_gestao("nm_gestao")
								if int(cd_gestao) = cdgestao then
									selected = "selected"
								End if%>
							<option value="<%=cdgestao%>" <%=selected%>><%=gestao%>
							</option>
						<%rs_gestao.movenext
						selected = ""
						loop%></td>
					<td>
					<%ordem_var = ordem_16
					ordem_num="ordem_16"
					ordem_campo="cd_gestao_ordem"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
<!-- --------------------------------------------------------------------------------- -->				


				<%bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_7 <> 0 then%><%ck_campo_7 ="checked"%><%end if%>
						<!--<input type="checkbox" name="campos" value="dt_os" <%=ck_campo_7%>>Período -->
						<!--<input type="checkbox" name="campos" value="dt_retorno_un" <%=ck_campo_7%>>Período-->
						<input type="checkbox" name="campo_periodo" value="1" <%=ck_campo_7%>>Período
					</td>
					<td colspan="4">
						<select name="dia_i">
							<%numero = 1
							do while NOT numero = 32
							if int(dia_i) = numero Then
							check = "selected"
							end if%>					
							<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<select name="mes_i">
							<%numero = 1
							do while NOT numero = 13
							if int(mes_i) = numero Then
							check = "selected"
							end if%>					
							<option value="<%=numero%>" <%=check%>><%=left(mes_selecionado(numero),3)%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<%if ano_i = "" then%><%ano_i=1900%><%end if%>
						<input type="text" name="ano_i" size="4" maxlength="4" value="<%=ano_i%>">
						até
						<select name="dia_f">
							<%numero = 1
							do while NOT numero = 32
								'if dia_f = "" then
								'teste = "nada"
								'Else
								'teste = "algo"
								'end if
								
								if int(dia_f) = numero Then
								check = "selected"
								'elseif
								end if%>					
							<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<select name="mes_f">
							<%numero = 1
							do while NOT numero = 13
								if mes_f = "" Then
									if mes_hoje = numero Then
									'numero = 12
									check = "selected"
									end if
								Elseif int(mes_f) = numero Then
									check = "selected"
								end if%>		
							<option value="<%=numero%>" <%=check%>><%=left(mes_selecionado(numero),3)%></option>
							<%numero = numero+1
						check = ""
						loop%>							
						</select> /
						<%if ano_f = "" then%><%ano_f=ano_atual%><%end if%> 
						<input type="text" name="ano_f" size="4" maxlength="4" value="<%=ano_f%>"> 
					</td>
					<td>
					<%ordem_var = ordem_17
					ordem_num="ordem_17"
					ordem_campo="dt_entrada"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>				
				<tr id="no_print">
					<td colspan="7">&nbsp;<font color="#ff0000"><%=alerta%> <%'=dia_f%></font></td>
					<td>
					<select name="sentido" size="1"><%if sentido = "DESC" Then%><%desc="selected"%><%end if%>
						<option value="ASC">Crescente</option>
						<option value="DESC" <%=desc%>>Decrescente</option>
					</select>
					</td>
				</tr>		
				<tr id="no_print">
					<td colspan="8" align="center">
						<input type="hidden" name="x" value="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="submit" value="Gerar Relatório">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input onclick="javascript:JsRedefinir('')" type="button" value="Limpar" title="Redefinir"> 
					</td>
				</tr>
			</table>
			<blockquote>	
			<table width="2" align="center" border="0" >
				<!--tr>
					<td colspan="100%">&nbsp;
					</td>
				</tr-->
			</form>
			<%
			'******************************
			'*							  *
			'* 	  1ª Parte - formulário   *
			'*                            *
			'******************************
		If campos = "" Then%>
				<tr>
					<td colspan="100%"><br><%response.write("<b> Nenhum item foi selecionado</b>")%>
					</td>
				</tr>
		<%Else%>
				<tr bgcolor="#c0c0c0">
					<td>&nbsp;Qtd</td>
<%if campo_ <> 0 Then%><td>&nbsp;Status</td><%end if%>
<%if campo_0 <> 0 Then%><td>&nbsp;O.S.</td><%end if%>
<%'if campo_0_1 <> 0 Then%><!--td>&nbsp;Status</td--><%'end if%>
<%if campo_7 <> 0 Then%><td>&nbsp;Data <%=tipo_data%></td><%end if%>
<%if campo_1 <> 0 Then%><td>&nbsp;Equipamento (P)</td><%end if%>
<%if campo_1_1 <> 0 Then%><td>&nbsp;Equipamento (M)</td><%end if%>
<%if campo_2_1 <> 0 Then%><td>&nbsp;N/S</td><%end if%>
<%if campo_2_2 <> 0 Then%><td>&nbsp;Patrimônio</td><%end if%>
<%if campo_2 <> 0 Then%><td>&nbsp;Unidade</td><%end if%>
<%if campo_3 <> 0 Then%><td>&nbsp;A. Tecnica</td><%end if%>
<%if campo_4 <> 0 Then%><td>&nbsp;Especialidades</td><%end if%>
<%if campo_5 <> 0 Then%><td>&nbsp;Avaliacao</td><%end if%>
<%if campo_6 <> 0 Then%><td>&nbsp;Marca</td><%end if%>
<%if campo_8 <> 0 Then%><td>&nbsp;O.S. Assist.</td><%end if%>
<%if campo_12 <> 0 Then%><td>&nbsp;Valor</td><%end if%>
<%if campo_13 <> 0 Then%><td>&nbsp;Motivo</td><%end if%>
<%if campo_14 <> 0 Then%><td>&nbsp;Situação</td><%end if%>
<%if campo_16 <> 0 Then%><td>&nbsp;Gestão</td><%end if%>

				</tr>
				<tr><td colspan="100%" id="ok_print"><img src="imagens/blackdot.gif" alt="" width="100%" height="1"></td></tr>
			<%
			
			'xsql = ""&selecao&" FROM VIEW_os_lista "&cond&" "&cond_mes&" "&cond_ano&" "&cond_ns&" "&cond_patrimonio&" "&cond_pat_vazio&" "&cond_eqp&" "&cond_und&" "&cond_forn&" "&cond_esp&" "&cond_aval&" "&cond_marca&" "&agrupamento&" "&ordem&" "&sentido&""'
			'xsql = ""&selecao&" FROM VIEW_os_lista "&cond&" "&cond_periodo&" "&cond_fechada&" "&cond_ns&" "&cond_patrimonio&" "&cond_pat_vazio&" "&cond_eqp&" "&cond_eqp_1&" "&cond_und&" "&cond_forn&" "&cond_esp&" "&cond_aval&" "&cond_marca&" "&cond_fecha&" "&cond_valor&" "&cond_motivo&" "&cond_situacao&"  "&cond_gestao&"  "&cond_os_assist&" "&agrupamento&" "&ordem&""'
			xsql = ""&selecao&" FROM VIEW_os_lista2 "&cond&" "&cond_periodo&" "&cond_fechada&" "&cond_ns&" "&cond_patrimonio&" "&cond_pat_vazio&" "&cond_eqp&" "&cond_eqp_1&" "&cond_und&" "&cond_forn&" "&cond_esp&" "&cond_aval&" "&cond_marca&" "&cond_fecha&" "&cond_valor&" "&cond_motivo&" "&cond_situacao&"  "&cond_gestao&"  "&cond_os_assist&" "&agrupamento&" "&ordem&""'
			Set rs = dbconn.execute(xsql)
			do while Not rs.EOF 
				if campo_0 >  0 then
					'cd_cod = rs("cd_codigo")
				end if
			
			If questao = "quantidade" then
			conta = rs("conta")
			
			
			Else
			
			'num_qtd = rs("num_qtd")
			'nm_unidade = rs("nm_unidade")
			'nm_fornecedor = rs("nm_fornecedor")
			'nm_marca = rs("nm_marca")
			'nm_especialidade = rs("nm_especialidade")
			'cd_fornecedor = rs("cd_fornecedor")
			'nm_avaliacao = rs("nm_avaliacao")
			'nm_marca = rs("nm_marca")
			
			end if
			
			if campo_ <> 0 Then
				fecha = rs("fecha")
			End if
			
			if campo_0_1 <> 0 then
				cd_cod = rs("cd_codigo")
			end if
			
			if campo_0 <> 0 Then
				
				num_os = rs("num_os")
				cd_unidade = rs("cd_unidade")
			End if
			if campo_2_1 <> 0 Then
				cd_ns = rs("cd_ns")
			End if
			if campo_2_2 <> 0 Then
				no_patrimonio = rs("no_patrimonio")
				no_pat_old = rs("no_pat_old")
			End if
			if campo_1 <> 0 Then
				nm_equipamento_novo = rs("nm_equipamento_novo")
			End if
				if campo_1_1 <> 0 Then
					Nm_equipamento = rs("Nm_equipamento")
				End if
				
			if campo_2 <> 0 Then
				'nm_unidade = rs("nm_unidade")
				nm_unidade = rs("nm_sigla")
				
			end if
			if campo_3 <> 0 Then
				nm_fornecedor = rs("nm_fornecedor")
			end if
			if campo_4 <> 0 Then
				nm_especialidade = rs("nm_especialidade")
			end if
			if campo_5 <> 0 Then
				nm_avaliacao = rs("nm_avaliacao")
			end if
			if campo_6 <> 0 Then
				nm_marca = rs("nm_marca")
			end if
			
			if campo_7 <> 0 Then
				if os_pendente = "" Then
					dt_os = rs("dt_entrada")
				elseif os_pendente = 0 Then
					dt_os = rs("dt_entrada")
				elseif os_pendente = 1 then
					dt_os = rs("dt_retorno_un")
				end if
			end if
			
			if campo_8 <> 0 Then
				num_os_assist = rs("num_os_assist")
			end if
			
			if campo_12 <> 0 Then
				cd_valor_orc = rs("cd_valor_orc")
			end if
			
			if campo_13 <> 0 Then
				nm_natureza_defeito = rs("nm_natureza_defeito")
			end if
			
			if campo_14 <> 0 Then
				nm_situacao = rs("nm_situacao")
			end if
			if campo_16 <> 0 Then
				cd_gestao = rs("cd_gestao")
				nm_gestao = rs("nm_gestao")
				cd_gestao_ordem = rs("cd_gestao_ordem")
			end if
			
			
			dt_os_d = zero(day(dt_os))
			dt_os_m = zero(month(dt_os))
			dt_os_a = year(dt_os)
			
			dt_os = dt_os_d&"/"&dt_os_m&"/"&dt_os_a%>
			
				<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:;" onmouseout="mOut(this,'');">
					<td valign="top">&nbsp;<%=zero(conta)%></td>
<%if campo_ <> 0 Then%><td valign="top">&nbsp;<%if fecha = 0 then%><span style="color:green;">A</span><%elseif fecha = 1 then%><span style="color:red;">F</span><%'else%><!--span style="color:red;">f</span--><%end if%></td><%end if%>
<%if campo_0 <> 0 Then%><td valign="top">
	<%if cd_user = 46 then '*** O.S.***%><a href="javascript:;" onclick="window.open('manutencao_2/manutencao_nova.asp?cd_codigo=<%=cd_cod%>&campo=cd_codigo&visual=1&jan=1', 'principal', 'width=651, height=500,scrollbars=1'); return false;"><%=zero(cd_unidade)%>.<%=manutencao(num_os)%></a>
	<%elseif cd_user = 10 then'*** Andamento ***%><a href="javascript:;" onclick="window.open('manutencao_2/manutencao_andamento2.asp?cd_codigo=<%=cd_cod%>&campo=cd_codigo&visual=1&jan=1', 'principal', 'width=651, height=500,scrollbars=1'); return false;"><%=zero(cd_unidade)%>.<%=manutencao(num_os)%></a>
	<%else%><a href="javascript:;" onclick="window.open('manutencao_2/manutencao_ver_janela2.asp?cd_codigo=<%=cd_cod%>&campo=cd_codigo&visual=1&jan=1', 'principal', 'width=600, height=500,scrollbars=1'); return false;"><%=zero(cd_unidade)%>.<%=manutencao(num_os)%></a>
	<%end if%>
	</td>
<%end if%>
<%if campo_7 <> 0 Then%><td valign="top"><%=dt_os%><%=left(mes_selecionado(dt_mes),3)%></td><%end if%>
<%if campo_1 <> 0 Then%><td valign="top"><%=nm_equipamento_novo%></td><%end if%>
<%if campo_1_1 <> 0 Then%><td valign="top"><%=Nm_equipamento%></td><%end if%>
<%if campo_2_1 <> 0 Then%><td valign="top">&nbsp;<%=cd_ns%></td><%end if%>
<%if campo_2_2 <> 0 Then%><td valign="top">&nbsp;<%=no_patrimonio%><span style="color:silver;">(<%=no_pat_old%>)</span></td><%end if%>
<%if campo_2 <> 0 Then%><td valign="top">&nbsp;<%=nm_unidade%></td><%end if%>
<%if campo_3 <> 0 Then%><td valign="top"><%=nm_fornecedor%></td><%end if%>
<%if campo_4 <> 0 Then%><td valign="top">&nbsp;<%=nm_especialidade%></td><%end if%>
<%if campo_5 <> 0 Then%><td valign="top">&nbsp;<%=nm_avaliacao%></td><%end if%>
<%if campo_6 <> 0 Then%><td valign="top"><%=nm_marca%></td><%end if%>
<%if campo_8 <> 0 Then%><td valign="top"><%=num_os_assist%></td><%end if%>
<%if campo_12 <> 0 Then%><td valign="top"><%=cd_valor_orc%></td><%end if%>
<%if campo_13 <> 0 Then%><td valign="top">
		<%if int(nm_natureza_defeito) = 1 then 
			response.write("Motivo desconhecido")
		elseif int(nm_natureza_defeito) = 2 then 
			response.write("Falha do operador")
		elseif int(nm_natureza_defeito) = 3 then
			response.write("Desgaste por uso")
		elseif int(nm_natureza_defeito) = 4 then
			response.write("Quebra por terceiros")
		elseif int(nm_natureza_defeito) = 5 then
			response.write("Solicitação de compras por terceiros")
		end if%>	
		</td><%end if%>
<%if campo_14 <> 0 Then%><td valign="top"><%=nm_situacao%></td><%end if%>
<%if campo_16 <> 0 Then%>
		<%'*** Abre a janela de verificação ***%>
		<%if nm_unidade = "" Then%><%str_sinal_un = "<>"%><%else%><%str_sinal_un = "="%><%end if%>
		<%titulo = "Gestão"
		condicao = "WHERE (dt_os BETWEEN |"&data_i&"| AND |"&data_f&"|) AND (nm_unidade "&str_sinal_un&" |"&nm_unidade&"|) AND (cd_gestao = "&cd_gestao&") AND fecha = "&fecha&""
		%><td valign="top" onclick="window.open('manutencao/manutencao_lista_janela.asp?condicao=<%=condicao%>&titulo=<%=titulo%>','visualizar','width=460,height=350,scrollbars=yes'); return false;"><%=nm_gestao%></td><%end if%>		
		
<!--td valign="top"><%'=fecha%></td-->
			
			</tr>
			<%nm_gestao = ""
			nm_situacao = ""
			
			rs.movenext
			geral = geral + conta
			loop
			
			'end if%>
			
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="1"></td></tr>					
			<tr bgcolor="#C0C0C0">
					<td colspan="100%">&nbsp;Total: <%=geral%></td>
				</tr>
		<%End if%>
				<tr id="ok_print">
						<td><img src="imagens/px.gif" width="25" height="1"></td>
						<td><img src="imagens/px.gif" width="20" height="1"></td>
<%if campo_0 <> 0 Then%><td><img src="imagens/px.gif" width="30" height="1"></td><%end if%>
<%if campo_7 <> 0 Then%><td><img src="imagens/px.gif" width="65" height="1"></td><%end if%>
<%if campo_1 <> 0 Then%><td><img src="imagens/px.gif" width="140" height="1"></td><%end if%>
<%if campo_1_1 <> 0 Then%><td><img src="imagens/px.gif" width="140" height="1"></td><%end if%>
<%if campo_2_1 <> 0 Then%><td><img src="imagens/px.gif" width="75" height="1"></td><%end if%>
<%if campo_2_2 <> 0 Then%><td><img src="imagens/px.gif" width="75" height="1"></td><%end if%>
<%if campo_2 <> 0 Then%><td><img src="imagens/px.gif" width="50" height="1"></td><%end if%>
<%if campo_3 <> 0 Then%><td><img src="imagens/px.gif" width="80" height="1"></td><%end if%>
<%if campo_4 <> 0 Then%><td><img src="imagens/px.gif" width="80" height="1"></td><%end if%>
<%if campo_5 <> 0 Then%><td><img src="imagens/px.gif" width="65" height="1"></td><%end if%>
<%if campo_6 <> 0 Then%><td><img src="imagens/px.gif" width="50" height="1"></td><%end if%>
				</tr>
			<tr id="no_print">
						<td><img src="imagens/px.gif" width="25" height="1"></td>
<%if campo_ <> 0 Then%><td><img src="imagens/px.gif" width="20" height="1"></td><%end if%>
<%if campo_0 <> 0 Then%><td><img src="imagens/px.gif" width="55" height="3"></td><%end if%>
<%if campo_7 <> 0 Then%><td><img src="imagens/px.gif" width="70" height="4"></td><%end if%>
<%if campo_1 <> 0 Then%><td><img src="imagens/px.gif" width="160" height="5"></td><%end if%>
<%if campo_1_1 <> 0 Then%><td><img src="imagens/px.gif" width="160" height="5"></td><%end if%>
<%if campo_2_1 <> 0 Then%><td><img src="imagens/px.gif" width="90" height="6"></td><%end if%>
<%if campo_2_2 <> 0 Then%><td><img src="imagens/px.gif" width="50" height="7"></td><%end if%>
<%if campo_2 <> 0 Then%><td><img src="imagens/px.gif" width="60" height="8"></td><%end if%>
<%if campo_3 <> 0 Then%><td><img src="imagens/px.gif" width="100" height="9"></td><%end if%>
<%if campo_4 <> 0 Then%><td><img src="imagens/px.gif" width="130" height="10"></td><%end if%>
<%if campo_5 <> 0 Then%><td><img src="imagens/px.gif" width="95" height="11"></td><%end if%>
<%if campo_6 <> 0 Then%><td><img src="imagens/px.gif" width="60" height="12"></td><%end if%>
<%if campo_8 <> 0 Then%><td><img src="imagens/px.gif" width="60" height="13"></td><%end if%>
<%if campo_12 <> 0 Then%><td><img src="imagens/px.gif" width="65" height="14"></td><%end if%>
<%if campo_13 <> 0 Then%><td><img src="imagens/px.gif" width="120" height="15"></td><%end if%>
<%if campo_14 <> 0 Then%><td><img src="imagens/px.gif" width="100" height="16"></td><%end if%>
<!--td><img src="imagens/blackdot.gif" width="100" height="17"></td-->
<%if campo_16 <> 0 Then%><td><img src="imagens/px.gif" width="100" height="18"></td><%end if%>
				</tr>				
			</table>
			<%if session_cd_usuario = 46 Then%>
			<br>
			<table>
				<tr class="txt">
					<td align="left" colspan="100%">
					<b>SELECT COUNT</b>(<%=select_option%>) as conta,<%=campos%><br>
					<b>FROM VIEW_os_lista2 </b><%=cond%> <%=cond_periodo%> <%'=cond_ano%> <%=cond_ns%> <%=cond_eqp%> <%=cond_eqp_1%> <%'=cond_pat%> <%=cond_patrimonio%> <%=cond_und%> <%=cond_forn%> <%=cond_esp%> <%=cond_aval%> <%=cond_marca%> <%=cond_valor%><br>
					<%agrupamento = replace(agrupamento,"GROUP BY","<b>GROUP BY</b>")%>
					<%=agrupamento%><br>
					<%ordem = replace(ordem,"ORDER BY","<b>ORDER BY</b>")%>
					<%=ordem%> <br>
					<b><%=sentido%></b></td>
				</tr>
			</table>
			<%end if%>
		</td><!-- Fim da tabela principal-->
	</tr>
</table></blockquote><br>
<br>

<SCRIPT LANGUAGE="javascript">

function JsRedefinir(cod_a, cod_b, cod_c, cod_d)
{
		document.location.href('relatorio_manut.asp?acao=redefinir');
}

</SCRIPT>




</body>
