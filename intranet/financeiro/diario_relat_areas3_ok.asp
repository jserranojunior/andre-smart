<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/javascripts.asp"-->
<!--#include file="../includes/util.asp"-->
<!--#include file="../css/geral.htm"-->
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
	document.form.ordem_7.value='';
	document.form.ordem_8.value='';
	document.form.ordem_9.value='';
	document.form.ordem_10.value='';
	
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
}

function manipulacao_valor(cd_diario,dt_vencimento_anterior,dt_mes,dt_ano,acao,user){
		shipinfo = document.getElementById('seleciona_periodo');
		hoje = new Date
		mes_hoje = hoje.getMonth()+1;
			if (mes_hoje < 10 ) {
				mes_hoje = '0'+ mes_hoje;
				ano_hoje = hoje.getFullYear();}				
			data_hoje = ano_hoje +''+ (mes_hoje);
		 	data_informada = dt_ano +''+ dt_mes;
			/*if (data_informada<data_hoje){
				window.alert ("O mês selecionado já está fechado.");  return (false);}
			else{	*/
				window.open('financeiro/diario_cad3.asp?cd_diario='+cd_diario+'&dia_sel='+dt_vencimento_anterior+'&mes_sel='+dt_mes+'&ano_sel='+dt_ano+'&acao='+acao+'','Cadastro','width=500,height=280,scrollbars=1');
				/*}*/
	}
// -->	
</script>
<%x = request("x")
ano_atual = Year(now)

cd_user = session("cd_codigo")

'*** DETALHES ***************
campo_periodo = request("campo_periodo")
campos = request("campos")
subtotais = request("subtotais")
'----------------------------
cd_area = request("cd_area")
cd_centrocusto = request("cd_centrocusto")
cd_conta = request("cd_conta")
cd_unidade = request("cd_unidade")
cd_tipo_valor = request("cd_tipo_valor")
cd_fornecedor = request("cd_fornecedor")
cd_modo_pagamento = request("cd_modo_pagamento")
'cd_cheque = request("cd_cheque")
'----------------------------
nm_pagamento = request("nm_pagamento")
nm_descricao = request("nm_descricao")
nm_cheque = request("nm_cheque")

'***** Verifica quais campos foram selecionados *****
	campo_1 = Instr(1,campos,"nm_valor",0)
	campo_2 = Instr(1,campos,"dt_vencimento",0)
	campo_3 = Instr(1,campos,"cd_area,nm_area",0)
	campo_4 = Instr(1,campos,"cd_centrocusto,nm_centrocusto,nm_cc",0)
	campo_5 = Instr(1,campos,"cd_conta,nm_conta",0)
	campo_6 = Instr(1,campos,"cd_unidade,nm_unidade,nm_sigla",0)
	campo_7 = Instr(1,campos,"cd_tipo_valor,nm_tipo,nm_tipo_abrv",0)
	campo_8 = Instr(1,campos,"cd_fornecedor,nm_fornecedor",0)
	campo_9 = Instr(1,campos,"nm_pagamento",0)
	campo_10 = Instr(1,campos,"cd_cheque,nm_cheque,cd_modo_pagamento",0)
	

'***** Tratamento da ordem de listagem ******
'*** 2ª etapa ***
ordem_total = request("ordem_total")
	ver_string_ordem = instr(ordem_total,",")
		if ver_string_ordem = 1 then
			ordem_total = mid(ordem_total,2,len(ordem_toral))
		end if
	
			if ordem_total <> "" Then
				ordem = "ORDER BY "&ordem_total&" "&sentido 
			else
				ordem = "ORDER BY dt_vencimento"
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
									
		'**** Prepara para o subtotais ****
		subtotal = ordem_total
			if subtotal <> "" Then
				pri_virg_subtotal = instr(1,subtotal,",",1)
				if pri_virg_subtotal > 0 then
					pri_subtotal = mid(subtotal,1,pri_virg_subtotal - 1)
				else
					pri_subtotal = mid(subtotal,1,len(subtotal))
				end if
			end if
		'***********************************
							
sentido = request("sentido")%>

<%'**** DATAS (Períodos) ********************
mes_hoje = month(now)
ano_hoje = year(now)
'*** Período do relatório *** 
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
	
	if dia_f <> "" AND ano_f&zero(mes_f)&zero(dia_f) < ano_i&zero(mes_i)&zero(dia_i) then
		mes_f = mes_i
		ano_f = ano_i
		dia_f = dia_i
	end if

	ver_ultimo_dia = UltimoDiaMes(mes_f, ano_f) 
	if int(ver_ultimo_dia) < int(dia_f) then '*** verifica se o ultimo dia é válido.
	dia_f = UltimoDiaMes(mes_f, ano_f)
		alerta = ("O dia da data final selecionado, foi substituído automaticamente pelo último dia válido do período.<br>")
	end if

	if dia_f = "" Then '*** Seleciona o ultimo dia do mes atual ao acessar o formulário vazio.
		dia_f = UltimoDiaMes(mes_hoje, ano_hoje)
		alerta = ("ultimo dia automático")
	end if
data_f = mes_f&"/"&dia_f&"/"&ano_f

'*** Data de pagamento ****
dia_i_p = request("dia_i_p")
dia_f_p = request("dia_f_p")

	data_i_p = mes_i_p&"/"&dia_i_p&"/"&ano_i_p
	data_f_p = mes_f_p&"/"&dia_f_p&"/"&ano_f_p

'***** FOCO DO RELATÓRIO *****************************************
'select_option = "nm_valor"

'***** DETALHAMENTO (FILTRAGEM) ***********************************
pre_selecao = " SELECT sum(nm_valor) as pre_total"
'******************************************************************
selecao = " SELECT cd_codigo,cd_diario,nm_valor, nm_fornecedor, nm_pagamento,dt_vencimento,nm_valor,cd_parcela,cd_qtd_parcelas,"&campos&" "
cond = " WHERE cd_qtd_parcelas >= 1 "	
cond_area = " AND cd_area = "&cd_area&""
cond_centrocusto = " AND cd_centrocusto = "&cd_centrocusto&""
cond_conta = " AND cd_conta = "&cd_conta&""
cond_unidade = " AND cd_unidade = "&cd_unidade&""
cond_tipo_valor = " AND cd_tipo_valor = "&cd_tipo_valor&""
cond_periodo = " AND (dt_vencimento BETWEEN '"&data_i&"' AND '"&data_f&"')"
cond_fornecedor = " AND cd_fornecedor = "&cd_fornecedor&""
cond_pagamento = " AND nm_pagamento like '%"&busca_inteligente(nm_descricao)&"%'"
cond_modo_pagto = " AND cd_modo_pagamento = '"&cd_modo_pagamento&"'"
'cond_cheque = " AND cd_cheque = '"&cd_cheque&"'"
cond_cheque = " AND nm_cheque like '"&nm_cheque&"%'"
cond_datacheque = " AND dt_pagamento between '"&data_i_p&"' AND '"&data_f_p&"'"

'***** AGRUPAMENTO ************************************************
	agrupamento = " GROUP BY cd_codigo,cd_diario,nm_valor,nm_fornecedor,nm_pagamento,dt_vencimento,nm_valor,cd_parcela,cd_qtd_parcelas,"&campos&""'
'****************************************************

If cd_area = "0" Then
	cond_area = ""
End if

If cd_centrocusto = "0" Then
	cond_centrocusto = ""
End if

If cd_conta = "0" Then
	cond_conta = ""
End if

If cd_unidade = "0" Then
	cond_unidade = ""
End if

If cd_tipo_valor = "0" Then
	cond_tipo_valor = ""
End if

if cd_fornecedor = "0" Then
	cond_fornecedor = ""
end if

if nm_descricao = "" Then
	cond_pagamento = ""
end if

'if nm_pagamento = "" Then
'	cond_pagamento = ""
'end if

'if cd_cheque = "0" Then
'	cond_cheque = ""
'end if

if nm_cheque = "" Then
	cond_cheque = ""
end if

if cd_modo_pagamento = "0" Then
	cond_modo_pagto = ""
end if

if dt_i_p = "" OR dt_f_p = "" Then
	cond_datacheque = ""
	dt_i_p = ""
	dt_i_f = ""
end if

%>
<body>
<!--br-->
<table align="center" border="0" style="border-collapse:collapse; border:white solid 1px">
	<tr>
		<td align="center" bgcolor="#C0C0C0" style="font-size:13px;"><b>Relatório de Gestão <%'=no_patrimonio_vazio%></b></td>
	</tr>
	<tr><td colspan="2"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="financeiro.asp" name="busca" id="busca">
						<input type="hidden" name="tipo" value="relat_areas">
					</form>				
				<tr><td align="center">
				<!--tr><td>|<a href="financeiro.asp?tipo=diario3"> Contas a pagar </a>|<a href="financeiro.asp?tipo=emissao_cheques"> Cheques emitidos </a>|<br>&nbsp;</td></tr-->
			<table border="1" bordercolor="#C0C0C0" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" align="center">
				<!--form action="financeiro.asp" method="post" name="form"-->
				<form action="diario_relat_areas3.asp" method="post" name="form">
				<input type="hidden" name="tipo" value="relat_areas">
				<!-- 3ª Etapa -->
				<input type="hidden" name="ordem_res">
				<input type="hidden" name="ordem_total" value="<%=ordem_total%>">
				<input type="hidden" name="ordem_inicial" value="<%if ordem_inicial = "" Then%>1<%else response.write(ordem_inicial) end if%>">
				<tr id="no_print" style="background-color:gray; color:white;">
					<td colspan="3">&nbsp;<b>MOSTRAR</b></td>
					<td colspan="4">&nbsp;<b>DETALHES</b></td>
					<td>&nbsp;<b>ORDEM</b></td>
				</tr>
				
				<%bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3">&nbsp;<a href="javascript:selecionar_tudo()">tudo</a> :: <a href="javascript:deselecionar_tudo()">nada</a>
					</td>
					<td colspan="4">&nbsp;</td>
					<td><a href="javascript:void();" onClick="limpa_ordem()">Limpa Ordem</a></td>
				</tr>
				<%bgc_linha = "FFFFFF"%>
				
				<%'***** VALORES ****************
				bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>e2e2e2">
					<td colspan="3"><%if campo_1 <> 0 then%><%ck_campo_1 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="nm_valor" <%=ck_campo_1%> >Valores
					</td>
					<td colspan="4">&nbsp;</td>
					<td>
					<%ordem_var = ordem_1
					ordem_num="ordem_1"
					ordem_campo="nm_valor"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
				<%'**** VENCIMENTO *************
				bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_2 <> 0 then%><%ck_campo_2 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="dt_vencimento"  <%=ck_campo_2%>>Vencimento
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
								if int(dia_f) = numero Then
								check = "selected"
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
					<%ordem_var = ordem_2
					ordem_num="ordem_2"
					ordem_campo="dt_vencimento"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
				<%'***** AREA ******************
				bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>e2e2e2">
					<td colspan="3"><%if campo_3 <> 0 then%><%ck_campo_3 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_area,nm_area" <%=ck_campo_3%>>Área
					</td>
					<td colspan="4">
						<%xsql = "SELECT * FROM TBL_financeiro_area"
						Set rs = dbconn.execute(xsql)%>
						<select name="cd_area">
							<option value="0">todos</option>
							<%while not rs.EOF
								strcd_area = rs("cd_codigo")
								strnm_area = rs("nm_area")%>
								<option value="<%=strcd_area%>" <%if abs(strcd_area) = abs(cd_area) Then response.write("SELECTED")%>><%=strnm_area%></option>
							<%rs.movenext
							wend%>
						</select></td>
					<td>
					<%ordem_var = ordem_3
					ordem_num="ordem_3"
					ordem_campo="nm_area"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
				<%'**** CENTRO DE CUSTO **********
				bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_4 <> 0 then%><%ck_campo_4 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_centrocusto,nm_centrocusto,nm_cc" <%=ck_campo_4%>>Centro de Custo
					</td>
					<td colspan="4"><%xsql = "SELECT * FROM TBL_financeiro_centrocusto"
						Set rs = dbconn.execute(xsql)%>
						<select name="cd_centrocusto">
							<option value="0">todos</option>
							<%while not rs.EOF
								strcd_centrocusto = rs("cd_codigo")
								strnm_centrocusto = rs("nm_centrocusto")
								dtrnm_cc = rs("nm_cc")%>
								<option value="<%=strcd_centrocusto%>" <%if abs(strcd_centrocusto) = abs(cd_centrocusto) Then response.write("SELECTED")%>><%=strnm_centrocusto%></option>
							<%rs.movenext
							wend%>
						</select></td>		
					<td>
					<%ordem_var = ordem_4
					ordem_num="ordem_4"
					ordem_campo="nm_centrocusto"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
				<%'**** CONTA **********
				bgc_linha = "e2e2e2"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_5 <> 0 then%><%ck_campo_5 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_conta,nm_conta" <%=ck_campo_5%>>Conta
					</td>
					<td colspan="4">
					<%'*** modo pagamento ***
					if no_patrimonio_vazio = 1 then%><%patrimonio_vazio = "checked"%><%else%><%patrimonio_vazio2 = "checked"%><%end if%>
					<%xsql = "SELECT * FROM TBL_financeiro_conta"
						Set rs = dbconn.execute(xsql)%>
						<select name="cd_conta">
							<option value="0">todos</option>
							<%while not rs.EOF
								strcd_conta = rs("cd_codigo")
								strnm_conta = rs("nm_conta")
								strnm_conta_abrv = rs("nm_conta_abrv")%>
								<option value="<%=strcd_conta%>" <%if abs(strcd_conta) = abs(cd_conta) Then response.write("SELECTED")%>><%=strnm_conta%></option>
							<%rs.movenext
							wend%>
						</select></td>		
					<td>
					<%ordem_var = ordem_5
					ordem_num="ordem_5"
					ordem_campo="nm_conta"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
				<%'**** UNIDADE **********
				bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_6 <> 0 then%><%ck_campo_6 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_unidade,nm_unidade,nm_sigla" <%=ck_campo_6%>>Unidade
					</td>
					<td colspan="4"><%xsql = "SELECT * FROM TBL_unidades where cd_status = 1"
						Set rs = dbconn.execute(xsql)%>
						<select name="cd_unidade">
							<option value="0">todos</option>
							<%while not rs.EOF
								strcd_unidade = rs("cd_codigo")
								strnm_unidade = rs("nm_unidade")
								strnm_sigla = rs("nm_sigla")%>
								<option value="<%=strcd_unidade%>" <%if abs(strcd_unidade) = abs(cd_unidade) Then response.write("SELECTED")%>><%=strnm_unidade%></option>
							<%rs.movenext
							wend%>
						</select></td>
					<td>
					<%ordem_var = ordem_6
					ordem_num="ordem_6"
					ordem_campo="nm_unidade"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
			
				<%'**** TIPO **********
				bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_7 <> 0 then%><%ck_campo_7 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_tipo_valor,nm_tipo,nm_tipo_abrv" <%=ck_campo_7%>>Tipo
					</td>
					<td colspan="4"><%xsql = "SELECT * FROM TBL_financeiro_tipo_valor"
						Set rs = dbconn.execute(xsql)%>
						<select name="cd_tipo_valor">
							<option value="0">todos</option>
							<%while not rs.EOF
								strcd_tipo = rs("cd_codigo")
								strnm_tipo = rs("nm_tipo")
								strnm_tipo_abrv = rs("nm_tipo_abrv")%>
								<option value="<%=strcd_tipo%>" <%if abs(strcd_tipo) = abs(cd_tipo_valor) Then response.write("SELECTED")%>><%=strnm_tipo%></option>
							<%rs.movenext
							wend%>
						</select></td>
					<td>
					<%ordem_var = ordem_7
					ordem_num="ordem_7"
					ordem_campo="nm_tipo"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
				<%'**** Fornecedores **********
				bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_8 <> 0 then%><%ck_campo_8 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_fornecedor,nm_fornecedor" <%=ck_campo_8%>>Favorecido
					</td>
					<td colspan="4"><%xsql = "SELECT * FROM TBL_fornecedor where cd_status = 1 order by nm_fornecedor"
						Set rs = dbconn.execute(xsql)%>
						<select name="cd_fornecedor">
							<option value="0">todos</option>
							<%while not rs.EOF
								strcd_fornecedor = rs("cd_codigo")
								strnm_fornecedor = rs("nm_fornecedor")%>
								<option value="<%=strcd_fornecedor%>" <%if abs(strcd_fornecedor) = abs(cd_fornecedor) Then response.write("SELECTED")%>><%=strnm_fornecedor%></option>
							<%rs.movenext
							wend%>
						</select></td>
					<td>
					<%ordem_var = ordem_8
					ordem_num="ordem_8"
					ordem_campo="nm_fornecedor"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
				<%'**** Histórico (pagamento)**********
				bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_9 <> 0 then%><%ck_campo_9 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="nm_pagamento" <%=ck_campo_9%>>Pagamento(histórico)
					</td>
					<td colspan="4"><input type="text" name="nm_descricao" value="<%=nm_descricao%>" size="50" maxlength="500"></td>
					<td>
					<%ordem_var = ordem_9
					ordem_num="ordem_9"
					ordem_campo="nm_pagamento"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
				<%'**** PAGAMENTO **********
				bgc_linha = "FFFFFF"%>
				<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
					<td colspan="3"><%if campo_10 <> 0 then%><%ck_campo_10 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_cheque,nm_cheque,cd_modo_pagamento,dt_pagamento" <%=ck_campo_10%>>Pagamento (cheque)
					</td>
					<td colspan="4"><%xsql = "SELECT * FROM TBL_financeiro_banco_pagamentos3 where cd_modo_pagamento = 1"%>
						<select name="cd_modo_pagamento">
							<option value="0">todos</option>
								
							<option value="1" <%if cd_modo_pagamento = 1 then response.write("selected")%>>Cheque</option>
							<option value="2" <%if cd_modo_pagamento = 2 then response.write("selected")%>>Transferência</option>
						</select>
						
						<%Set rs = dbconn.execute(xsql)%>
						<!--select name="cd_cheque">
							<option value="0">Todos</option>
							<%while not rs.EOF
								strcd_cheque = rs("cd_codigo")
								strnm_cheque = rs("nm_cheque")
								strcd_modo_pagamento = rs("cd_modo_pagamento")
								'strcd_conta_banco = rs("cd_conta_banco")
								'strdt_pagamento = rs("dt_pagamento")%>
								<option value="<%=strcd_cheque%>" <%if abs(strcd_cheque) = abs(cd_cheque) Then response.write("SELECTED")%>><%=strnm_cheque%></option>
							<%rs.movenext
							wend%>
						</select-->
						&nbsp; N° Cheque: <input type="text" name="nm_cheque" value="<%=nm_cheque%>" size="10" maxlength="500">
						
						efetuados entre: <select name="dia_i_p">
							<option value="0"></option>
							<%numero = 1
							do while NOT numero = 32
							if int(dia_i_p) = numero Then
							check = "selected"
							end if%>					
							<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>
						e
						<select name="dia_f_p">
							<option value="0"></option>
							<%numero = 1
							do while NOT numero = 32
								if int(dia_f_p) = numero Then
								check = "selected"
								end if%>					
							<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>
						
						
						</td>
					<td>
					<%ordem_var = ordem_10
					ordem_num="ordem_10"
					ordem_campo="cd_cheque"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
				</tr>
				
				<tr bgcolor="#eaeaea"><%if subtotais > 0 then%><%ck_subtotais = "checked"%><%end if%>
					<td colspan="3"><input type="checkbox" name="subtotais" value="1" <%=ck_subtotais%>>&nbsp;Subtotais.</td>
					<td colspan="6" style="face_size:7px; color:silver;" valign="baseline">&nbsp; Obs.: Para obter os subtotais desejados, ordene a busca pela coluna à direita.</td>					
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
					<td colspan="3"><img src="imagens/px.gif" width="140" height="1"></td>
					<td colspan="4"><img src="imagens/px.gif" width="485" height="1"></td>
					<td><img src="imagens/px.gif" width="115" height="1"></td>
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
					<td>&nbsp;</td>
<%if campo_3 <> 0 Then%><td>&nbsp;Área</td><%end if%>
<%if campo_4 <> 0 Then%><td>&nbsp;C.Custo</td><%end if%>
<%if campo_5 <> 0 Then%><td>&nbsp;Conta</td><%end if%>
<%if campo_6 <> 0 Then%><td>&nbsp;Unidade</td><%end if%>
<%if campo_7 <> 0 Then%><td>&nbsp;Tipo</td><%end if%>
<%if campo_8 <> 0 Then%><td>&nbsp;Favorecido - Histórico</td><%end if%>
<%if campo_10 <> 0 Then%><td>&nbsp;Pagto.</td><%end if%>
<%if campo_2 <> 0 Then%><td>&nbsp;Vencimento</td><%end if%>
<%if campo_1 <> 0 Then%><td align="center">&nbsp;Valor</td><%end if%>



				</tr>
				<tr><td colspan="100%" id="ok_print"><img src="imagens/blackdot.gif" alt="" width="100%" height="1"></td></tr>
<%x = "10"
IF x = "10" then

			'espacos = 1
			num_linha = 1
			cor = 1
			nr = 0
			
			'xsql = ""&selecao&" FROM VIEW_os_lista2 "&cond&" "&cond_periodo&" "&cond_fechada&" "&cond_ns&" "&cond_patrimonio&" "&cond_pat_vazio&" "&cond_eqp&" "&cond_eqp_1&" "&cond_und&" "&cond_forn&" "&cond_esp&" "&cond_aval&" "&cond_marca&" "&cond_fecha&" "&cond_valor&" "&cond_motivo&" "&cond_situacao&"  "&cond_gestao&"  "&cond_os_assist&" "&agrupamento&" "&ordem&""'
			'xsql = ""&selecao&" FROM VIEW_financeiro_valores3 "&cond&" "&cond_periodo&" "&cond_fechada&" "&cond_ns&" "&cond_patrimonio&" "&cond_pat_vazio&" "&cond_eqp&" "&cond_eqp_1&" "&cond_und&" "&cond_forn&" "&cond_esp&" "&cond_aval&" "&cond_marca&" "&cond_fecha&" "&cond_valor&" "&cond_motivo&" "&cond_situacao&"  "&cond_gestao&"  "&cond_os_assist&" "&agrupamento&""' "&ordem&""'
			'xsql = ""&selecao&" FROM VIEW_financeiro_valores3 "&cond&" "&cond_area&" "&cond_centrocusto&" "&cond_conta&" "&cond_unidade&" "&cond_tipo_valor&" "&cond_periodo&" "&cond_fornecedor&" "&cond_pagamento&" "&agrupamento&" "&ordem&""' "&cond_ns&" "&cond_patrimonio&" "&cond_pat_vazio&" "&cond_eqp&" "&cond_eqp_1&" "&cond_und&" "&cond_forn&" "&cond_esp&" "&cond_aval&" "&cond_marca&" "&cond_fecha&" "&cond_valor&" "&cond_motivo&" "&cond_situacao&"  "&cond_gestao&"  "&cond_os_assist&" "&agrupamento&""' "&ordem&""'
			'xsql = ""&pre_selecao&" FROM VIEW_financeiro_valores3 "&cond&" "&cond_area&" "&cond_centrocusto&" "&cond_conta&" "&cond_unidade&" "&cond_tipo_valor&" "&cond_periodo&" "&cond_fornecedor&" "&cond_pagamento&" "&cond_modo_pagto&" "&cond_cheque&" "&cond_datacheque&" "&agrupamento&" "&ordem&""'
			xsql = ""&pre_selecao&" FROM VIEW_financeiro_valores3 "&cond&" "&cond_area&" "&cond_centrocusto&" "&cond_conta&" "&cond_unidade&" "&cond_tipo_valor&" "&cond_periodo&" "&cond_fornecedor&" "&cond_pagamento&" "&cond_modo_pagto&" "&cond_cheque&" "&cond_datacheque&" "'
			Set rs = dbconn.execute(xsql)
			while not rs.EOF
				pre_total = rs("pre_total")
				'pre_total = rs(campo_pre_selecao)
			rs.movenext
			wend
			
			
			xsql = ""&selecao&" FROM VIEW_financeiro_valores3 "&cond&" "&cond_area&" "&cond_centrocusto&" "&cond_conta&" "&cond_unidade&" "&cond_tipo_valor&" "&cond_periodo&" "&cond_fornecedor&" "&cond_pagamento&" "&cond_modo_pagto&" "&cond_cheque&" "&cond_datacheque&" "&agrupamento&" "&ordem&""'
			Set rs = dbconn.execute(xsql)
			do while Not rs.EOF 
				cd_codigo = rs("cd_codigo")
				cd_diario = rs("cd_diario")
			
			'if campo_1 <> 0 Then
				nm_valor = rs("nm_valor")
				'sub_total_valor = sub_total_valor + nm_valor
				total_valor = total_valor + nm_valor
			'End if
			
			'if campo_2 <> 0 Then
				dt_vencimento = rs("dt_vencimento")
					dt_vencimento = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&right(year(dt_vencimento),2)
			'end if
				
			if campo_3 <> 0 Then
				cd_area = rs("cd_area")
				nm_area = rs("nm_area")
					espacos = espacos + 1
			end if
			
			if campo_4 <> 0 Then
				cd_centrocusto = rs("cd_centrocusto")
				nm_centrocusto = rs("nm_centrocusto")
				nm_cc = rs("nm_cc")
					espacos = espacos + 1
			end if
			
			if campo_5 <> 0 Then
				cd_conta = rs("cd_conta")
				nm_conta = rs("nm_conta")
					espacos = espacos + 1
			end if
			
			if campo_6 <> 0 Then
				cd_unidade = rs("cd_unidade")
				nm_unidade = rs("nm_unidade")
				nm_sigla = rs("nm_sigla")
					espacos = espacos + 1
			end if
			
			if campo_7 <> 0 Then
				cd_tipo_valor = rs("cd_tipo_valor")
				nm_tipo = rs("nm_tipo")
				nm_tipo_abrv = rs("nm_tipo_abrv")
					espacos = espacos + 1
			end if
			
			if campo_8 <> 0 Then
				cd_fornecedor = rs("cd_fornecedor")
				nm_fornecedor = rs("nm_fornecedor")
					espacos = espacos + 1
			end if
			
			if campo_10 <> 0 Then
				cd_cheque = rs("cd_cheque")
				nm_cheque = rs("nm_cheque")
				cd_modo_pagamento = rs("cd_modo_pagamento")
					if cd_modo_pagamento = 1 then
						nm_modo_pagamento = "ch"
					elseif cd_modo_pagamento then
						nm_modo_pagamento = "transf"
					else
						nm_modo_pagamento = ""
					end if
						if isNull(cd_modo_pagamento) Then
							cd_modo_pagamento = 0
						end if
					espacos = espacos + 1
				dt_pagamento = rs("dt_pagamento")
					dt_pagamento_red = day(dt_pagamento)&"/"&mes_selecionado(month(dt_pagamento))
			end if
			
			
			if pri_subtotal <> "" Then
				compara1 = rs(pri_subtotal)
				if isNULL(compara1) Then
					compara1 = "nada"
				end if
			end if
			
			nm_fornecedor = rs("nm_fornecedor")
			nm_pagamento = rs("nm_pagamento")
			cd_qtd_parcelas = rs("cd_qtd_parcelas")
			cd_parcela = rs("cd_parcela")
			'			
			
			dt_os_d = zero(day(dt_os))
			dt_os_m = zero(month(dt_os))
			dt_os_a = year(dt_os)
			
			dt_os = dt_os_d&"/"&dt_os_m&"/"&dt_os_a%>
			
				<%'*** Exibe o subtotal depois da ultima linha (1) ***
				if compara1 <> grava_info AND sub_total_valor > 0 AND subtotais <> "" then%>
					<tr style="background-color:#b5b5b5;">
						<td colspan="8">&nbsp;</td>
						<td align="center"><%if sub_total_valor <> "" Then 
								'porcentagem_valor_total = sub_total_valor/pre_total
								response.write("<b>"&formatNumber(sub_total_valor/pre_total*100,2)&"%</b>")
							end if%></td>
						<td align="right"><b><i><%if sub_total_valor <> "" Then response.write(formatNumber(sub_total_valor,2))%></i></b></td>
					</tr>
					<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%sub_total_valor = 0
				end if%>
				<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
					<td valign="top">&nbsp;<%=zero(num_linha)%><%'=" - "&grava_info&"<br>"&compara1%></td>
						<%if campo_0 <> 0 Then%><td valign="top">
							<%if cd_user = 46 then '*** O.S.***%><a href="javascript:;" onclick="window.open('manutencao_2/manutencao_nova.asp?cd_codigo=<%=cd_cod%>&campo=cd_codigo&visual=1&jan=1', 'principal', 'width=651, height=500,scrollbars=1'); return false;"><%=zero(cd_unidade)%>.<%=manutencao(num_os)%></a>
							<%elseif cd_user = 10 then'*** Andamento ***%><a href="javascript:;" onclick="window.open('manutencao_2/manutencao_andamento2.asp?cd_codigo=<%=cd_cod%>&campo=cd_codigo&visual=1&jan=1', 'principal', 'width=651, height=500,scrollbars=1'); return false;"><%=zero(cd_unidade)%>.<%=manutencao(num_os)%></a>
							<%else%><a href="javascript:;" onclick="window.open('manutencao_2/manutencao_ver_janela2.asp?cd_codigo=<%=cd_cod%>&campo=cd_codigo&visual=1&jan=1', 'principal', 'width=600, height=500,scrollbars=1'); return false;"><%=zero(cd_unidade)%>.<%=manutencao(num_os)%></a>
							<%end if%>
							</td>
						<%end if%>
						
						<%if campo_3 <> 0 Then%><td valign="top"><%response.write(nm_area)%></td><%end if%>
						<%if campo_4 <> 0 Then%><td valign="top"><%response.write(nm_cc)%></td><%end if%>
						<%if campo_5 <> 0 Then%><td valign="top"><%response.write(nm_conta)%></td><%end if%>
						<%if campo_6 <> 0 Then%><td valign="top"><%response.write(nm_sigla)%></td><%end if%>
						<%if campo_7 <> 0 Then%><td valign="top" align="center"><%response.write(nm_tipo_abrv)%></td><%end if%>
						<%if campo_8 <> 0 Then%>
							<td><span onClick="window.open('financeiro/diario_lista_pag3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=month(dt_vencimento)%>&ano_sel=<%=year(dt_vencimento)%>','Visualizar','width=500,height=<%if cd_tipo_valor=1 then%>240<%elseif cd_tipo_valor=2 then%>480<%elseif cd_tipo_valor=3 then%>460<%end if%>,scrollbars=1')" return false;" style="border-style:none;">
								<%="<b>"&nm_fornecedor &"</b>"%><%if nm_pagamento<> "" Then%><%=" - "&nm_pagamento&" <!--span style='color:"&cor_linha&"'>"&cd_codigo&"</span-->"%><%end if%>
								<%if cd_qtd_parcelas > 1 then response.write("&nbsp; <span style='color:red;'><b>"&cd_parcela&"<b>/"&cd_qtd_parcelas&"</span>")%></td><%end if%></span>
						<%if campo_10 <> 0 Then%><td valign="top" align="center"><a href="javascript:void(0);" title="<%="Pago em:"&dt_pagamento_red%>"><%response.write(nm_modo_pagamento)%><%if cd_modo_pagamento = 1 then response.write(":"&nm_cheque)%></a><%'=compara1%></td><%end if%>
						<%if campo_2 <> 0 Then%><td valign="top" align="center"><%response.write(dt_vencimento)%></td><%end if%>
						<%if campo_1 <> 0 Then%><td valign="top" align="right" onClick="manipulacao_valor('<%=cd_diario%>','<%=dt_vencimento_anterior%>','<%=month(dt_vencimento)%>','<%=year(dt_vencimento)%>','editar_valor<%'=str_acao%>','<%=cd_user%>')";><%if isnumeric(nm_valor) then response.write(formatnumber(nm_valor,2))%></td><%end if%>

		
		
<!--td valign="top"><%'=fecha%></td-->
			
			</tr>
			<%
			if cor > 0 then
				cor_linha = "#d7d7d7"
				cor_linha2 = "#d7d7d7"
				cor = 0
			else
				cor_linha = "#FFFFFF"
				cor_linha2 = "#e9e9e9"
				cor = 1
			end if
	
			nm_gestao = ""
			nm_situacao = ""
			sub_total_valor = sub_total_valor + nm_valor
			
			rs.movenext
			grava_info = compara1
			
			nr = 1
			geral = geral + conta
			num_linha = num_linha + 1
			loop
			
			'end if%>
			
			<%'*** Exibe o subtotal depois da ultima linha (2) ***
			'if compara1 <> grava_info AND sub_total_valor > 0 AND subtotais <> "" then
			if sub_total_valor > 0 AND subtotais <> "" AND pri_subtotal <> "" then%>
					<tr style="background-color:#b5b5b5;">
						<td colspan="8">&nbsp;</td>
						<td align="center"><%if sub_total_valor <> "" Then 
								'porcentagem_valor_total = sub_total_valor/pre_total
								response.write("<b>"&formatNumber(sub_total_valor/pre_total,3)*100&"%</b>")
							end if%></td>
						<td align="right"><b><i><%if sub_total_valor <> "" Then response.write(formatNumber(sub_total_valor,2))%></i></b></td>
					</tr>
			<%sub_total_valor = 0
			end if%>
			<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="1"></td></tr>					
			<tr bgcolor="#C0C0C0" style="font-size:12px;">
					<%if campo_3 <> 0 Then%><td></td><%end if%>
					<%if campo_4 <> 0 Then%><td></td><%end if%>
					<%if campo_5 <> 0 Then%><td></td><%end if%>
					<%if campo_6 <> 0 Then%><td></td><%end if%>
					<%if campo_7 <> 0 Then%><td></td><%end if%>
					<%if campo_10 <> 0 Then%><td></td><%end if%>
					<%'if campo_2 <> 0 Then%><td></td><%'end if%>
					<%'if campo_1 <> 0 Then%><td></td><%'end if%>
					
					<td align="center">&nbsp;<b>TOTAL:</b></td>
					<td align="right"><b><%=formatnumber(total_valor,2)%></b></td>
				</tr>
		<%End if%>
				<tr class="ok_print">
						<td><img src="imagens/px.gif" width="25" height="1"></td>
						<td><img src="imagens/px.gif" width="20" height="1">xxx</td>
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
			<tr class="no_print">
						<td><img src="imagens/px.gif" width="25" height="1"></td>
<%if campo_3 <> 0 Then%><td><img src="imagens/blackdot.gif" width="55" height="1"></td><%end if%>
<%if campo_4 <> 0 Then%><td><img src="imagens/blackdot.gif" width="40" height="1"></td><%end if%>
<%if campo_5 <> 0 Then%><td><img src="imagens/blackdot.gif" width="85" height="1"></td><%end if%>
<%if campo_6 <> 0 Then%><td><img src="imagens/blackdot.gif" width="50" height="1"></td><%end if%>
<%if campo_7 <> 0 Then%><td><img src="imagens/blackdot.gif" width="40" height="1"></td><%end if%>
<%if campo_8 <> 0 Then%><td><img src="imagens/blackdot.gif" width="250" height="1"></td><%end if%>
<%if campo_10 <> 0 Then%><td><img src="imagens/blackdot.gif" width="30" height="1"></td><%end if%>
<%if campo_2 <> 0 Then%><td><img src="imagens/blackdot.gif" width="60" height="1"></td><%end if%>
<%if campo_1 <> 0 Then%><td><img src="imagens/blackdot.gif" width="75" height="1"></td><%end if%>

<%END IF%>				
				</tr>				
			</table>
			<%if session_cd_usuario = 46 Then%>
			<br>
			<table>
				<!--tr class="txt">
					<td align="left" colspan="100%">
					<b>SELECT SUM</b>(<%=select_option%>) as conta,<%=campos%><br>
					<b>FROM VIEW_financeiro_valores3 </b><%=cond%> <%=cond_periodo%> <%'=cond_ano%> <%=cond_ns%> <%=cond_eqp%> <%=cond_eqp_1%> <%'=cond_pat%> <%=cond_patrimonio%> <%=cond_und%> <%=cond_forn%> <%=cond_esp%> <%=cond_aval%> <%=cond_marca%> <%=cond_valor%><br>
					<%agrupamento = replace(agrupamento,"GROUP BY","<b>GROUP BY</b>")%>
					<%=agrupamento%><br>
					<%ordem = replace(ordem,"ORDER BY","<b>ORDER BY</b>")%>
					<%=ordem%> <br>
					<b><%=sentido%></b></td>
				</tr-->
				<tr>
					<td>
						<%=selecao&" FROM VIEW_financeiro_valores3 "&cond&" "&cond_area&" "&cond_centrocusto&" "&cond_conta&" "&cond_unidade&" "&cond_tipo_valor&" "&cond_periodo&" "&cond_fornecedor&" "&cond_pagamento&" "&agrupamento&" "&ordem%>
					</td>
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
