<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<!-- #include file="../css/geral.htm"-->
<body>
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
	
	document.form.ordem_inicial.value='1';
	
	document.ordem_1.style.display='inline';
	document.ordem_2.style.display='inline';
	document.ordem_3.style.display='inline';
	document.ordem_4.style.display='inline';
	document.ordem_5.style.display='inline';
	}

function adiciona_pgto(a,lista_pgto,check1,check2){
	lista_pgto = lista_pgto+','+a
	document.form.total_pgto.value=(lista_pgto);
	var el1=document.getElementById(check1);
		el1.style.display=(el1.style.display!='none'?'none':'');
	var el2=document.getElementById(check2);
		el2.style.display=(el2.style.display!='none'?'none':'');		
	}
	
function subtrai_pgto(b,lista_pgto,check2,check1){
	lista_pgto = lista_pgto.replace(','+b,'');
	document.form.total_pgto.value=(lista_pgto);
	var el2=document.getElementById(check2);
		el2.style.display=(el2.style.display!='none'?'none':'');
	var el1=document.getElementById(check1);
		el1.style.display=(el1.style.display!='none'?'none':'');
	}

function monta_pgto(lista_pgto){
	mywindow = window.open("orcamentos_aprovados_pacote.asp?lista_pagto="+lista_pgto+"&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>", "Emissao_Cheques", "location=0,status=1,width=700,height=400,scrollbars=1, resizable=1");  
	mywindow.moveTo(0, 0);
	}
// -->	
</script>
<%cd_user = session("cd_codigo")
pasta_loc = "manutencao_2\"
arquivo_loc = "cad_orc_aprov.asp"

tipo = request("tipo")
cd_orcamento = request("cd_orcamento")
campos = request("campos")

cd_fornecedor = request("cd_fornecedor")
nr_orcamento = request("nr_orcamento")

dt_diai = request("dt_diai")
dt_mesi = request("dt_mesi")
dt_anoi = request("dt_anoi")

dt_diaf = request("dt_diaf")
dt_mesf = request("dt_mesf")
dt_anof = request("dt_anof")

	if dt_diaf = "" OR dt_mesf = "" OR dt_anof = ""  Then
		if dt_diai <> "" OR dt_mesi <> "" OR dt_anoi <> "" Then
			dt_diaf = ultimodiames(dt_mesi,dt_anoi)
			dt_mesf = request("dt_mesi")
			dt_anof = request("dt_anoi")
		end if
	elseif ultimodiames(dt_mesi,dt_anoi) < dt_diaf then
		dt_diaf = ultimodiames(dt_mesf,dt_anof)
	end if

tipo_nao_entregue = request("tipo_nao_entregue")
tipo_data_orc = request("tipo_data_orc")
	if tipo_data_orc = 1 then
		tipo_campo_data = "dt_orcamento"
		periodo_filtro = " between '"&dt_mesi&"/"&dt_diai&"/"&dt_anoi&"' AND '"&dt_mesf&"/"&dt_diaf&"/"&dt_anof&"'"
	elseif tipo_data_orc = 2 then
		tipo_campo_data = "dt_aprov_orc"
		periodo_filtro = " between '"&dt_mesi&"/"&dt_diai&"/"&dt_anoi&"' AND '"&dt_mesf&"/"&dt_diaf&"/"&dt_anof&"'"
	elseif tipo_data_orc = 3 then
		tipo_campo_data = "dt_entrega"
		periodo_filtro = " between '"&dt_mesi&"/"&dt_diai&"/"&dt_anoi&"' AND '"&dt_mesf&"/"&dt_diaf&"/"&dt_anof&"'"
	end if

nm_valor = request("nm_valor")
nr_parcela = request("nr_parcela")

'****************************************************
'***** Verifica quais campos foram selecionados *****
	'campo_forn = Instr(1,campos,"cd_fornecedor,nm_fornecedor",0)
	'response.write("<br>"&campos)
'****************************************************

If cd_fornecedor = "0" OR cd_fornecedor = "" then
	cond_forn = "" 
else
	cond_forn = " AND cd_fornecedor = "&cd_fornecedor&""
End if

if nr_orcamento = "" Then
	cond_n_orc = "" 
else
	cond_n_orc = " AND nr_orcamento like '"&nr_orcamento&"%' "
end if

if dt_diai = "" AND dt_mesi = "" AND dt_anoi = "" Then
	'cond_periodo_i = " "
	'
	dt_mesi = zero(month(now))
	dt_diai = zero("1")
	dt_anoi = year(now)
	
	dt_mesf = zero(month(now))
	dt_diaf = ultimodiames(month(now),year(now))
	dt_anof = year(now)
	cond_periodo_i = " AND dt_orcamento between '"&dt_mesi&"/"&dt_diai&"/"&dt_anoi&"' AND '"&dt_mesf&"/"&dt_diaf&"/"&dt_anof&"'"
else
	'cond_periodo_i = " AND dt_orcamento between '"&dt_mesi&"/"&dt_diai&"/"&dt_anoi&"' AND '"&dt_mesf&"/"&dt_diaf&"/"&dt_anof&"'"
	cond_periodo_i = " AND "&tipo_campo_data&" between '"&dt_mesi&"/"&dt_diai&"/"&dt_anoi&"' AND '"&dt_mesf&"/"&dt_diaf&"/"&dt_anof&"'"
	'cond_periodo_i = " AND "&tipo_campo_data &" "& periodo_filtro
end if

if tipo_nao_entregue = "" then
	cond_nao_entregue = ""
else
	cond_nao_entregue = " AND dt_entrega IS NULL "
end if

If nm_valor = "" then
	cond_valor = " " 
else
	cond_valor = " AND nm_valor like '%"&nm_valor&"%' "
End if

If nr_parcela = "" then
	cond_parcela = "" 
else
	cond_parcela = " AND nr_parcela = '"&nr_parcela&"' "
End if

'***** Tratamento da ordem de listagem ******
'*** 2ª etapa ***
ordem_total = request("ordem_total")
	
	'** Datas **
	verif_dt_campo1 = instr(1,ordem_total,"dt_orcamento",1)
	verif_dt_campo2 = instr(1,ordem_total,"dt_aprov_orc",1)
	verif_dt_campo3 = instr(1,ordem_total,"dt_entrega",1)
	
		if verif_dt_campo1 = 1 AND tipo_campo_data = "dt_aprov_orc" then
			ordem_total = replace(ordem_total,"dt_orcamento","dt_aprov_orc")
		elseif verif_dt_campo1 = 1 AND tipo_campo_data = "dt_entrega" then
			ordem_total = replace(ordem_total,"dt_orcamento","dt_entrega")
		end if
		
		if verif_dt_campo2 = 1 AND tipo_campo_data = "dt_orcamento" then
			ordem_total = replace(ordem_total,"dt_aprov_orc","dt_orcamento")
		elseif verif_dt_campo2 = 1 AND tipo_campo_data = "dt_entrega" then
			ordem_total = replace(ordem_total,"dt_aprov_orc","dt_entrega")
		end if
		
		if verif_dt_campo3 = 1 AND tipo_campo_data = "dt_orcamento" then
			ordem_total = replace(ordem_total,"dt_entrega","dt_orcamento")
		elseif verif_dt_campo3 = 1 AND tipo_campo_data = "dt_aprov_orc" then
			ordem_total = replace(ordem_total,"dt_entrega","dt_aprov_orc")
		end if
		
	'response.write(verif_dt_campo1&","&verif_dt_campo2&","&verif_dt_campo3)
	

	ver_string_ordem = instr(ordem_total,",")
		if ver_string_ordem = 1 then
			ordem_total = mid(ordem_total,2,len(ordem_toral))
		else
			campo_sub = ordem_total
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
		
		'**** Prepara para os subtotais ****
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
		
sentido = request("sentido")
subtotais = request("subtotais")%>

<%'if tipo = "cad_orc_aprov" then%>
<!--#include file="../includes/arquivo_loc.asp"-->
	<table align="center">
		
		<tr>
			<td align="center" colspan="7" style="background-color:gray; color:white; font-"><b>Gestão dos Orçamentos Aprovados</b></td>
		</tr>
		<tr>
			<td colspan="8">&nbsp;
		<%if tipo = "financ" then
			action = "cad_orc_aprov"
		elseif tipo = "cad_orc_aprov" then
			action = "manutencao2"
		end if%>
		<form name="form" action="<%=action%>.asp" method="get" id="form">
			<input type="hidden" name="total_pgto" value="">
			<input type="hidden" name="tipo" value="<%=tipo%>">
			<input type="hidden" name="ordem_res">
			<input type="hidden" name="ordem_total" value="<%=ordem_total%>">
			<input type="hidden" name="ordem_inicial" value="<%if ordem_inicial = "" Then%>1<%else response.write(ordem_inicial) end if%>">
			</td>
			<td><a href="javascript:void();" onClick="limpa_ordem()">Limpa Ordem</a></td>
		</tr>
		<%bgc_linha = "dddFFF"%>
		<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
			<td colspan="5"><b>Empresa</b><br>
				<select name="cd_fornecedor">
					<option value="0">Todos</option>
					<%xsql_forn = "Select * From TBL_fornecedor order by nm_fornecedor"
					Set rs_forn = dbconn.execute(xsql_forn)
						while not rs_forn.EOF
							cd_forn = rs_forn("cd_codigo")
							nm_forn = rs_forn("nm_fornecedor")%>
							<option value="<%=cd_forn%>" <%if abs(cd_fornecedor) = abs(cd_forn) then response.write("selected")%>><%=nm_forn%></option>
						<%rs_forn.movenext
						wend%>
				</select>
			</td>
			<td>
			<%ordem_var = ordem_1
			ordem_num="ordem_1"
			ordem_campo="nm_fornecedor"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
		</tr>
		<%bgc_linha = "dddFFF"%>
		<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
			<td colspan="5"><b>Nº Orç.</b><br>
				<img src="../imagens/px.gif" width="120" height="1"><br>
				<input type="text" name="nr_orcamento" size="10" maxlength="20" value="<%=nr_orcamento%>"></td>
			<%'strsql ="TBL_Fornecedor order by nm_fornecedor"
			'Set rs_forn = dbconn.execute(strsql)%>
			</td>
			<td>
			<%ordem_var = ordem_2
			ordem_num="ordem_2"
			ordem_campo="nr_orcamento"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
		</tr>
		<%bgc_linha = "dddFFF"%>
		<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>"><%=tipo_data_orc%>
			<td align="center"><b>Abertura</b><br><input type="radio" name="tipo_data_orc" value="1" <%if tipo_data_orc = "" or tipo_data_orc = 1 then response.write("checked")%>></td>
			<td align="center"><b>Aprov.</b><br><input type="radio" name="tipo_data_orc" value="2"<%if tipo_data_orc = 2 then response.write("checked")%>></td>
			<td align="center"><b>Devol.</b><br><input type="radio" name="tipo_data_orc" value="3"<%if tipo_data_orc = 3 then response.write("checked")%>></td>
			<td><b>Período De:</b><br>
				<img src="../imagens/px.gif" width="170" height="1"><br>
				<input type="text" name="dt_diai" size="2" maxlength="2" value="<%=dt_diai%>" id="diai" onkeyup="javascript:JumpField(this,'mesi');">/
				<input type="text" name="dt_mesi" size="2" maxlength="2" value="<%=dt_mesi%>" id="mesi" onkeyup="javascript:JumpField(this,'anoi');">/
				<input type="text" name="dt_anoi" size="4" maxlength="4" value="<%=dt_anoi%>" id="anoi" onkeyup="javascript:JumpField(this,'diaf');"></td>
			<td><b>Até</b><br><img src="../imagens/px.gif" width="170" height="1"><br>
				<input type="text" name="dt_diaf" size="2" maxlength="2" value="<%=dt_diaf%>" id="diaf" onkeyup="javascript:JumpField(this,'mesf');">/
				<input type="text" name="dt_mesf" size="2" maxlength="2" value="<%=dt_mesf%>" id="mesf" onkeyup="javascript:JumpField(this,'anof');">/
				<input type="text" name="dt_anof" size="4" maxlength="4" value="<%=dt_anof%>" id="anof"></td>
			<td>
			<%ordem_var = ordem_3
			ordem_num="ordem_3"
			ordem_campo=tipo_campo_data'"dt_orcamento"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
		</tr>
		<%bgc_linha = "dddFFF"%>
		<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
			<td align="center" colspan="3"><b>Não entregue</b><br><input type="checkbox" name="tipo_nao_entregue" value="1" <%if tipo_nao_entregue = "1"  then response.write("checked")%>></td>
			<!--td><b>Período De:</b><br>
				<img src="../imagens/px.gif" width="170" height="1"><br>
				<input type="text" name="dt_diai" size="2" maxlength="2" value="<%=dt_diai%>" id="diai" onkeyup="javascript:JumpField(this,'mesi');">/
				<input type="text" name="dt_mesi" size="2" maxlength="2" value="<%=dt_mesi%>" id="mesi" onkeyup="javascript:JumpField(this,'anoi');">/
				<input type="text" name="dt_anoi" size="4" maxlength="4" value="<%=dt_anoi%>" id="anoi" onkeyup="javascript:JumpField(this,'diaf');"></td>
			<td><b>Até</b><br><img src="../imagens/px.gif" width="170" height="1"><br>
				<input type="text" name="dt_diaf" size="2" maxlength="2" value="<%=dt_diaf%>" id="diaf" onkeyup="javascript:JumpField(this,'mesf');">/
				<input type="text" name="dt_mesf" size="2" maxlength="2" value="<%=dt_mesf%>" id="mesf" onkeyup="javascript:JumpField(this,'anof');">/
				<input type="text" name="dt_anof" size="4" maxlength="4" value="<%=dt_anof%>" id="anof"></td-->
			<td colspan="2">&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<%bgc_linha = "dddFFF"%>
		<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
			<td colspan="5"><b>valor</b><br>
				<img src="../imagens/px.gif" width="120" height="1"><br>
				<input type="text" name="nm_valor" size="10" maxlength="20" value="<%=nm_valor%>" onKeyup="valores();" style="text-align:right"></td>
			<td>
			<%ordem_var = ordem_4
			ordem_num="ordem_4"
			ordem_campo="nm_valor"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
		</tr>
		<%bgc_linha = "dddFFF"%>
		<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
			<td colspan="5"><b>Qtd. Parcelas.</b><br>
				<img src="../imagens/px.gif" width="80" height="1"><br>
				<input type="text" name="nr_parcela" size="2" maxlength="3" value="<%=nr_parcela%>"></td>
			<td>
			<%ordem_var = ordem_52
			ordem_num="ordem_5"
			ordem_campo="nm_valor"%>
			<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td>
		</tr>
		<tr>
			<td colspan="3"><input type="checkbox" name="subtotais" value="1" <%if subtotais = 1 then response.write("Checked")%>>&nbsp;Subtotais.</td>
			<td><input type="submit" value="Filtrar"></td>
		</tr>
		</form>
	</table>

<%'end if%>
<br class="no_print">
<table align="center" class="no_print">
	<tr class="no_print">
		<%if tipo = "financ" Then%>
			<td align="center" colspan="1">
				<input type="button" name="pagtos" value="Montar lista para pagamento" onclick="monta_pgto(document.form.total_pgto.value);" class="no_print">
			</td>
		<%elseif tipo = "cad_orc_aprov" then%>
			<td>
				<input type="submit" name="novo_orc" value="Novo Orçamento" onclick="window.open('../manutencao_2/orcamentos_aprovados_cad.asp?tipo=<%=tipo%>,cd_funcionario=<%=strcod%>','orccad','width=500,height=280,toolbar=0,location=0,status=0,scrollbars=0,resizable=1'); return false;">
			</td>
		<%end if%>
	</tr>
</table>

<br class="no_print">

<table align="center" border="1" style="">
	<tr>
		<td align="center" colspan="12" style="background-color:gray; color:white; font-size:14px;">Lista de orçamentos cadastrados <%'=" - "& pri_subtotal%></td>
	</tr>
	<%
	'condicoes = cond_forn 
	
	num = 1
	strsql ="SELECT * FROM View_manutencao_orcamento Where cd_fornecedor > '0' "&cond_forn&" "&cond_n_orc&" "&cond_periodo_i&" "&cond_nao_entregue&" "&cond_valor&" "&cond_parcela&" "&ordem&""
		Set rs = dbconn.execute(strsql)
			while not rs.EOF
				cd_orcamento = rs("cd_codigo")
				cd_fornecedor = rs("cd_fornecedor")
				nm_fornecedor = rs("nm_fornecedor")
				nr_orcamento = rs("nr_orcamento")
				dt_orcamento = rs("dt_orcamento")
				dt_aprov_orc = rs("dt_aprov_orc")
				dt_entrega = rs("dt_entrega")
				nm_valor = rs("nm_valor")
					'total_nm_valor = total_nm_valor + nm_valor
					if nm_valor <> "" Then
						total_nm_valor = abs(total_nm_valor) + abs(nm_valor)
					else
						total_nm_valor = 0
					end if
				nr_parcela = rs("nr_parcela")
					if nr_parcela > 1 then
						nr_valor_parcela = nm_valor / nr_parcela
					else
						nr_valor_parcela = 0
					end if
				cd_diario = rs("cd_diario")
					'************** Verifica se já consta no relatório de orçamentos *********************
					strsql_2 ="SELECT cd_codigo FROM TBL_manutencao_orc_aprov WHERE cd_orcamento = "&cd_orcamento
					Set rs_2 = dbconn.execute(strsql_2)
					if not rs_2.EOF Then
						cadastrado = 1
					end if
					
					
					
				
				'**** Comparação Subtotal ****
				if pri_subtotal <> "" Then
					compara1 = rs(pri_subtotal)
					if isNULL(compara1) Then
						compara1 = "nada"
					end if
				end if
				
			if num = 1 then%>
				<tr style="background-color:silver">
					<td><b>Num.</b><br><img src='../imagens/px.gif' width='20' height='1'></td>
					<%'if tipo = "cad_orc_aprov" Then%>
						<!--b>&nbsp;</b-->
					<%if tipo = "financ" Then%>
						<td><b>&nbsp;</b><br><img src='../imagens/px.gif' width='20' height='1'></td>
					<%end if%>	
					<td><b>Fornecedor</b><br><img src='../imagens/px.gif' width='200' height='1'></td>
					<td><b>Nº Orçamento</b><br><img src='../imagens/ox.gif' width='90' height='1'></td>
					<td><b>Data orç.</b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<td><b>Data aprov.</b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<td><b>Data entrega.</b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<td><b>Valor Total</b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<td><b>Qtd parc.</b><br><img src='../imagens/px.gif' width='50' height='1'></td>
					<td><b>Valor parcela</b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<td><b>Diario</b><br><img src='../imagens/px.gif' width='40' height='1'></td>
					<td colspan="2">&nbsp;</td>
				</tr>
			<%end if%>
			<%conta_os = 0
			strsql = "SELECT * FROM VIEW_os_lista2 WHERE num_os_assist = '"&nr_orcamento&"' "
					Set rs_ver = dbconn.execute(strsql)
						while not rs_ver.EOF
						dt_retorno = rs_ver("dt_retorno")
						
							if not isnull(dt_retorno) then 
								back_color = "#ccffcc"
							else
								back_color = "white"
							end if
						conta_os = conta_os + 1
						rs_ver.movenext
						wend%>
			<%'*** Exibe o subtotal depois da ultima linha (1) ***
				'if compara1 <> grava_info AND sub_total_valor > 0 AND subtotais <> "" then
				'if compara1 <> grava_info AND sub_total_valor > 0 AND subtotais <> "" AND pri_subtotal <> "" then
				if compara1 <> grava_info AND subtotais <> "" AND pri_subtotal <> "" then
				'if compara1 <> grava_info AND subtotais <> "" then%>
					<tr style="background-color:#b5b5b5;">
						<td colspan="3">&nbsp;</td>
						
						<td align="center"><%if sub_total_valor <> "" Then 
								'porcentagem_valor_total = sub_total_valor/pre_total
								'response.write("<b>"&formatNumber(sub_total_valor/pre_total*100,2)&"%</b>")
							end if%></td>
						<td align="right"><b><i><%if sub_total_valor <> "" Then response.write(formatNumber(sub_total_valor,2))%></i></b></td>
						<td>&nbsp;</td>
						<td colspan="3">&nbsp;</td>
					</tr>
					<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%sub_total_valor = 0
				end if%>
			
			<tr style="background-color:<%=back_color%>;">
				<td align="center"><%=num%><%'=conta_os%></td>
				<%if tipo = "financ" Then%>
					<td align="center"><%'=cd_modo_pagamento%> <%check1 = "check_1_"&cd_orcamento
						check2 = "check_2_"&cd_orcamento
						'if nm_valor = "0" Then 
						
							if cd_valor <> "" Then 
								str_acao = "editar_valor"
							else
								str_acao="inserir_valor"
							end if
						
						if cd_suspenso = 1 Then%>
							<img src="../../imagens/pagto-suspenso.gif" alt="Pagamento suspenso&#13;para o mês atual" width="12" height="12" border="0">
						<%elseif not IsNumeric(cd_valor) Then%>
							<img src="../../imagens/check_proibido.gif" alt="Inclua o valor" width="12" height="12" border="0" onClick="manipulacao_valor('<%=cd_orcamento%>','<%=dt_vencimento_anterior%>','<%=dt_mes%>','<%=dt_ano%>','<%=str_acao%>','<%=cd_user%>')"; return false;><%=cd_orcamento%>,<%=dt_vencimento_anterior%>,<%=dt_mes%>,<%=dt_ano%>,<%=str_acao%>,<%=cd_user%>
						<%elseif cd_confirma_pagto = int(1) AND cd_pagto_agendado = int(1) then%>
							<img src="../../imagens/check_ok.gif" alt="Pago" width="25" height="12" border="0">
						<%elseif cd_pagto_agendado = int(1) then%>
							<!--img src="../../imagens/check_aguarda.gif" alt="lista cheque(<%=nm_cheque%>)" width="25" height="12" border="0" onClick="window.open('financeiro/diario_pagamentos3.asp?cd_conta_banco=<%=cd_conta_banco%>&nm_cheque=<%=nm_cheque%>','lista_cheque','width=700,height=500,scrollbars=1')"-->
							<%if cd_modo_pagamento = "1" then%>
								<img src="../../imagens/pagto-cheque.gif" alt="lista cheque (<%=nm_cheque%>)" width="12" height="12" border="0" onClick="window.open('financeiro/diario_pagamentos3.asp?cd_pagamento=<%=cd_pagamento%>','lista_cheque','width=700,height=500,scrollbars=1')">
							<%elseif cd_modo_pagamento = "2" then%>
								<img src="../../imagens/pagto-transf.gif" alt="Transferência (<%=zero(day(dt_pagto))&"/"&left(mes_selecionado(month(dt_pagto)),3)&")"%>" width="12" height="12" border="0" onClick="window.open('financeiro/diario_pagamentos3.asp?cd_pagamento=<%=cd_pagamento%>','lista_transf','width=700,height=500,scrollbars=1')">
							<%end if%>
							</td>
						<%else%>	
							<!--img src="../../imagens/check_inativo.gif" alt="Inclui" width="25" height="12" border="0" onClick="adiciona_pgto('<%=cd_valor%>',document.diario.total_pgto.value,'<%=check1%>','<%=check2%>')" id="<%=check1%>" name="<%=check1%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>-->
							<img src="../../imagens/pagto-inativo.gif" alt="Inclui" width="12" height="12" border="0" onClick="adiciona_pgto('<%=cd_orcamento%>',document.form.total_pgto.value,'<%=check1%>','<%=check2%>')" id="<%=check1%>" name="<%=check1%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>>
							<img src="../../imagens/check_ativo_2.gif" alt="Exclui" width="15" height="12" border="0" onClick="subtrai_pgto('<%=cd_orcamento%>',document.form.total_pgto.value,'<%=check2%>','<%=check1%>')" id="<%=check2%>" name="<%=check2%>" <%'if ordem_var <> "" Then%>style="display:none;"<%'end if%>><%'=cd_orcamento%></td>
						<%end if%>
				<%end if%>
				<td><%=nm_fornecedor%></td>
				<td><a href="javascript:void(0);" onclick="window.open('<%if tipo = "financ" then response.write("../")%>manutencao_2/orcamentos_aprovados_os.asp?cd_orcamento=<%=cd_orcamento%>', 'orcamento<%=cd_orcamento%>', 'width=<%if nr_parcela > 1 then%>612<%else%>587<%end if%>, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=nr_orcamento%></a><%if cd_user = 46 AND cadastrado = 1 Then Response.write(" *")%></td>
				<td align="center"><%=zero(day(dt_orcamento))&"/"&zero(month(dt_orcamento))&"/"&right(year(dt_orcamento),2)%></td>
				<td align="center"><%=zero(day(dt_aprov_orc))&"/"&zero(month(dt_aprov_orc))&"/"&right(year(dt_aprov_orc),2)%></td>
				<td align="center"><%=zero(day(dt_entrega))&"/"&zero(month(dt_entrega))&"/"&right(year(dt_entrega),2)%></td>
				<td align="right"><%if tipo = "financ" then%><a href="javascript:void(0);" return false;" onclick="window.open('../financeiro/diario_cad3.asp?cd_origem=6.<%=cd_orcamento%>&cd_forn=<%=cd_fornecedor%>&cd_equip=<%=cd_equipamento%>&cd_valor_orc=<%=replace(nm_valor,".","")%>&cd_tipo_orc=<%=cd_gestao%>&cd_unidade=<%=cd_unidade%>&campo=cd_codigo&visual=1&jan=1', 'pagamento_<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%if nm_valor <> "" Then response.write(FormatNumber(nm_valor,2))%></a><%else%><%if nm_valor <> "" Then response.write(formatnumber(nm_valor,2))%><%end if%></td>
				<td align="center"><%=nr_parcela%></td>
				<td align="right"><%if nr_valor_parcela > 0 then response.write(formatnumber(nr_valor_parcela,2)) Else response.write("&nbsp;") end if%></td>
				<td align="center">
					<%if cd_diario <> "" Then%>
					<span onClick="window.open('financeiro/diario_lista_pag3.asp?cd_diario=<%=cd_diario%>&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>','Visualizar','width=500,height=240,scrollbars=1')" return false;" style="border-style:none;">
						<%'=cd_diario%><img src="../imagens/check_ok.gif" alt="Incluso no diário" width="25" height="12" border="0">
					</span>
					<%end if%>
				</td>
				<%if tipo <> "financ" then%>
					<td><img src="../imagens/ic_editar.gif" alt="" width="13" height="14" border="0" onclick="window.open('manutencao_2/orcamentos_aprovados_cad.asp?cd_orcamento=<%=cd_orcamento%>', 'orccad', 'width=500, height=300, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"></td>
					<td><img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="javascript:JsDelete(<%=cd_orcamento%>)"></td>
				<%else%>
					<td colspan="2">&nbsp;</td>
				<%end if%>
			</tr>
			
			<%conta_os = 0
			sub_total_valor = sub_total_valor + nm_valor
			grava_info = compara1
			num = num + 1
			cadastrado = 0
			rs.movenext
			wend%>
			
			<%'*** Exibe o subtotal depois da ultima linha (2) ***
			'if compara1 <> grava_info AND sub_total_valor > 0 AND subtotais <> "" then
			'if compara1 <> grava_info AND sub_total_valor > 0 AND subtotais <> "" AND pri_subtotal <> "" then
			if subtotais <> "" AND pri_subtotal <> "" then%>
					<tr style="background-color:#b5b5b5;">
						<td colspan="3">&nbsp;</td>
						<td align="center"><%if sub_total_valor <> "" Then 
								'porcentagem_valor_total = sub_total_valor/pre_total
								'response.write("<b>"&formatNumber(sub_total_valor/pre_total,3)*100&"%</b>")
							end if%></td>
						<td align="right"><b><i><%if sub_total_valor <> "" Then response.write(formatNumber(sub_total_valor,2))%></i></b></td>
						<td>&nbsp;</td>
						<td colspan="3">&nbsp;</td>
					</tr>
			<%sub_total_valor = 0
			end if%>
			
			<tr style="background-color:silver;">
				<td colspan="5">&nbsp;</td>
				<td align="center"><b>Total</b></td>
				<td align="right"><b><%if total_nm_valor <> "" Then response.write(formatNumber(total_nm_valor,2))%></b></td>
				<td colspan="5">&nbsp;</td>
			</tr>
</table>
</body>

<%if tipo <> "financ" then%>			
	<SCRIPT LANGUAGE="javascript">
		formataMoeda(document.forms.form.nm_valor, 2);	
		
		function JsDelete(cod)
		   {
			if (confirm ("Tem certeza que deseja excluir o orçamento?"))
		  {
		  document.location.href=('manutencao_2/acoes/manutencao_acao.asp?cd_orcamento='+cod+'&acao=3&filtro=orc_cad_del');
			}
			  }
	</SCRIPT>
<%end if%>
