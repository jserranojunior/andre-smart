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
arquivo_loc = "rel_orc_aprov_2.asp"

tipo = request("tipo")
cd_orcamento = request("cd_orcamento")
campos = request("campos")

cd_fornecedor = request("cd_fornecedor")
nr_orcamento = request("nr_orcamento")

tipo_nao_entregue = request("tipo_nao_entregue")
tipo_campo_data = "dt_aprov_orc"

nm_valor = request("nm_valor")
nr_parcela = request("nr_parcela")


dt_ano = request("ano_sel")
dt_mes = request("mes_sel")
	if dt_mes = "" OR dt_ano = "" then
		dt_mes = month(now)
		dt_ano = year(now)
	end if
		
		dia_final = ultimodiames(dt_mes,dt_ano)
		
		mes_anterior = dt_mes - 1
		ano_anterior = dt_ano
			if mes_anterior < 1 then
				mes_anterior = 12
				ano_anterior = dt_ano - 1
			end if
			dia_final_anterior = ultimodiames(mes_anterior,ano_anterior)
			
		mes_posterior = dt_mes + 1
		ano_posterior = dt_ano
			if mes_posterior > 12 then
				mes_posterior = 1
				ano_posterior = dt_ano + 1
			end if
			dia_final_posterior = ultimodiames(mes_posterior,ano_posterior)
	'pos_atual = right(dt_ano,2)&zero(month(dt_mes))


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
			else
				ordem = "ORDER BY dt_entrega, dt_aprov_orc"				
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
	<table align="center" class="no_print;">
		
		<tr>
			<td align="center" colspan="7" style="background-color:gray; color:white; font-"><b>Relatórios para provisão de<br>pagamento de orçamentos aprovados</b></td>
		</tr>
		
		<%if tipo = "financ" then
			'action = "cad_orc_aprov"
			action = "rel_orc_aprov_2"
		elseif tipo = "cad_orc_aprov" then
			'action = "manutencao2"
			action = "rel_orc_aprov_2"
		elseif tipo = "rel_orcamentos" then
			action = "manutencao2_2"
			'action = "rel_orc_aprov"
		elseif tipo = "man_rep_aprov" then
			action = "manutencao2"
		end if%>
		<form name="form" action="<%=action%>.asp" method="get" id="form">
			<input type="hidden" name="tipo" value="<%=tipo%>">
	
		<%bgc_linha = "dddFFF"%>
		<tr>
			<td><b>M&ecirc;s</b></td>
			<td><b>Ano</b></td>
		</tr>
		<tr>
			
			<td><select name="mes_sel">
						<option value=""></option>
						<%for imes = 1 to 12%>
							<option value="<%=imes%>" <%if abs(imes)=abs(dt_mes) then response.write("SELECTED")%>><%=mes_selecionado(imes)%></option>
						<%next%>
				</select></td>	
			<td><input type="text" name="ano_sel" maxlength="4" size="4" value="<%=dt_ano%>">&nbsp;<input type="submit" name="OK" value="OK" width="4" height="5"></td>
		</tr>
		<tr>
			<td align="center" colspan="2">
				<a href="<%=action%>.asp?tipo=<%=tipo%>&mes_sel=<%=mes_anterior%>&ano_sel=<%=ano_anterior%>"><< Anterior </a>&nbsp; &nbsp; &nbsp; 
				<a href="<%=action%>.asp?tipo=<%=tipo%>&mes_sel=<%=month(now)%>&ano_sel=<%=year(now)%>">Atual</a> &nbsp; &nbsp; &nbsp; 
				<a href="<%=action%>.asp?tipo=<%=tipo%>&mes_sel=<%=mes_posterior%>&ano_sel=<%=ano_posterior%>">Próximo>></a></td>
		<tr>
		</form>
	</table>

<%'end if%>
<br class="no_print">
<!-- <table align="center" class="no_print">
	<tr>
		<%if tipo = "financ" Then%>
			<td align="center" colspan="1">
				<input type="button" name="pagtos" value="Montar lista para pagamento" onclick="monta_pgto(document.form.total_pgto.value);" class="no_print">
			</td>
		<%elseif tipo = "cad_orc_aprov" then%>
			<td align="center" colspan="1">
				<input type="submit" name="novo_orc" value="Novo Orçamento" onclick="window.open('../manutencao_2/orcamentos_aprovados_cad.asp?tipo=<%=tipo%>,cd_funcionario=<%=strcod%>','orccad','width=500,height=200,toolbar=0,location=0,status=0,scrollbars=0,resizable=1'); return false;">
			</td>
		<%end if%>
	</tr>
</table>
-->

<br class="no_print">

<table align="center" border="1" style="">
	<%
	'condicoes = cond_forn 
	
	'num = 1
	maior_qtd_parcelas = 0
	cancela_subtotal = 0
	ultima_dt_entrega = NULL
	
	strsql ="sp_orcamento_s01 @funcao=S, @dt_inicial='"&dt_ano&"-"&dt_mes&"-1', @dt_final='"&dt_ano&"-"&dt_mes&"-"&dia_final&"'"
		Set rs = dbconn.execute(strsql)
			while not rs.EOF
				cd_orcamento = rs("cd_codigo")
				cd_fornecedor = rs("cd_fornecedor")
				nm_fornecedor = rs("nm_fornecedor")
				nr_orcamento = rs("nr_orcamento")
				'dt_orcamento = rs("dt_orcamento")
				dt_aprov_orc = rs("dt_aprov_orc")
				dt_entrega = rs("dt_entrega")
					if isnull(dt_entrega) Then
						dt_entrega = NULL
					else
						dt_entrega = zero(day(dt_entrega))&"/"&zero(month(dt_entrega))&"/"&right(year(dt_entrega),2)
					end if
				nm_valor = rs("nm_valor")
					'total_nm_valor = total_nm_valor + nm_valor
					total_nm_valor = abs(total_nm_valor) + abs(nm_valor)
				nr_parcela = rs("nr_parcela")
					if nr_parcela >= 1 then
						nr_valor_parcela = nm_valor / nr_parcela
					else
						nr_valor_parcela = 0
					end if
				
				if nr_parcela > maior_qtd_parcelas then
					maior_qtd_parcelas = nr_parcela
				end if
				
				dt_inicio = rs("dt_i")
				dt_fim = rs("dt_f")
				
				coluna_divisora = rs("coluna_divisora")
					'if coluna_divisora = 3 then
					'	cab_relatorio = "Compromissos - Pagamentos no &quot;Contas a Pagar&quot; - "
					'	cor_relat = "#d1e9d8"
					'elseif coluna_divisora = 2 then
					'	cab_relatorio = "Orçamentos Aprovados recentes -"
					'	cor_relat = "#f1f5ba"
					'elseif coluna_divisora = 1 then
					'	cab_relatorio = "Orçamentos Aguardando Aprovação - "
					'	cor_relat = "#ec9997"
					'else
					'	cab_relatorio = "Erro no processo - Verificar - "
					'end if
				
				'ultima_divisao = coluna_divisora 
				
					pos_atual = right(dt_ano,2)&zero(dt_mes)
					pos_aprov = right(year(dt_aprov_orc),2)&zero(month(dt_aprov_orc))
					pos_dif = pos_atual - pos_aprov
					qtd_parc_real = (nr_parcela - pos_dif)
					
					IF coluna_divisora = 1 Then
						difereca_dt = DATEDIFF("m","1/"&month(dt_orcamento)&"/"&year(dt_orcamento),"1/"&zero(dt_mes)&"/"&dt_ano)
						cab_relatorio = "Orçamentos Aguardando Aprovação - "
						cor_relat = "#ec9997"
					ELSEIF coluna_divisora = 2 Then
						difereca_dt = DATEDIFF("m","1/"&month(dt_aprov_orc)&"/"&year(dt_aprov_orc),"1/"&zero(dt_mes)&"/"&dt_ano)
						cab_relatorio = "Orçamentos Aprovados recentes -"
						cor_relat = "#f1f5ba"
					ELSEIF coluna_divisora = 3 Then
						difereca_dt = DATEDIFF("m","1/"&month(dt_entrega)&"/"&year(dt_entrega),"1/"&zero(dt_mes)&"/"&dt_ano)
						cab_relatorio = "Compromissos - Pagamentos no &quot;Contas a Pagar&quot; - "
						cor_relat = "#d1e9d8"
					ELSE
						diferenca_dt = 0
						cab_relatorio = "Erro no processo - Verificar - "
						cor_relat = "#eaeaea"
					END IF
				
			'*****************************************************
			'*** Cabeçalho e Rodapé ******************************
			'*****************************************************
			if ultima_divisao <> coluna_divisora then
				if num > 1 Then%>
				<tr>
					<td colspan="4">&nbsp;</td>
					<td colspan="2" align=""right">Subtotal</td>
					<%for i = 1 to 6
						array_parcela = split(array_subtotal,";")
							for each parc_item in array_parcela
								'array_subtotal&";"&i&":"&nr_valor_parcela
								if instr(1,parc_item,i&":",1) then
									if parc_item <> "" then
										parc_item = replace(parc_item,i&":","")
											subtotal = subtotal+abs(parc_item)
									end if
								end if
							next%>
						<td align="right"><%if i = 1 then response.write("<b>")%><%if subtotal > 0 Then response.write(formatnumber(subtotal,2))%><%if i = 1 then response.write("</b>")%></td>
					<%subtotal = 0
					next%>
				</tr>
				<tr><td colspan="12">&nbsp;</td></tr>
			<%end if%>
				<tr>
					<td align="center" colspan="12" style="background-color:gray; color:white; font-size:14px;"><%=cab_relatorio&mes_selecionado(dt_mes)&"/"&dt_ano%><%'=dt_ano&"-"&dt_mes&"-"&dt_dia%></td>
				</tr>
				<tr style="background-color:silver; page-break-before: always;">
					<td><b>Num.</b><br><img src='../imagens/px.gif' width='20' height='1'></td>
					<td><b>Nº Orç.</b><br><img src='../imagens/ox.gif' width='60' height='1'></td>
					<td><b>Fornecedor</b><br><img src='../imagens/px.gif' width='200' height='1'></td>
					<td><b>Data aprov.</b><br><img src='../imagens/px.gif' width='60' height='1'></td>
					<!--td><b>Data entrega.</b><br><img src='../imagens/px.gif' width='70' height='1'></td-->
					<!--td><b>Data I/F</b><br><img src='../imagens/px.gif' width='70' height='1'></td-->
					<td><b>Valor Total</b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<td><b>Qtd parc.</b><br><img src='../imagens/px.gif' width='50' height='1'></td>
					<%dt_mes_parcela = dt_mes -1
					dt_ano_parcela = dt_ano
					for i = 1 to 6
						dt_mes_parcela = dt_mes_parcela + 1
							if dt_mes_parcela > 12 then
								dt_mes_parcela = 1
								dt_ano_parcela = dt_ano + 1
							end if%>
								<td align="center"><b><%=left(mes_selecionado(dt_mes_parcela),3)&"/"&right(dt_ano_parcela,2)%></b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<%next%>
					<td>O.S. - Entrega<br><img src='../imagens/px.gif' width='150' height='1'></td>
					
				</tr>
			<%array_subtotal=""
			num = 1
			end if%>
			
			<tr style="background-color:<%=cor_relat%>;">
				<td align="center"><%=num%></td>
				<td><a href="javascript:void(0);" onclick="window.open('<%if tipo = "financ" then response.write("../")%>manutencao_2/orcamentos_aprovados_os.asp?cd_orcamento=<%=cd_orcamento%>', 'orcamento<%=cd_orcamento%>', 'width=520, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=nr_orcamento%></a></td>
				<td><%=nm_fornecedor%> <%=" - "&nr_parcela&" - "&difereca_dt&" = "&qtd_num_parc_rest&" ("&nr_parcela-difereca_dt&")"%>
				</td>
				<td align="center"><%if not ISNULL(dt_aprov_orc) then response.write(zero(day(dt_aprov_orc))&"/"&zero(month(dt_aprov_orc))&"/"&right(year(dt_aprov_orc),2))%></td>
				<!--td align="center"><%=dt_entrega%></td-->
				<!--td align="center"><%=dt_inicio&" - "&dt_fim%></td-->
				
				<td align="right"><%if tipo = "financ" then%><a href="javascript:void(0);" return false;" onclick="window.open('../financeiro/diario_cad3.asp?cd_origem=6.<%=cd_orcamento%>&cd_forn=<%=cd_fornecedor%>&cd_equip=<%=cd_equipamento%>&cd_valor_orc=<%=replace(nm_valor,".","")%>&cd_tipo_orc=<%=cd_gestao%>&cd_unidade=<%=cd_unidade%>&campo=cd_codigo&visual=1&jan=1', 'pagamento_<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%if nm_valor <> "" Then response.write(FormatNumber(nm_valor,2))%></a><%else%><%if nm_valor <> "" Then response.write(formatnumber(nm_valor,2))%><%end if%></td>
				<td align="center"><%=nr_parcela%></td>
				
					<%'******************************************
					'*** Ditribuição das Parcelas - Divisão 3 ***
					'********************************************
					if coluna_divisora	= 3 then
						qtd_parc_rest = nr_parcela-difereca_dt +1
						for i = 1 to qtd_parc_rest
							array_subtotal = array_subtotal&";"&i&":"&nr_valor_parcela
							%>
							<td align="right"><%if i = 1 then response.write("<b>")%><%'=i%> <%=formatnumber(nr_valor_parcela,2)%><%if i = 1 then response.write("</b>")%><br><img src='../imagens/px.gif' width='70' height='1'></td>
						<%next
						coluna_preeenchimento = 6 - qtd_parc_rest
					
					'********************************************
					'*** Ditribuição das Parcelas - Divisão 2 ***
					'********************************************
					elseif coluna_divisora	= 2 then
						nr_parcela_max = 6
							if nr_parcela > 6 Then nr_parcela = nr_parcela_max
						
						for i = 1 to nr_parcela
							array_subtotal = array_subtotal&";"&i&":"&nr_valor_parcela
							%>
							<td align="right"><%if i = 1 then response.write("<b>")%><%'=i%> <%=formatnumber(nr_valor_parcela,2)%><%if i = 1 then response.write("</b>")%><br><img src='../imagens/px.gif' width='70' height='1'></td>
						<%next
						coluna_preeenchimento = 6 - nr_parcelas
						
					end if
						for i = 1 to coluna_preeenchimento%>
							<td align="right"><b><%'=i%></b> <%'=coluna_preeenchimento&"<br>"%><img src='../imagens/px.gif' width='70' height='1'></td>
						<%next%>
				
				
				
				<td>	
						<%if coluna_divisora < 3 then
						
						'strsql = "SELECT * FROM VIEW_os_lista2 WHERE num_os_assist = '"&nr_orcamento&"' ORDER BY dt_resposta_orc"
						strsql = "SELECT * FROM VIEW_os_lista2 WHERE num_os_assist = '"&nr_orcamento&"' ORDER BY dt_resposta_orc"
						Set rs_2 = dbconn.execute(strsql)
						num_os_2 = ""
						conta_os = 1
						ultima_entrega_os = ""
						marca_inconsistencia = 0
			
						while not rs_2.EOF
							cd_unidade_2 = rs_2("cd_unidade_os")
							num_os_2 = rs_2("num_os")
							dt_retorno_2 = rs_2("dt_retorno")
							dt_resposta_orc = rs_2("dt_resposta_orc")
							
							'****** Verifica Inconsistencias *******
							'if ultima_resposta_orc <> dt_resposta_orc  AND conta_os > 1 THEN
							'	marca_inconsistencia = 1
							'	response.write("<b>-></b>")
							
							if ultima_entrega_os <> dt_retorno_2  AND conta_os > 1 THEN
							'	marca_inconsistencia = 1
							'	response.write("<b>-></b>")
								
							ELSEIF  ISNULL(dt_retorno_2) THEN
								marca_inconsistencia = 1
								response.write("<b>--></b>")
							
							ELSEIF  ISNULL(dt_resposta_orc) THEN
								marca_inconsistencia = 1
								response.write("<b>*-></b>")
								
							ELSEIF num_os_2 = "" Then
								marca_inconsistencia = 1
								response.write("<b>***></b>")
							else
								'marca_inconsistencia = 0
							end if
							
							'response.write("("&conta_os&")")
							%>
							
							<%="<b>"&zero(cd_unidade_2)&"."&num_os_2&"</b> - "&dt_retorno_2%><%'=":"&ultima_entrega_os%><br>
						<%
						ultima_entrega_os = dt_retorno_2
						ultima_resposta_orc = dt_resposta_orc
						conta_os = conta_os +1
						
						rs_2.movenext
						wend
						
							IF num_os_2 = "" Then
								marca_inconsistencia = 1
								response.write("<b>-->> </b>")
							end if
							
							'if marca_inconsistencia = 1 AND not ISNULL(dt_retorno_2) AND num_os_2 <> "" then
							if marca_inconsistencia = 1 then
								'response.write("<br><b>xxxXXXxxx</b>")
								
							'elseif marca_inconsistencia = 0 AND not ISNULL(dt_retorno_2) AND num_os_2 <> "" then
							elseif marca_inconsistencia = 0 then
								'********* Efetua a alteração da data de entrega no Orçamento****
								'response.write("@cd_orcamento='"&cd_orcamento&"'<br> @dt_entrega="&year(dt_retorno_2)&"-"&month(dt_retorno_2)&"-"&day(dt_retorno_2)&"")
								'xsql = "up_orcamento_baixa @cd_orcamento='"&cd_orcamento&"', @dt_entrega='"&year(dt_retorno_2)&"-"&month(dt_retorno_2)&"-"&day(dt_retorno_2)&"'"
								'response.write("dt_aprov_orc ='"&year(dt_resposta_orc)&"-"&month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"'  cd_codigo="&cd_orcamento)
								'xsql = "UPDATE TBL_manutencao_orcamento SET dt_aprov_orc ='"&year(dt_resposta_orc)&"-"&month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"' WHERE  cd_codigo="&cd_orcamento&""
								
								'dbconn.execute(xsql)
								response.write("^^^^^^^^^^")
							else
								response.write("...")
							end if
						marca_inconsistencia = 0
						ultima_entrega_os = ""
						end if
						
						conta_os = 1%>
				</td>
				
				
				
						
			</tr>
			
			<%conta_os = 0
			dt_mes_parcela = 0
			coluna_preeenchimento = 0
			
			
			'pos_aprov = 0
			pos_dif = 0
			qtd_parc_real = 1
			
			grava_info = compara1
			num = num + 1
			ultima_divisao = coluna_divisora 
			xptop = ""
			rs.movenext
			wend
			
			'response.write(listaParcelas_total%>
			
			
					
					<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr style="background-color:#b5b5b5;">
						<td colspan="4">&nbsp;</td>
						<td colspan="2" align="center">Subtotal</td>
						<%dt_mes_parcela = dt_mes -1
							for i = 1 to 6
								dt_mes_parcela = dt_mes_parcela + 1
									if dt_mes_parcela > 12 then
										dt_mes_parcela = 1
									end if
									
									'array_parcela = array_parcela&";"&dt_mes_parcela&"!"&nr_valor_parcela
									
							array_parcela = split(listaParcelas_total,";")
							for each parc_item in array_parcela
								'response.write(parc_item&"<br>")
									
								if instr(1,parc_item,dt_ano_parcela&dt_mes_parcela&"!",1) then
									'response.write(parc_item&"<br>")
									if parc_item <> "" then
										parc_item = replace(parc_item,dt_ano_parcela&dt_mes_parcela&"!","")
											'response.write(parc_item&"<br>")
											subtotal = subtotal+abs(parc_item)
									end if'soma = soma+parc_item&"_"
										
								end if
							next%>
										<td align="right"><b><%'=dt_mes_parcela%></b> <%=formatnumber(subtotal,2)%><%'=subtotal%><br><img src='../imagens/px.gif' width='70' height='1'></td>
							<%subtotal = 0
							next%>
					</tr>
			<%'sub_total_valor = 0
			'end if%>
			<!-- <tr><td Colspan = "7"><%'=listaParcelas%></td></tr>-->
			<tr style="background-color:silver;">
				<td colspan="3">&nbsp;</td>
				<td align="center"><b>Total</b></td>
				<td align="right"><b><%if total_nm_valor <> "" Then response.write(formatNumber(total_nm_valor,2))%></b></td>
				<td colspan="7">&nbsp;</td>
				
			</tr>
					</table>
				</td>
			</tr>
			<%'=response.write(listaParcelas_total)%>			
</table>
<br>
<br>
</body>

<%if tipo <> "financ" then%>			
	<SCRIPT LANGUAGE="javascript">
		formataMoeda(document.forms.form.nm_valor, 2);	
		
		function JsDelete(cod)
		   {
			if (confirm ("Tem certeza que deseja excluir o orçamento?"))
		  {
		  document.location.href('manutencao_2/acoes/manutencao_acao.asp?cd_orcamento='+cod+'&acao=3&filtro=orc_cad_del');
			}
			  }
	</SCRIPT>
<%end if%>
