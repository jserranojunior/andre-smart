<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->

<%'*** Dados do usuário ***
cd_user = request("cd_user")
data_atual = request("data_atual")
acao = request("acao")

cd_diario = request("cd_diario")

cd_qtd_parcelas = request("cd_qtd_parcelas")
	if cd_qtd_parcelas = "" then
		cd_qtd_parcelas = 1
	end if
cd_parcela = request("cd_parcela")
	if cd_parcela = "" then
		cd_parcela = 1
	end if
cd_tipo_valor = request("cd_tipo_valor")

mes_atual = month(now)
ano_atual = year(now)
	data_atual = ano_atual&zero(mes_atual)

dia_vencimento = request("dia_vencimento")
mes_vencimento = request("mes_vencimento")
ano_vencimento = request("ano_vencimento")
	data_escolhida = ano_vencimento&mes_vencimento
	'*** Calculo de data (Parcelado) ***
		if acao = "inserir" AND cd_parcela > 1 AND cd_tipo_valor = 3 then
			for i_venc = 1 to cd_parcela
				mes_vencimento = mes_vencimento - 1
				ano_vencimento = ano_vencimento
					if mes_vencimento < 1 then
						mes_vencimento = 12
						ano_vencimento = ano_vencimento - 1
					end if
				'response.write("<br>"&mes_vencimento)
			next
			'mes_vencimento = mes_vencimento + 1
		end if
	'************************************
	ano_inicial = request("ano_inicial")
	mes_inicial = request("mes_inicial")
	dt_vencimento = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento
		
		if mes_inicial <> "" Then
			dt_inicial = ""&mes_inicial&"/1/"&ano_inicial&""
		else
			dt_inicial = "'"&mes_vencimento&"/1/"&ano_vencimento&"'"
			dt_final = "NULL"
		end if
		

cd_fornecedor = request("cd_fornecedor")
nm_descricao = request("nm_descricao")
cd_valor = request("cd_valor")
nm_valor = request("nm_valor")
	if nm_valor <> "" Then nm_valor = replace(nm_valor,".","")
	if nm_valor <> "" Then nm_valor = replace(nm_valor,",",".")
	'if nm_valor <> "" Then formatNumber(replace(nm_valor,",","."))
cd_suspenso = request("cd_suspenso")
	if  cd_suspenso <> 1 then
		cd_suspenso = "NULL"
	elseif cd_suspenso = 1 then
		nm_valor = "0.00"
	end if
	
nm_obs = request("nm_obs")


'*** Valor Extra - data final = ultimo dia do mês do pagamento
if cd_diario = "" AND cd_tipo_valor = 1 AND cd_qtd_parcelas <= 1 then
	dt_final = "'"&mes_vencimento&"/"&UltimoDiaMes(mes_vencimento,ano_vencimento)&"/"&ano_vencimento&"'"
end if

'*** Valor parcelado - calcula a data final ***
if cd_tipo_valor = 3 AND cd_qtd_parcelas > 1 Then
	mes_final = mes_vencimento-1
	ano_final = ano_vencimento
	for i_parcelas = 1 to cd_qtd_parcelas
		mes_final = mes_final + 1
		if mes_final > 12 then
			mes_final = 1
			ano_final = ano_final + 1
		end if
		dt_final = "'"&mes_final&"/"&UltimoDiaMes(mes_final,ano_final)&"/"&ano_final&"'"
	next
end if


cd_area = request("cd_area")
cd_centrocusto = request("cd_centrocusto")
cd_conta = request("cd_conta")
cd_unidade = request("cd_unidade")

'**** Dados do cheque ****
lista_pagto = request("lista_pagto")
cd_conta_banco = request("cd_conta_banco")
nm_cheque = request("nm_cheque")
dia_pagamento = request("dia_pagamento")
mes_pagamento = request("mes_pagamento")
ano_pagamento = request("ano_pagamento")
	dt_pagto = mes_pagamento&"/"&dia_pagamento&"/"&ano_pagamento
cd_confirma_pagto = request("cd_confirma_pagto")
modo_pagamento = request("modo_pagamento")
	if modo_pagamento = 2 then
		cd_conta_banco = "NULL"
		nm_cheque = "NULL"
	end if

'**** Dados do faturamento ****
dt_faturamento = request("dt_faturamento")



	'x = 1
	x = ""

IF x = "" THEN
	'**********************************************
	'*** 		Grava no banco de dados 		***
	'**********************************************
	IF acao = "inserir" THEN
		'*** Registra no diário ***
		xsql = "up_financeiro_diario_insert3 @dt_inicial="&dt_inicial&", @dt_final="&dt_final&", @pre_ordem="&dia_vencimento&", @ordem_padrao="&dia_vencimento&", @cd_fornecedor='"&cd_fornecedor&"', @nm_descricao='"&nm_descricao&"', @cd_area='"&cd_area&"', @cd_centrocusto='"&cd_centrocusto&"', @cd_conta='"&cd_conta&"', @cd_unidade='"&cd_unidade&"', @cd_tipo_valor='"&cd_tipo_valor&"', @cd_qtd_parcelas='"&cd_qtd_parcelas&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
				'*** Seleciona o registro recém incluído 
				strsql = "Select * From TBL_financeiro_diario3 Where dt_inicial="&dt_inicial&" AND cd_fornecedor='"&cd_fornecedor&"' AND cd_area='"&cd_area&"' AND cd_centrocusto='"&cd_centrocusto&"' AND cd_conta='"&cd_conta&"' AND cd_unidade='"&cd_unidade&"' AND cd_tipo_valor='"&cd_tipo_valor&"' AND cd_qtd_parcelas='"&cd_qtd_parcelas&"' AND cd_user='"&cd_user&"' AND dt_cad='"&data_atual&"'"
				Set rs_diario = dbconn.execute(strsql)
					if not rs_diario.EOF then	
						cd_diario = rs_diario("cd_codigo")
						response.write("<br>Diario OK <br>")
						'*** insere o valor e o vencimento, conforme a quantidade de parcelas ***
							for i_parc = 1 to cd_qtd_parcelas
								if i_parc = 1 then
									dt_vencimento = dt_vencimento
								elseif i_parc > 1 then
									mes_vencimento = mes_vencimento + 1
									ultimodia = ultimodiames(mes_vencimento,ano_vencimento)
										if mes_vencimento > 12 then
											mes_vencimento = 1
											ano_vencimento = ano_vencimento + 1
										end if
											dt_vencimento = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento
												if dia_vencimento > ultimodia then
													dt_vencimento = mes_vencimento&"/"&ultimodia&"/"&ano_vencimento
												end if
									response.write(dt_vencimento&" - ")
								end if
								'*** Insere o valor do pagamento e data ***
								xsql = "up_financeiro_valor_insert3 @cd_diario="&cd_diario&", @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela='"&i_parc&"', @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_suspenso="&cd_suspenso&", @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
								dbconn.execute(xsql)
							next
					end if
		'response.redirect("../diario_cad3.asp?acao=inserir&mes_sel="&int(mes_vencimento)&"&ano_sel="&ano_vencimento&"")
		
	'**********************************************
	'*** 		Edita o banco de dados			***
	'**********************************************
	ELSEIF acao = "editar" THEN
		'*** Altera os dados do pagamento ***
		xsql = "up_financeiro_diario_update3 @cd_diario='"&cd_diario&"', @nm_descricao='"&nm_descricao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_tipo_valor='"&cd_tipo_valor&"', @ordem_padrao='"&dia_vencimento&"', @cd_qtd_parcelas='"&cd_qtd_parcelas&"', @cd_area='"&cd_area&"', @cd_centrocusto='"&cd_centrocusto&"', @cd_conta='"&cd_conta&"', @cd_unidade='"&cd_unidade&"', @dt_inicial='"&dt_inicial&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
		
	ELSEIF acao = "inserir_valor" THEN
		'*** Insere o valor e data do pagamento já existente***
		xsql = "up_financeiro_valor_insert3 @cd_diario="&cd_diario&", @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela=1, @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_suspenso="&cd_suspenso&", @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
			'*** Altera a pre ordem para que o pagamento se reposicione automaticamente***
			If data_escolhida >= data_atual then
				xsql = "up_financeiro_diario_update_ordem3 @cd_diario='"&cd_diario&"', @pre_ordem="&dia_vencimento&", @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
				dbconn.execute(xsql)
				mens = "valor alterado!"
				response.write("mes_ok_ins")
			end if
		
	ELSEIF acao = "editar_valor" THEN
		'*** Altera o valor e data do pagamento já existente***
		xsql = "up_financeiro_valor_update3 @cd_valor="&cd_valor&", @cd_diario='"&cd_diario&"', @cd_tipo_valor='"&cd_tipo_valor&"', @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela="&cd_parcela&", @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_suspenso="&cd_suspenso&", @dt_final='"&month(dt_vencimento)&"/"&UltimoDiaMes(month(dt_vencimento),year(dt_vencimento))&"/"&year(dt_vencimento)&"',@cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
			'*** Altera a pre ordem para que o pagamento se reposicione automaticamente***
			If data_escolhida >= data_atual then
				xsql = "up_financeiro_diario_update_ordem3 @cd_diario='"&cd_diario&"', @pre_ordem="&dia_vencimento&", @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
				dbconn.execute(xsql)
				mens = "valor alterado!"
				response.write("mes_ok_upd")
			end if
	ELSEIF acao = "edita_obs" THEN
		'*** Altera a observação já existente***
		xsql = "up_financeiro_obs_update3 @cd_valor="&cd_valor&", @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
		
		
	'**************************************************
	'*** 		Insere os dados	do pagamento		***
	'**************************************************
	ELSEIF acao = "insere_cheque" THEN
		'*** registra o cheque  (TBL_)***
			xsql = "up_financeiro_pagamento_insert3 @nm_lista_pagto='"&lista_pagto&"', @cd_conta_banco="&cd_conta_banco&", @nm_cheque="&nm_cheque&", @dt_pagto='"&dt_pagto&"', @cd_modo_pagto='"&modo_pagamento&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
			dbconn.execute(xsql)
				'*** Seleciona o registro recém incluído 
				'strsql = "Select * From TBL_financeiro_banco_pagamentos3 Where nm_lista_pagto='"&lista_pagto&"' AND cd_conta_banco="&cd_conta_banco&" AND nm_cheque="&nm_cheque&""
				strsql = "Select * From TBL_financeiro_banco_pagamentos3 Where nm_lista_pagto='"&lista_pagto&"'"
				Set rs_cheque = dbconn.execute(strsql)
					'if not rs_cheque.EOF Then
						cod_cheque = rs_cheque("cd_codigo")
					'end if
				
					'*** Insere as informações do cheque para pagamento (TBL_)
					pagto_array = split(lista_pagto,",")
						For Each pagto_item In pagto_array
							xsql = "up_financeiro_pagamento_update3 @cd_valor='"&pagto_item&"', @cd_conta_banco="&cd_conta_banco&", @nm_cheque="&nm_cheque&", @cd_cheque="&cod_cheque&", @dt_pagto='"&dt_pagto&"', @cd_modo_pagto='"&modo_pagamento&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
							dbconn.execute(xsql)
						next
	
	ELSEIF acao = "edita_cheque" THEN
	'**************************************************
	'*** 		Insere os dados	do pagamento		***
	'**************************************************
	ELSEIF acao = "excluir_pagamento" THEN
		'*** Exclui a conta a pagar ***
		xsql = "up_financeiro_diario_delete3 @cd_diario='"&cd_diario&"'"
		dbconn.execute(xsql)
			'*** Exclui os valores referentes a conta ***
			xsql = "up_financeiro_valor_delete3 @cd_diario='"&cd_diario&"'"
			dbconn.execute(xsql)
			response.write("Excluído!<br>")
			
	Elseif acao = "fatura" then
	'******************************************************
	'*** 		Insere dados para geração de futura		***
	'******************************************************
		'*** verifica se há registro para o mês selecionado ***
		separa_mes = instr(1,dt_faturamento,"/",1)
		
		mes_fatura = mid(dt_faturamento,1,separa_mes-1)
		
		strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE MONTH(dt_faturamento) = "&mes_fatura&" AND (YEAR(dt_faturamento) = "&year(dt_faturamento)&")"
		Set rs_verif = dbconn.execute(strsql)
			if not rs_verif.EOF then
				cd_fatura = rs_verif("cd_codigo")
				response.write(cd_fatura)
			else 'rs_verif.movenext
				cd_fatura = ""
				response.write(day(dt_faturamento))
			end if
		
		if cd_fatura = "" then
			xsql = "SELECT * FROM TBL_unidades where cd_status = 1 and cd_hospital = 1"
			Set rs = dbconn.execute(xsql)%>
				<%while not rs.EOF
					strcd_unidade = rs("cd_codigo")
					
						qtd_proc1 = request("qtd_proc1_"&strcd_unidade)
						qtd_proc2 = request("qtd_proc2_"&strcd_unidade)
						qtd_proc3 = request("qtd_proc3_"&strcd_unidade)
						qtd_empr = request("qtd_empr_"&strcd_unidade)
						
						dia_vencimento = request("dia_vencimento_"&strcd_unidade)
						nm_obs = request("nm_obs_"&strcd_unidade)
						
						response.write(acao&","&strcd_unidade&","&qtd_proc1&","&qtd_proc2&","&qtd_proc3&","&qtd_empr&",venc:"&dia_vencimento&","&dt_faturamento&","&cd_tipo_proc&"<br>")
						
						xsql = "up_financeiro_fatura_insert3 @cd_unidade="&strcd_unidade&", @dia_vencimento="&dia_vencimento&", @qtd_proc1='"&qtd_proc1&"', @qtd_proc2='"&qtd_proc2&"', @qtd_proc3='"&qtd_proc3&"', @qtd_empr='"&qtd_empr&"', @dt_faturamento='"&dt_faturamento&"', @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
						dbconn.execute(xsql)
						
				rs.movenext
				wend
				response.redirect("../../financeiro.asp?tipo=fatura&cd_unidade="&cd_unidade&"")
		else
			response.write("<br>seria editado o registro <br>")
			xsql = "SELECT * FROM TBL_unidades where cd_status = 1 and cd_hospital = 1"
			Set rs = dbconn.execute(xsql)%>
				<%while not rs.EOF
					strcd_unidade = rs("cd_codigo")
					
						cd_fatura = request("cd_fatura_"&strcd_unidade)
						qtd_proc1 = request("qtd_proc1_"&strcd_unidade)
						qtd_proc2 = request("qtd_proc2_"&strcd_unidade)
						qtd_proc3 = request("qtd_proc3_"&strcd_unidade)
						qtd_empr = request("qtd_empr_"&strcd_unidade)
						
						dia_vencimento = request("dia_vencimento_"&strcd_unidade)
						nm_obs = request("nm_obs_"&strcd_unidade)
						
						response.write(cd_fatura&","&acao&","&strcd_unidade&","&qtd_proc1&","&qtd_proc2&","&qtd_proc3&","&qtd_empr&",venc:"&dia_vencimento&","&dt_faturamento&","&cd_tipo_proc&"<br>")
						
						xsql = "up_financeiro_fatura_update3 @cd_fatura="&cd_fatura&", @dia_vencimento='"&dia_vencimento&"', @qtd_proc1='"&qtd_proc1&"', @qtd_proc2='"&qtd_proc2&"', @qtd_proc3='"&qtd_proc3&"', @qtd_empr='"&qtd_empr&"', @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
						dbconn.execute(xsql)
						
				rs.movenext
				wend
				response.redirect("../../financeiro.asp?tipo=fatura&cd_unidade="&cd_unidade&"")
		end if
	
	END IF

END IF
%>
<body onload="window.close();">
<br>
cd_diario = <%=cd_diario%><br>
cd_valor = <%=cd_valor%><br>
<br>
mens = <%=mens%><br>
acao = <%=acao%><br>

<br>
ordem_padrao = <%=dia_vencimento%><br>
pre_ordem = <%=dia_vencimento%><br><br>

dia_vencimento = <%=dia_vencimento%><br>
mes_vencimento = <%=mes_vencimento%><br>
ano_vencimento = <%=ano_vencimento%><br><br>

dia_final = <%=UltimoDiaMes(mes_final,ano_final)%><br>
mes_final = <%=mes_final%><br>
ano_final = <%=ano_final%><br><br>

cd_suspenso = <%=cd_suspenso%>

cd_fornecedor = <%=cd_fornecedor%><br>
nm_pagamento = <%=nm_pagamento%><br>
nm_obs = <%=nm_obs%><br><br>

nm_valor = <%=nm_valor%><br><br>

cd_area = <%=cd_area%><br>
cd_centrocusto = <%=cd_centrocusto%><br>
cd_conta = <%=cd_conta%><br>
cd_unidade = <%=cd_unidade%><br><br>

cd_qtd_parcelas = <%=cd_qtd_parcelas%><br>
cd_parcela = <%=cd_parcela%><br><br>


*** Vencimento ***: <%=mes_vencimento&"/"&ano_vencimento%><br>

<!--
if cd_tipo_valor = 3 AND cd_qtd_parcelas > 1 Then
		mes_final = mes_vencimento-1
		ano_final = ano_vencimento
		for i_parcelas = 1 to cd_qtd_parcelas
			mes_final = mes_final + 1
			if mes_final > 12 then
				mes_final = 1
				ano_final = ano_final + 1
			end if
			dt_final = "'"&mes_final&"/"&UltimoDiaMes(mes_final,ano_final)&"/"&ano_final&"'"
		next
-->

cd_user = <%=cd_user%><br>
data_atual = <%=data_atual%><br>
nm_obs = <%=nm_obs%><br><br>

lista_pagto = <%=lista_pagto%><br>
cd_conta_banco = <%=cd_conta_banco%><br>
nm_cheque = <%=nm_cheque%><br>
dt_pagto = <%=dt_pagto%><br>
cd_pagamento_agendado = <%=cd_pagamento_agendado%><br>
cd_confirma_pagto = <%=cd_confirma_pagto%><br>
cod_cheque = <%=cod_cheque%><br>
<br>
data_escolhida = <%=data_escolhida%><br>
data_atual= <%=data_atual%><br>

</body>

