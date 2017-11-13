<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->

<%cd_diario = request("cd_diario")

dia_vencimento = request("dia_vencimento")
mes_vencimento = request("mes_vencimento")
ano_vencimento = request("ano_vencimento")
	dt_vencimento = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento
		dt_inicial = "'"&mes_vencimento&"/1/"&ano_vencimento&"'"
		dt_final = "NULL"

cd_fornecedor = request("cd_fornecedor")
nm_descricao = request("nm_descricao")
cd_valor = request("cd_valor")
nm_valor = request("nm_valor")
	if nm_valor <> "" Then nm_valor = replace(nm_valor,".","")
	if nm_valor <> "" Then nm_valor = replace(nm_valor,",",".")
	'if nm_valor <> "" Then formatNumber(replace(nm_valor,",","."))
	
nm_obs = request("nm_obs")
cd_qtd_parcelas = request("cd_qtd_parcelas")
	if cd_qtd_parcelas = "" then
		cd_qtd_parcelas = 1
	end if

cd_tipo_valor = request("cd_tipo_valor")
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


'*** Dados do usuário ***
cd_user = request("cd_user")
data_atual = request("data_atual")
acao = request("acao")

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
								if mes_vencimento > 12 then
										mes_vencimento = 1
										ano_vencimento = ano_vencimento + 1
									end if
								dt_vencimento = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento
							end if
							'*** Insere o valor do pagamento e data ***
							xsql = "up_financeiro_valor_insert3 @cd_diario="&cd_diario&", @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela='"&i_parc&"', @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
							dbconn.execute(xsql)
						next
				end if
	response.redirect("../diario_cad3.asp?acao=inserir&mes_sel="&int(mes_vencimento)&"&ano_sel="&ano_vencimento&"")
	
'**********************************************
'*** 		Edita o banco de dados			***
'**********************************************
ELSEIF acao = "editar" THEN
	'*** Altera os dados do pagamento ***
	xsql = "up_financeiro_diario_update3 @cd_diario='"&cd_diario&"', @nm_descricao='"&nm_descricao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_tipo_valor='"&cd_tipo_valor&"', @ordem_padrao='"&dia_vencimento&"', @cd_qtd_parcelas='"&cd_qtd_parcelas&"', @cd_area='"&cd_area&"', @cd_centrocusto='"&cd_centrocusto&"', @cd_conta='"&cd_conta&"', @cd_unidade='"&cd_unidade&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
	dbconn.execute(xsql)
	
ELSEIF acao = "inserir_valor" THEN
	'*** Insere o valor e data do pagamento já existente***
	xsql = "up_financeiro_valor_insert3 @cd_diario="&cd_diario&", @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela=1, @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
	dbconn.execute(xsql)
		'*** Altera a pre ordem para que o pagamento se reposicione automaticamente***
		xsql = "up_financeiro_diario_update_ordem3 @cd_diario='"&cd_diario&"', @pre_ordem="&dia_vencimento&", @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
		mens = "valor alterado!"
	
ELSEIF acao = "editar_valor" THEN
	'*** Altera o valor e data do pagamento já existente***
	xsql = "up_financeiro_valor_update3 @cd_valor="&cd_valor&", @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela=1, @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
	dbconn.execute(xsql)
		'*** Altera a pre ordem para que o pagamento se reposicione automaticamente***
		xsql = "up_financeiro_diario_update_ordem3 @cd_diario='"&cd_diario&"', @pre_ordem="&dia_vencimento&", @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
		mens = "valor alterado!"
	
ELSEIF acao = "edita_obs" THEN
	'*** Altera a observação já existente***
	xsql = "up_financeiro_obs_update3 @cd_valor="&cd_valor&", @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
	dbconn.execute(xsql)
	
ELSEIF acao = "insere_cheque" THEN
	'*** registra o cheque ***
		'xsql = "up_financeiro_cheque_insert3 @nm_lista_pagto='"&lista_pagto&"', @cd_conta_banco="&cd_conta_banco&", @cd_cheque="&cd_cheque&", @dt_pagto='"&dt_pagto&"', @cd_pagamento_agendado='"&dt_pagto&"',cd_confirma_pagto='"&dt_pagto&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		xsql = "up_financeiro_cheque_insert3 @nm_lista_pagto='"&lista_pagto&"', @cd_conta_banco="&cd_conta_banco&", @nm_cheque="&nm_cheque&", @dt_pagto='"&dt_pagto&"',                       @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
			'*** Seleciona o registro recém incluído 
			strsql = "Select * From TBL_financeiro_banco_cheque3 Where nm_lista_pagto='"&lista_pagto&"' AND cd_conta_banco="&cd_conta_banco&" AND nm_cheque="&nm_cheque&""
			Set rs_cheque = dbconn.execute(strsql)
				'if not rs_cheque.EOF Then
					cod_cheque = rs_cheque("cd_codigo")
				'end if
			
			
				'*** Insere as informações do cheque para pagamento
				pagto_array = split(lista_pagto,",")
					For Each pagto_item In pagto_array
						xsql = "up_financeiro_cheque_update3 @cd_valor='"&pagto_item&"', @cd_conta_banco='"&cd_conta_banco&"', @nm_cheque='"&nm_cheque&"', @cd_cheque='"&cod_cheque&"', @dt_pagto='"&dt_pagto&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
						dbconn.execute(xsql)
					next
ELSEIF acao = "edita_cheque" THEN

END IF
%>
<body onload="window.closes();">
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

cd_fornecedor = <%=cd_fornecedor%><br>
nm_pagamento = <%=nm_pagamento%><br>
nm_obs = <%=nm_obs%><br><br>

nm_valor = <%=nm_valor%><br><br>

cd_area = <%=cd_area%><br>
cd_centrocusto = <%=cd_centrocusto%><br>
cd_conta = <%=cd_conta%><br>
cd_unidade = <%=cd_unidade%><br><br>

cd_user = <%=cd_user%><br>
data_atual = <%=data_atual%><br>
nm_obs = <%=nm_obs%><br><br>

lista_pagto = <%=lista_pagto%><br>
cd_conta_banco = <%=cd_conta_banco%><br>
nm_cheque = <%=nm_cheque%><br>
dt_pagto = <%=dt_pagto%><br>
cd_pagamento_agendado = <%=cd_pagamento_agendado%><br>
cd_confirma_pagto = <%=cd_confirma_pagto%><br>
cod_cheque = <%=cod_cheque%>
</body>

