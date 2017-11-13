<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->

<%'*** Dados do usuário ***
cd_user = request("cd_user")
data_atual = request("data_atual")
acao = request("acao")

cd_diario = request("cd_diario")
cd_os = request("cd_os")

cd_atrasadas = request("cd_atrasadas")
'cd_os = request("cd_os")
cd_origem = request("cd_origem")

str_origem = request("cd_origem")
	sep_origem1 = instr(1,str_origem,".",1)
	if sep_origem1 > 0 then
		cd_origem = mid(str_origem,1,sep_origem1)
		codigo_origem = mid(str_origem,sep_origem1, len(str_origem))
			cd_origem = replace(cd_origem,".","")
			codigo_origem = replace(codigo_origem,".","")
	end if

cd_qtd_parcelas = request("cd_qtd_parcelas")
	if cd_qtd_parcelas = "" then
		cd_qtd_parcelas = 1
	end if
cd_qtd_parcelas_atual = request("cd_qtd_parcelas_atual")

cd_parcela = request("cd_parcela")
	if cd_parcela = "" then
		cd_parcela = 1
	end if
cd_tipo_valor = request("cd_tipo_valor")
cd_tipo_valor_atual = request("cd_tipo_valor_atual")

mes_atual = month(now)
ano_atual = year(now)
	data_hoje = ano_atual&zero(mes_atual)
	
dt_vencimento_ori = request("dt_vencimento_ori")

dia_vencimento = request("dia_vencimento")
mes_vencimento = request("mes_vencimento")
ano_vencimento = request("ano_vencimento")

	'if dia_vencimento > ultimodiames(mes_vencimento,ano_vencimento)  then 
	'	dia_vencimento = ultimodiames(mes_vencimento,ano_vencimento)
	'end if
	if int(dia_vencimento) > int(ultimodiames(mes_vencimento,ano_vencimento))  then 
		dia_vencimento = ultimodiames(mes_vencimento,ano_vencimento)
	end if

	data_escolhida = ano_vencimento&mes_vencimento
	dt_vencimento_novo = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento
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

mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

cd_cancela = request("cd_cancela")
if cd_cancela = 1 then '*** Somente em caso de cancelamento da conta a pagar
	mes_final = mes_vencimento - 1
	ano_final = ano_vencimento
		if mes_final < 1 then
			mes_final = 12
			ano_final = ano_final -1
		end if
	dt_final = "'"&mes_final&"/"&UltimoDiaMes(mes_final,ano_final)&"/"&ano_final&"'"
else	'*** verifica e altera ou ratifica a data final
	ano_final = request("ano_final")
	mes_final = request("mes_final")
		if ano_final = "" AND mes_final = "" Then
			dt_final = "NULL"
		else
			dt_final = "'"&mes_final&"/"&UltimoDiaMes(mes_final,ano_final)&"/"&ano_final&"'"
		end if
end if

cd_fornecedor = request("cd_fornecedor")
nm_descricao = request("nm_descricao")'&" - "&cd_os
cd_valor = request("cd_valor")
nm_valor = request("nm_valor")
	if nm_valor <> "" Then nm_valor = replace(nm_valor,".","")
	if nm_valor <> "" Then nm_valor = replace(nm_valor,",",".")
	'if nm_valor <> "" Then formatNumber(replace(nm_valor,",","."))
	
nm_valor_atual = request("nm_valor_atual")
	if nm_valor_atual <> "" Then nm_valor_atual = replace(nm_valor_atual,".","")
	if nm_valor_atual <> "" Then nm_valor_atual = replace(nm_valor_atual,",",".")
	'if nm_valor <> "" Then formatNumber(replace(nm_valor,",","."))
	
cd_pendencia = request("cd_pendencia")

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
	'ultimodia_pagamento = ultimodia_mes(mes_pagamento,ano_pagamento) 
		dt_pagto = mes_pagamento&"/"&dia_pagamento&"/"&ano_pagamento
		
cd_confirma_pagto = request("cd_confirma_pagto")
modo_pagamento = request("modo_pagamento")
	if modo_pagamento = 2 then
		cd_conta_banco = "NULL"
		nm_cheque = "NULL"
	end if
cd_status = request("cd_status")
nm_valor_cheque = request("nm_valor_cheque")

dados_contas = request("dados_contas")

'**** Dados do faturamento ****
dt_faturamento = request("dt_faturamento")

'str_origem = "6.15"

	'x = 1
	'x = ""
	
	'response.write("<br>************************up_financeiro_pagamento_update3 @cd_valor='"&pagto_item&"', <br>@cd_conta_banco="&cd_conta_banco&", <br>@nm_cheque="&nm_cheque&", <br>@cd_cheque="&cod_cheque&", <br>@dt_pagto='"&dt_pagto&"', <br>@cd_modo_pagto='"&modo_pagamento&"', <br>@nm_valor_cheque='"&nm_valor_cheque&"', <br>@cd_user='"&cd_user&"', <br>@data_atual='"&data_atual&"'<br>*********************<br>")

IF x = "" THEN
	'**********************************************
	'*** 		Grava no banco de dados 		***
	'**********************************************
	IF acao = "inserir" THEN
		'*** Registra no diário ***
		xsql = "up_financeiro_diario_insert3 @dt_inicial="&dt_inicial&", @dt_final="&dt_final&", @pre_ordem="&dia_vencimento&", @ordem_padrao="&dia_vencimento&", @cd_fornecedor='"&cd_fornecedor&"', @nm_descricao='"&nm_descricao&"', @cd_area='"&cd_area&"', @cd_centrocusto='"&cd_centrocusto&"', @cd_conta='"&cd_conta&"', @cd_unidade='"&cd_unidade&"', @cd_tipo_valor='"&cd_tipo_valor&"', @cd_qtd_parcelas='"&cd_qtd_parcelas&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"', @cd_origem='"&str_origem&"'"
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
									'if i_parc  > 1 then nm_valor = 0
								end if
								'*** Insere o valor do pagamento e data ***
								xsql = "up_financeiro_valor_insert3 @cd_diario="&cd_diario&", @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela='"&i_parc&"', @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_suspenso="&cd_suspenso&", @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
								dbconn.execute(xsql)
								
									'if cd_os <> "" Then
									'	xsql = "up_andamento_update_pagto2 @cd_os="&cd_os&",@cd_pagto=1"
									'	dbconn.execute(xsql)
									'end if
									
									if codigo_origem <> "" Then
										xsql = "up_orcamento_update_pagto @cd_codigo="&codigo_origem&", @cd_pagto=1"
										dbconn.execute(xsql)
									end if
							next
					end if
		if cd_origem <> "" Then%>
		<body onload="window.close();">
		<%else
			response.redirect("../diario_cad3.asp?acao=inserir&mes_sel="&int(mes_vencimento)&"&ano_sel="&ano_vencimento&"")
		end if




	'**********************************************
	'*** 		Edita o banco de dados			***
	'**********************************************
	ELSEIF acao = "editar" THEN
		'*** Altera os dados do pagamento ***
		xsql = "up_financeiro_diario_update3 @cd_diario='"&cd_diario&"', @nm_descricao='"&nm_descricao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_tipo_valor='"&cd_tipo_valor&"', @ordem_padrao='"&dia_vencimento&"', @cd_qtd_parcelas='"&cd_qtd_parcelas&"', @cd_area='"&cd_area&"', @cd_centrocusto='"&cd_centrocusto&"', @cd_conta='"&cd_conta&"', @cd_unidade='"&cd_unidade&"', @dt_inicial='"&dt_inicial&"', @dt_final="&dt_final&", @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)





	ELSEIF acao = "inserir_valor" THEN
		'*** Insere o valor e data do pagamento já existente***
		xsql = "up_financeiro_valor_insert3 @cd_diario="&cd_diario&", @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela=1, @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_suspenso="&cd_suspenso&", @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
			'*** Altera a pre ordem para que o pagamento se reposicione automaticamente***
			If data_escolhida >= data_hoje then
				xsql = "up_financeiro_diario_update_ordem3 @cd_diario='"&cd_diario&"', @pre_ordem="&dia_vencimento&", @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
				dbconn.execute(xsql)
				mens = "valor alterado!"
				response.write("mes_ok_ins")
			end if





	ELSEIF acao = "editar_valor" THEN
		'*** Altera o valor e data do pagamento já existente***
		xsql = "up_financeiro_valor_update3_1 @cd_valor="&cd_valor&", @cd_diario='"&cd_diario&"', @cd_tipo_valor='"&cd_tipo_valor&"', @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela="&cd_parcela&", @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_suspenso="&cd_suspenso&", @dt_final="&dt_final&",@cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		'xsql = "up_financeiro_valor_update3_1 @cd_valor="&cd_valor&", @cd_diario='"&cd_diario&"', @cd_tipo_valor='"&cd_tipo_valor&"', @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela="&cd_parcela&", @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_suspenso="&cd_suspenso&", @dt_final='"&month(dt_vencimento)&"/"&UltimoDiaMes(month(dt_vencimento),year(dt_vencimento))&"/"&year(dt_vencimento)&"',@cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		''xsql = "up_financeiro_valor_update3 @cd_valor="&cd_valor&", @cd_diario='"&cd_diario&"', @cd_tipo_valor='"&cd_tipo_valor&"', @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela="&cd_parcela&", @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_suspenso="&cd_suspenso&", @dt_final='"&month(dt_vencimento)&"/"&UltimoDiaMes(month(dt_vencimento))&"',@cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
		
			'********************************************************
			'*** Verifica se houve alteração no tipo de pagamento ***
			'********************************************************
			if cd_tipo_valor <> cd_tipo_valor_atual then
				if cd_tipo_valor = 3 then ' Altera para parcelado, calcula e insere as demais parcelas e datas
					'*** verifica as parcelas já existentes ***
					
						
						'if cd_qtd_parcelas_atual = 1 then cd_qtd_parcelas_atual = 2 - 
						
						for i_parc = cd_qtd_parcelas_atual to cd_qtd_parcelas
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
											if int(dia_vencimento) > int(ultimodia) then
												dt_vencimento = mes_vencimento&"/"&ultimodia&"/"&ano_vencimento
											else
												dt_vencimento = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento
											end if
							end if
							'*** Insere o valor do pagamento e data das parcelas***
							ultimo_dia_final = ultimodiames(mes_vencimento,ano_vencimento)
							
							xsql = "up_financeiro_valor_insert3 @cd_diario="&cd_diario&", @nm_pagamento='"&nm_descricao&"', @nm_valor='"&nm_valor&"', @cd_parcela='"&i_parc&"', @dt_vencimento='"&dt_vencimento&"', @nm_obs='"&nm_obs&"', @cd_suspenso="&cd_suspenso&", @cd_user="&cd_user&", @data_atual='"&data&"'"
							dbconn.execute(xsql)
							num_ultima_parc = i_parc
							'response.write("diario="&cd_diario&" - nm_pagamento='"&nm_descricao&"' - valor='"&nm_valor&"' - parcela='"&i_parc&"' - venc='"&dt_vencimento&"' - nm_obs='"&nm_obs&"' - suspenso="&cd_suspenso&" - user="&cd_user&" -  data_atual='"&data&"'<br>")
						next
						
						if cd_cancela = 1 then
						'*** Altera a data final do pagamento - dt_vencimento
						strsql = "UPDATE TBL_financeiro_diario3 SET dt_final='"&mes_vencimento&"/"&ultimodia&"/"&ano_vencimento&"', cd_qtd_parcelas='"&num_ultima_parc&"' WHERE cd_codigo="&cd_diario&""
						'dbconn.execute(strsql)
						end if
						
				'******** FIXO **********
				elseif cd_tipo_valor = 2 then 'Altera Para Fixo e se certifica que a data de témino esteja Null
					if int(cd_tipo_valor_atual) <> int(cd_tipo_valor) then
						strsql = "UPDATE TBL_financeiro_diario3 SET dt_final='12/31/2013', cd_qtd_parcelas=1 WHERE cd_codigo="&cd_diario&""
					else
						strsql = "UPDATE TBL_financeiro_diario3 SET dt_final='12/31/2013', cd_qtd_parcelas=1 WHERE cd_codigo="&cd_diario&""
					end if
					
					dbconn.execute(strsql)
					
						'*** verifica se exisetem parcelas para esse pagamento ***
						strsql = "SELECT * FROM TBL_financeiro_valores3 WHERE cd_diario="&cd_diario&" AND dt_vencimento NOT BETWEEN '"&mes_vencimento&"/1/"&ano_vencimento&"' AND '"&mes_vencimento&"/"&ultimodiames(mes_vencimento,ano_vencimento)&"/"&ano_vencimento&"'"
						SET rs_pagto = dbconn.execute(strsql)
						while not rs_pagto.EOF 
							cd_pagto = rs_pagto("cd_codigo")
							'*** Apaga as parcelas existemtes ***
							strsql = "DELETE FROM TBL_financeiro_valores3 WHERE cd_codigo="&cd_pagto&""
							dbconn.execute(strsql)
							response.write("<br>>>> *** diario="&cd_pagto&"*** <<< <br>")
						rs_pagto.movenext
						wend
					
					
				'******** EXTRA **********	
				elseif cd_tipo_valor = 1 then ''Altera para Extra e define a data de término do pagto para o mês atual
					 response.write("<br>>>>termino:"&dt_vencimento&"<<<<br>")
					 strsql = "UPDATE TBL_financeiro_diario3 SET cd_qtd_parcelas=1, dt_final='"&mes_vencimento&"/"&ultimodiames(mes_vencimento,ano_vencimento)&"/"&ano_vencimento&"' WHERE cd_codigo="&cd_diario&""
					 dbconn.execute(strsql)
					
						'*** verifica se exisetem parcelas para esse pagamento ***
						strsql = "SELECT * FROM TBL_financeiro_valores3 WHERE cd_diario="&cd_diario&" AND dt_vencimento NOT BETWEEN '"&mes_vencimento&"/1/"&ano_vencimento&"' AND '"&mes_vencimento&"/"&ultimodiames(mes_vencimento,ano_vencimento)&"/"&ano_vencimento&"'"
						SET rs_pagto = dbconn.execute(strsql)
						while not rs_pagto.EOF 
							cd_pagto = rs_pagto("cd_codigo")
							'*** Apaga as parcelas existemtes ***
							strsql = "DELETE FROM TBL_financeiro_valores3 WHERE cd_codigo="&cd_pagto&""
							dbconn.execute(strsql)
							response.write("<br>>>> *** diario="&cd_pagto&"*** <<< <br>")
						rs_pagto.movenext
						wend
				end if
				
				'*** verifica se existe pendencia no período ***
				strsql_verpend = "Select * From TBL_financeiro_pendencias where cd_valor='"&cd_valor&"' and dt_inicio_pend <= '"&mes_vencimento&"/1/"&ano_vencimento&"' and dt_fim_pend is null"
				Set rs = dbconn.execute(strsql_verpend)
				if not rs.EOF then
					pend_existe = 1
				end if
				
				if cd_pendencia > 0 AND pend_existe = "" then
					nm_valor = replace(nm_valor,".",",") '*** troca o ponto pela virgula ***
					xsql_pend = "up_financeiro_pendencia_insert @cd_valor="&cd_valor&", @cd_diario="&cd_diario&", @nm_valor='"&nm_valor&"', @dt_inicio_pend='"&mes_vencimento&"/1/"&ano_vencimento&"', @cd_user="&cd_user&", @dt_cad='"&data_atual&"'"
					dbconn.execute(xsql_pend)
					
				end if
				
				
			end if
			
		
			'*** Altera a pre ordem para que o pagamento se reposicione automaticamente***
			If data_escolhida >= data_hoje then
				xsql = "up_financeiro_diario_update_ordem3 @cd_diario='"&cd_diario&"', @pre_ordem="&dia_vencimento&", @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
				'dbconn.execute(xsql)
				mens = "valor alterado!"
				response.write("mes_ok_upd")
			end if





	ELSEIF acao = "edita_obs" THEN
		'*** Altera a observação já existente***
		xsql = "up_financeiro_obs_update3 @cd_valor="&cd_valor&", @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
		
		
	'******************************************************************
	'*** 		Insere os dados	do pagamento (pagamento/transf) 	***
	'******************************************************************
	ELSEIF acao = "insere_cheque" THEN
		'*** Verifica se o cheque  já existe ***
		strsql = "Select * From TBL_financeiro_banco_pagamentos3 Where cd_conta_banco="&cd_conta_banco&" and nm_cheque="&nm_cheque&""
		Set rs_cheque_ver = dbconn.execute(strsql)
			if not rs_cheque_ver.EOF Then
			cheque_existe = rs_cheque_ver("cd_codigo")
			'*** Aqui vai redirecionar ***
			response.redirect("../diario_pagamentos3.asp?cd_conta_banco="&cd_conta_banco&"&nm_cheque="&nm_cheque&"&dia_pagto="&dia_pagamento&"&mes_pagto="&mes_pagamento&"&ano_pagto="&ano_pagamento&"&nm_valor_cheque="&nm_valor_cheque&"&cd_status="&cd_status&"&nm_obs="&nm_obs&"&mens=O_cheque_já_existe")'?tipo=fatura&cd_unidade="&cd_unidade&"&mes_sel="&mes_sel&"&ano_sel="&ano_sel&"")
			else
				'*** registra o cheque  (TBL_)***
				xsql = "up_financeiro_pagamento_insert3 @nm_lista_pagto='"&lista_pagto&"', @cd_conta_banco="&cd_conta_banco&", @nm_cheque="&nm_cheque&", @dt_pagto='"&dt_pagto&"', @cd_modo_pagto='"&modo_pagamento&"', @cd_status="&cd_status&", @nm_valor_cheque='"&nm_valor_cheque&"', @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
				dbconn.execute(xsql)
					'*** Seleciona o registro recém incluído 
					'strsql = "Select * From TBL_financeiro_banco_pagamentos3 Where nm_lista_pagto='"&lista_pagto&"' AND cd_conta_banco="&cd_conta_banco&" AND nm_cheque="&nm_cheque&""
					if lista_pagto <> "" Then
						strsql = "Select * From TBL_financeiro_banco_pagamentos3 Where nm_lista_pagto='"&lista_pagto&"'"
						Set rs_cheque = dbconn.execute(strsql)
							'if not rs_cheque.EOF Then
								cod_cheque = rs_cheque("cd_codigo")
								response.write(cod_cheque)
							'end if
						
							'*** Insere as informações do cheque no registro do pagamento 
							pagto_array = split(lista_pagto,",")
								For Each pagto_item In pagto_array
									nm_valor_cheque = replace(nm_valor_cheque,",",".") 'substitui a virgula por ponto
									xsql = "up_financeiro_pagamento_update3 @cd_valor='"&pagto_item&"', @cd_conta_banco="&cd_conta_banco&", @nm_cheque="&nm_cheque&", @cd_cheque="&cod_cheque&", @dt_pagto='"&dt_pagto&"', @cd_modo_pagto='"&modo_pagamento&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
									dbconn.execute(xsql)
								next
					end if
			end if
	
	ELSEIF acao = "edita_cheque" THEN
	'**************************************************
	'*** 		Insere os dados	do pagamento		***
	'**************************************************
	
	ELSEIF acao = "excluir_pagamento_reg" THEN
	'**************************************************
	'*** 		Exclui os dados	do pagamento		***
	'**************************************************
	'*** Exclui a conta a pagar ***
		xsql = "up_financeiro_diario_delete3 @cd_diario='"&cd_diario&"'"
		dbconn.execute(xsql)
			'*** Exclui os valores referentes a conta ***
			xsql = "up_financeiro_valor_delete3 @cd_diario='"&cd_diario&"'"
			dbconn.execute(xsql)
			response.write("pagamento totalmente Excluído!"&cd_diario&"<br>")
				if codigo_origem <> "" Then
					xsql = "up_orcamento_update_pagto @cd_codigo="&codigo_origem&",@cd_pagto=null"
					dbconn.execute(xsql)
					response.write("+++Orçamento restabelecido!+++<br>")%>
					
				<%end if
				
	ELSEIF acao = "excluir_pagamento_valor" THEN
	'**************************************************
	'*** 		Exclui o valor de um pagamento		***
	'**************************************************
	'*** Exclui um valor de um pagamento***
			'*** Exclui os valores referentes a conta ***
			xsql = "up_financeiro_valor_indiv_delete3 @cd_valor='"&cd_valor&"'"
			dbconn.execute(xsql)
			%><body onload="window.close();">
	<%
	ELSEIF acao = "fatura_2" then
	'******************************************************
	'*** 		Insere dados para geração de futura		***
	'******************************************************
	'*** Verifica se já há registro para a fatura solicitada ***
		strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE MONTH(dt_faturamento) = "&mes_fatura&" AND (YEAR(dt_faturamento) = "&year(dt_faturamento)&")"
		Set rs_verif = dbconn.execute(strsql)
			if not rs_verif.EOF then
				cd_fatura = rs_verif("cd_codigo")
				'response.write("-cheio-")
				response.write("cd_fatura:"&cd_fatura&"<br>")
			else 'rs_verif.movenext
				cd_fatura = ""
				response.write("cd_fatura:-vazio-")
			end if
	
	
	
	ELSEIF acao = "fatura" then
	'******************************************************
	'*** 		Insere dados para geração de futura		***
	'******************************************************
		'*** verifica se há registro para o mês selecionado ***
		response.write("<br>:::FATURA:::<br>")
		separa_mes = instr(1,dt_faturamento,"/",1)
		mes_fatura = mid(dt_faturamento,1,separa_mes-1)
		
		strsql = "SELECT * FROM TBL_financeiro_fatura3 WHERE MONTH(dt_faturamento) = "&mes_fatura&" AND (YEAR(dt_faturamento) = "&year(dt_faturamento)&")"
		Set rs_verif = dbconn.execute(strsql)
			if not rs_verif.EOF then
				cd_fatura = rs_verif("cd_codigo")
				'response.write("-cheio-")
				response.write("cd_fatura:"&cd_fatura&"<br>")
			else 'rs_verif.movenext
				cd_fatura = ""
				response.write("cd_fatura:-vazio-")
			end if
		
	
		'*** insere dados	***
		if cd_fatura = "" then
			xsql = "SELECT * FROM TBL_unidades where cd_status = 1 and cd_hospital = 1"
			Set rs = dbconn.execute(xsql)%>
				<%while not rs.EOF
					strcd_unidade = rs("cd_codigo")
					
					'*** Grava os procedimentos normais ***
					xsql = "SELECT * FROM View_unidades_procedimento_tipo where cd_unidade = "&strcd_unidade&""
					SET rs_proc = dbconn.execute(xsql)
						while not rs_proc.EOF
							cd_procedimento_tipo = rs_proc("cd_procedimento_tipo")
							
							qtd_proc = request("qtd_proc_"&strcd_unidade&"_"&cd_procedimento_tipo)
							if qtd_proc = "" then qtd_proc = 0
							'qtd_empr = request("qtd_empr_"&strcd_unidade)
							cd_tipo_proc = request("cd_tipo_proc_"&strcd_unidade)
					
							dia_vencimento = request("dia_vencimento_"&strcd_unidade)
							dt_recebimento = request("dt_recebimento_"&strcd_unidade&"_0")
								if dt_recebimento <> "" then
									dt_recebimento = "'"&dt_recebimento&"'"
								else
									dt_recebimento = "NULL"
								end if
							nm_obs = request("nm_obs_"&strcd_unidade)
					
							'if qtd_proc <> "" then
								response.write("<br>UNIDADE:"&strcd_unidade&"<br>QTD PROC:"&qtd_proc&"<BR>PROC TIPO:"&cd_procedimento_tipo&"<br>RECEB:"&dt_recebimento&"<br>FAT:"&dt_faturamento&"<br>TIPO PROC:"&cd_tipo_proc&"<br>")
							
								xsql = "up_financeiro_fatura_insert3 @cd_unidade="&strcd_unidade&", @cd_procedimento_tipo="&cd_procedimento_tipo&", @qtd_proc='"&qtd_proc&"', @dt_faturamento='"&dt_faturamento&"', @dt_recebimento="&dt_recebimento&", @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
								dbconn.execute(xsql)
							
						rs_proc.movenext
						wend
					'*** Grava os empréstimos ***
						qtd_empr = request("qtd_empr_"&strcd_unidade)
						dia_vencimento = request("dia_vencimento_"&strcd_unidade)
						dt_recebimento = request("dt_recebimento_"&strcd_unidade&"_0")
							if dt_recebimento <> "" then
									dt_recebimento = "'"&dt_recebimento&"'"
								else
									dt_recebimento = "NULL"
								end if
							'if qtd_empr <> "" then
								'response.write(strcd_unidade&","&qtd_empr&",5,receb:"&dt_recebimento&","&dt_faturamento&"<br>")
							
								'xsql = "up_financeiro_fatura_insert3 @cd_unidade="&strcd_unidade&", @qtd_proc='"&qtd_proc&"', @qtd_proc2='"&qtd_proc2&"', @qtd_proc3='"&qtd_proc3&"', @qtd_empr='"&qtd_empr&"', @dt_faturamento='"&dt_faturamento&"', @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
								xsql = "up_financeiro_fatura_insert3 @cd_unidade="&strcd_unidade&", @cd_procedimento_tipo=5,    @qtd_proc='"&qtd_empr&"', @dt_faturamento='"&dt_faturamento&"', @dt_recebimento="&dt_recebimento&", @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
								dbconn.execute(xsql)
								
							'end if
						'rs_proc.movenext
						'wend
						
						
				rs.movenext
				wend
				response.redirect("../fatura3.asp?tipo=fatura&cd_unidade=0&mes_sel="&mes_sel&"&ano_sel="&ano_sel&"")
				
		'*** EDITA OS REGISTROS ***
		else
			xsql = "SELECT * FROM TBL_unidades where cd_status = 1 and cd_hospital = 1"
			Set rs = dbconn.execute(xsql)%>
				<%while not rs.EOF
					strcd_unidade = rs("cd_codigo")
					
					'*** Grava os procedimentos normais ***
					xsql = "SELECT * FROM View_unidades_procedimento_tipo where cd_unidade = "&strcd_unidade&""
					SET rs_proc = dbconn.execute(xsql)
						while not rs_proc.EOF
							cd_procedimento_tipo = rs_proc("cd_procedimento_tipo")
							
							
							cd_fatura = request("cd_fatura_"&strcd_unidade&"_"&cd_procedimento_tipo)
							qtd_proc = request("qtd_proc_"&strcd_unidade&"_"&cd_procedimento_tipo)
								if qtd_proc = "" Then qtd_proc = 0
							'qtd_empr = request("qtd_empr_"&strcd_unidade)
							'cd_tipo_proc = request("cd_tipo_proc_"&strcd_unidade)
					
							dia_vencimento = request("dia_vencimento_"&strcd_unidade)
								if dt_recebimento = "" Then 
									dt_recebimento = "NULL"
								else
									dt_recebimento = "'"&dt_recebimento&"'"
								end if
							dt_recebimento = request("dt_recebimento_"&strcd_unidade&"_0")
								if dt_recebimento = "" Then 
									dt_recebimento = "NULL"
								else
									dt_recebimento = "'"&dt_recebimento&"'"
								end if
							nm_obs = request("nm_obs_"&strcd_unidade)
					
							'if qtd_proc <> "" then
								response.write("<br>FATURA:"&cd_fatura&"<br>UNID:"&strcd_unidade&"<br>QTD_PROC:"&qtd_proc&"<br>TIPO:"&cd_procedimento_tipo&"<br>RECEB:"&dt_recebimento&"<br>FAT:"&dt_faturamento&"<br>CD TIPO PROC"&cd_tipo_proc&"<br>")
							
								if cd_fatura <> "" Then
									xsql = "up_financeiro_fatura_update3 @cd_fatura="&cd_fatura&", @cd_procedimento_tipo="&cd_procedimento_tipo&", @qtd_proc='"&qtd_proc&"', @dt_faturamento='"&dt_faturamento&"', @dt_recebimento="&dt_recebimento&", @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
									dbconn.execute(xsql)
								end if
							'end if
						rs_proc.movenext
						wend
					'*** Grava os empréstimos ***
						cd_fatura = request("cd_fatura_"&strcd_unidade&"_5")
						qtd_empr = request("qtd_empr_"&strcd_unidade)
						dia_vencimento = request("dia_vencimento_"&strcd_unidade)
						dt_recebimento = request("dt_recebimento_"&strcd_unidade&"_0")
							
							'if qtd_empr <> "" then
								response.write(cd_fatura&","&strcd_unidade&","&qtd_empr&",5,receb:"&dt_recebimento&","&dt_faturamento&"<br><br>")
							
								'xsql = "up_financeiro_fatura_update3 @cd_fatura="&cd_fatura&",@cd_unidade="&strcd_unidade&", @cd_procedimento_tipo=5,    @qtd_proc='"&qtd_empr&"', @dt_faturamento='"&dt_faturamento&"', @dt_recebimento='"&dt_recebimento&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
							if cd_fatura <> "" Then
								xsql = "up_financeiro_fatura_update3 @cd_fatura="&cd_fatura&", @cd_procedimento_tipo=5, @qtd_proc='"&qtd_empr&"', @dt_faturamento='"&dt_faturamento&"', @dt_recebimento='"&dt_recebimento&"', @cd_user="&cd_user&", @data_atual='"&data_atual&"'"
								dbconn.execute(xsql)
							end if	
							'end if
						'rs_proc.movenext
						'wend
						
						
				rs.movenext
				wend
				response.redirect("../fatura3.asp?tipo=fatura&cd_unidade=0&mes_sel="&mes_sel&"&ano_sel="&ano_sel&"")
		end if
	ELSEIF acao = "adiar_recebimento" then
	'**********************************************************************
	'*** 		Informa o atraso no pagamento mensal pelo hospital		***
	'**********************************************************************
	
	'response.write(day(dt_vencimento)&"/"&month(dt_vencimento)&"/"&year(dt_vencimento))
	'response.write("<br>********<br>@cd_unidade="&cd_unidade&", <br>@dt_vencimento_ori='"&dt_vencimento_ori&"', <br>@dt_vencimento_novo='"&dt_vencimento_novo&"', <br>@nm_valor="&nm_valor&",<br>@cd_user="&cd_user&",<br>@dt_cad='"&data_atual&"'<br>********<br>")
	xsql = "up_financeiro_fatura_adia_pgto_insert @cd_unidade="&cd_unidade&", @dt_vencimento_ori='"&dt_vencimento_ori&"', @dt_vencimento_novo='"&dt_vencimento_novo&"', @nm_valor="&nm_valor&",@cd_user="&cd_user&",@dt_cad='"&data_atual&"' "
	dbconn.execute(xsql)
	
	ELSEIF acao = "adiar_rec_editar" then
	'**********************************************************************
	'*** 		Informa o atraso no pagamento mensal pelo hospital		***
	'**********************************************************************
	
	'response.write(day(dt_vencimento)&"/"&month(dt_vencimento)&"/"&year(dt_vencimento))
	'response.write("<br>********<br>@cd_unidade="&cd_unidade&", <br>@dt_vencimento_ori='"&dt_vencimento_ori&"', <br>@dt_vencimento_novo='"&dt_vencimento_novo&"', <br>@nm_valor="&nm_valor&",<br>@cd_user="&cd_user&",<br>@dt_cad='"&data_atual&"'<br>********<br>")
	xsql = "up_financeiro_fatura_adia_pagto_update @cd_codigo="&cd_atrasadas&", @dt_vencimento_novo='"&dt_vencimento_novo&"',@cd_user="&cd_user&",@data_atual='"&data_atual&"' "
	dbconn.execute(xsql)
	response.write("update")
	
	ELSEIF acao = "adiar_rec_excluir" then
	'**********************************************************************
	'*** 		Informa o atraso no pagamento mensal pelo hospital		***
	'**********************************************************************
	
	'response.write(day(dt_vencimento)&"/"&month(dt_vencimento)&"/"&year(dt_vencimento))
	'response.write("<br>********<br>@cd_unidade="&cd_unidade&", <br>@dt_vencimento_ori='"&dt_vencimento_ori&"', <br>@dt_vencimento_novo='"&dt_vencimento_novo&"', <br>@nm_valor="&nm_valor&",<br>@cd_user="&cd_user&",<br>@dt_cad='"&data_atual&"'<br>********<br>")
	xsql = "up_financeiro_fatura_adia_pagto_excluir @cd_codigo="&cd_atrasadas&" "
	dbconn.execute(xsql)
	response.write("excluir"&cd_atrasadas)
	
	
	END IF

END IF
%>
<body onload="window.close();">
<br>
cd_diario = <%=cd_diario%><br>
cd_valor = <%=cd_valor%><br>
cd_orcamento = <%=cd_orcamento%><br>
codigo_origem = <%=codigo_origem%><br>
<br>

<br>

<br>
mens = <%=mens%><br>
acao = <%=acao%><br>

<br>
ordem_padrao = <%=dia_vencimento%><br>
pre_ordem = <%=dia_vencimento%><br><br>

dia_vencimento = <%=dia_vencimento%><br>
mes_vencimento = <%=mes_vencimento%><br>
ano_vencimento = <%=ano_vencimento%><br><br>

cd_cancela = <%=cd_cancela%><br><br>

dia_final = <%'=UltimoDiaMes(mes_final,ano_final)%><br>
mes_final = <%=mes_final%><br>
ano_final = <%=ano_final%><br><br>

cd_suspenso = <%=cd_suspenso%>

cd_fornecedor = <%=cd_fornecedor%><br>
nm_pagamento = <%=nm_pagamento%><br>
nm_obs = <%=nm_obs%><br><br>

nm_valor = <%=nm_valor%><br>
cd_tipo_valor = <%=cd_tipo_valor%><br>
cd_tipo_valor_atual <%=cd_tipo_valor_atual%><br>

cd_area = <%=cd_area%><br>
cd_centrocusto = <%=cd_centrocusto%><br>
cd_conta = <%=cd_conta%><br>
cd_unidade = <%=cd_unidade%><br><br>

cd_qtd_parcelas = <%=cd_qtd_parcelas%><br>
cd_parcela = <%=cd_parcela%><br><br>


*** Vencimento ***: <%=mes_vencimento&"/"&ano_vencimento%><br>
cheque_existe: <%=cheque_existe%><br>

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
status = <%=cd_status%><br><br>

data_escolhida = <%=data_escolhida%><br>
data_atual= <%=data_atual%><br>
str_origem = <%=str_origem%><br>
cd_os = <%=cd_os%><br>

<br>
@cd_valor="&<%=cd_pagto%>&", <br>
@cd_diario="&<%=cd_diario%>&", <br>
@nm_valor='"&<%=nm_valor%>&"', <br>
@dt_inicio_pendencia='"&<%=mes_vencimento&"/1/"&ano_vencimento%>&"', <br>
@cd_user="&<%=cd_user%>&", <br>
@dt_cad='"&<%=data_atual%>&"'"
</body>

