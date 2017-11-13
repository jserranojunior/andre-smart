<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->

<%dia_vencimento = request("dia_vencimento")
mes_vencimento = request("mes_vencimento")
ano_vencimento = request("ano_vencimento")
	dt_vencimento = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento

cd_diario = request("cd_diario")

acao = request("acao")
cd_fornecedor = request("cd_fornecedor")
nm_pagamento = request("nm_pagamento")
nm_valor = request("nm_valor")
nm_obs = request("nm_obs")
cd_tipo_pagamento = request("cd_tipo_pagamento")
cd_parcela = request("cd_parcela")
	if cd_parcela = "" then
		cd_parcela = 1
	end if

	if cd_tipo_pagamento = 3 then
		'nm_valor = nm_valor/cd_parcela
	end if
	


cd_area = request("cd_area")
cd_centrocusto = request("cd_centrocusto")
cd_conta = request("cd_conta")
cd_unidade = request("cd_unidade")

cd_user = request("cd_user")
data_atual = request("data_atual")

'**********************************************
'*** 		Grava no banco de dados 		***
'**********************************************
if acao = "incluir" then
	'*** Inclui a despesa ***
	mes_prox = mes_vencimento
	ano_prox = ano_vencimento
	For i_parc = 1 to cd_parcela
		if i_parc = 1 then
			'Parcela 1 será creditada no mes selecionado
			dt_vencimento = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento
		elseif i_parc > 1 then
			'As demais parcelas serão creditadas nos meses sub-sequentes
				mes_seq = i_parc - 1
				
				mes_prox = mes_prox + 1
				response.write(mes_prox)
				if mes_prox > 12 then
					mes_prox = 1
					ano_prox = ano_prox + 1
					
				end if
				
			dt_vencimento = mes_prox&"/"&dia_vencimento&"/"&ano_prox
			
		end if
		xsql = "up_financeiro_despesas_descricao_insert @nm_despesa='"&nm_pagamento&"', @cd_cedente='"&cd_fornecedor&"', @nm_obs='"&nm_obs&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
		dbconn.execute(xsql)
		
			strsql = "Select * From TBL_financeiro_despesas_descricao Where nm_despesa = '"&nm_pagamento&"' AND dt_cad='"&data_atual&"'"
			Set rs = dbconn.execute(strsql)
				if not rs.EOF then	cd_pagamento = rs("cd_codigo")
				
		'*** Registra no diário ***
		xsql = "up_financeiro_diario_insert @cd_area='"&cd_area&"', @cd_centrocusto='"&cd_centrocusto&"', @cd_conta='"&cd_conta&"', @cd_unidade='"&cd_unidade&"', @dt_vencimento='"&dt_vencimento&"',@cd_descricao='"&cd_pagamento&"', @nm_valor='"&nm_valor&"', @cd_tipo_valor='"&cd_tipo_pagamento&"', @cd_parcela='"&i_parc&"', @cd_qtd_parcelas='"&cd_parcela&"', @cd_user = '"&cd_user&"', @data_atual = '"&data_atual&"'"
		dbconn.execute(xsql)
			
			'***resgata o cod do cadastro recem feito ***
			strsql = "Select * From TBL_financeiro_diario Where cd_area='"&cd_area&"', @cd_centrocusto='"&cd_centrocusto&"', @cd_conta='"&cd_conta&"', @cd_unidade='"&cd_unidade&"', @dt_vencimento='"&dt_vencimento&"',@cd_descricao='"&cd_pagamento&"', @nm_valor='"&nm_valor&"', @cd_tipo_valor='"&cd_tipo_pagamento&"'"
			Set rs = dbconn.execute(strsql)
				if not rs.EOF then	cd_pagamento = rs("cd_codigo")
		
		'*** Inclui o valor ***
		xsql = "up_financeiro_valor_insert2 @cd_diario='"&cd_diario&"', @dt_vencimento='"&mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento&"', @nm_valor='"&nm_valor&"', @cd_user = '"&cd_user&"', @data_atual = '"&data_atual&"'"
						dbconn.execute(xsql)
						mens = "inserido"
	next
	
	response.redirect("../diario_cad.asp?action="&action&"&mes_sel="&mes_vencimento&"&ano_sel="&ano_vencimento&"")
	
elseif acao = "editar" Then
	'*** Verifica se já existe valor para a data
		strsql = "Select * From TBL_financeiro_valores Where cd_diario = '"&cd_diario&"' AND dt_vencimento='"&mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento&"'"
		Set rs = dbconn.execute(strsql)
				while not rs.EOF 
					cd_valor = rs("cd_codigo")
				rs.movenext
				wend

					if cd_valor = "" Then ' INSERE VALOR
						xsql = "up_financeiro_valor_insert2 @cd_diario='"&cd_diario&"', @dt_vencimento='"&mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento&"', @nm_valor='"&nm_valor&"', @cd_user = '"&cd_user&"', @data_atual = '"&data_atual&"'"
						dbconn.execute(xsql)
						mens = "inserido"
					else 'ATUALIZA VALORES
						xsql = "up_financeiro_valor_update2 @cd_valor='"&cd_valor&"', @dt_vencimento='"&dt_vencimento&"', @nm_valor='"&nm_valor&"', @cd_user='"&cd_user&"', @data_atual='"&data_atual&"'"
						dbconn.execute(xsql)
						mens = "atualizado"
					end if
end if
%>
<body onload="window.close();">
cd_diario = <%=cd_diario%><br>
cd_valor = <%=cd_valor%><br>
<br>
mens = <%=mens%><br>
<br>

dia_vencimento = <%=dia_vencimento%><br>
mes_vencimento = <%=mes_vencimento%><br>
ano_vencimento = <%=ano_vencimento%><br><br>

cd_fornecedor = <%=cd_fornecedor%><br>
nm_pagamento = <%=nm_pagamento%><br>
nm_obs = <%=nm_obs%><br>
nm_valor = <%=nm_valor%><br><br>


cd_area = <%=cd_area%><br>
cd_centrocusto = <%=cd_centrocusto%><br>
cd_conta = <%=cd_conta%><br>
cd_unidade = <%=cd_unidade%><br><br>

cd_user = <%=cd_user%><br>
data_atual = <%=data_atual%>
</body>

