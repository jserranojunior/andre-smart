
  <!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
   <!--#include file="../includes/inc_area_restrita.asp"-->

<%
cd_user = session("cd_codigo")
jan = request("jan")
	if jan = "" Then
		jan = 0
	end if
caminho = request("caminho")
cd_ciente = request("cd_ciente")

acao = request("acao")
cd_tipo = request("cd_tipo")
cd_apaga = request("cd_apaga")
cd_usuario = request("cd_usuario")

cd_codigo = request("cd_codigo")
cd_plan_prev = request("cd_plan_prev")
cd_pat_codigo = request("cd_pat_codigo")
cd_equipamento = request("cd_equipamento")
cd_pn = request("cd_pn")
cd_ns = request("cd_ns")
no_patrimonio = trim(request("no_patrimonio"))
nm_tipo = request("nm_tipo")
cd_especialidade = request("cd_especialidade")
nm_motivo = request("nm_motivo")
cd_nf = request("cd_nf")
	if cd_nf <> "" Then
		cd_nf = "'"&cd_nf&"'"
	else
		cd_nf = "NULL"
	end if

cd_comprador = request("cd_comprador")
cd_unidade_compra = request("cd_unidade_compra")


cd_unidade = request("cd_unidade") '----
cd_rack = request("cd_rack")

cd_unidade_2 = request("cd_unidade_2")
cd_rack_2 = request("cd_rack_2")

cd_unidade_comodato = request("cd_unidade_comodato")
cd_rack_comodato = request("cd_rack_comodato")

	und_movimento_atual = request("und_movimento_atual")
	rack_movimento_atual = request("rack_movimento_atual")

	status_movimento_atual = request("status_movimento_atual")

cd_unit = request("cd_unit")


num_hospital = request("num_hospital")
cd_status = request("cd_status")
cd_marca = request("cd_marca")
cd_categoria = request("cd_categoria")
tipo_manut = request("tipo_manut")
nm_obs = request("nm_obs")

cd_dia = request("cd_dia")
cd_mes = request("cd_mes")
cd_ano = request("cd_ano")

'*****************************************
dt_plan_mes_manut = request("dt_plan_mes_manut")
dt_plan_ano_manut = request("dt_plan_ano_manut")
	dt_plan_manut = "'"&dt_plan_mes_manut&"/1/"&dt_plan_ano_manut&"'"
	
	cd_patrimonio_garantia = request("cd_patrimonio_garantia")
		if cd_patrimonio_garantia = 1 then
			dt_patrimonio_garantia = dt_plan_manut
		else
			dt_patrimonio_garantia = "NULL"
		end if

'*****************************************
cd_preventiva = request("cd_preventiva")
	if cd_preventiva = "" Then
		cd_preventiva = 0
	end if
dt_periodicidade_prev = request("dt_periodicidade_prev")
	if dt_periodicidade_prev = "" Then
		dt_periodicidade_prev = "NULL"
	end if
dt_plan_mes_prev = request("dt_plan_mes_prev")
dt_plan_ano_prev = request("dt_plan_ano_prev")
	dt_plan_prev = "'"&dt_plan_mes_prev&"/1/"&dt_plan_ano_prev&"'"
'*********************************************
cd_calibracao = request("cd_calibracao")
	if cd_calibracao = "" Then
		cd_calibracao = 0
	end if
dt_periodicidade_cal = request("dt_periodicidade_cal")
	if dt_periodicidade_cal = "" Then
		dt_periodicidade_cal = "NULL"
	end if
dt_plan_mes_cal = request("dt_plan_mes_cal")
dt_plan_ano_cal = request("dt_plan_ano_cal")
	dt_plan_cal = "'"&dt_plan_mes_cal&"/1/"&dt_plan_ano_cal&"'"
'*******************************************
cd_seg_elet = request("cd_seg_elet")
	if cd_seg_elet = "" Then
		cd_seg_elet = 0
	end if
dt_periodicidade_elet = request("dt_periodicidade_elet")
	if dt_periodicidade_elet = "" Then
		dt_periodicidade_elet = "NULL"
	end if
dt_plan_mes_elet = request("dt_plan_mes_elet")
dt_plan_ano_elet = request("dt_plan_ano_elet")
	dt_plan_elet = "'"&dt_plan_mes_elet&"/1/"&dt_plan_ano_elet&"'"
'*********************************************
cd_fornecedor_prev = request("cd_fornecedor_prev")
cd_fornecedor_cal = request("cd_fornecedor_cal")
cd_fornecedor_elet = request("cd_fornecedor_elet")
'*********************************************

dt_obs = request("dt_obs")


nao_constar = request("nao_constar")
	if nao_constar = "" Then
		nao_constar = "1"
	end if

dia = day(now)
mes = month(now)
ano = year(now)
cd_hora = hour(now)
cd_minuto = minute(now)
cd_segundo = second(now)
dt_hoje = mes&"/"&dia&"/"&ano&" "&cd_hora&":"&cd_minuto&":"&cd_segundo

	dt_atualizacao = cd_mes&"/"&cd_dia&"/"&cd_ano&" "&cd_hora&":"&cd_minuto&":"&cd_segundo
	dt_data = dt_atualizacao
	dt_descarte = dt_atualizacao

voltar = request("voltar")

if cd_apaga <> "" Then
	acao = "apaga"
end if

'Cargo
cd_funcao = request("cd_funcao")

'Emprestimo
cd_emprestimo = request("cd_emprestimo")

	cd_nova_situacao = request("cd_nova_situacao")
		limite_sitn = instr(1,cd_nova_situacao,":",1)
		limite_sit2n = instr(1,cd_nova_situacao,"{",1)
		limite_sit3n = instr(1,cd_nova_situacao,"}",1)
		limite_sit4n = instr(1,cd_nova_situacao,"[",1)
		
		response.write("*** ---"&cd_nova_situacao&"--- ***<br><br>")
	
	
	cd_situacao = request("cd_situacao")
		limite_sit = instr(1,cd_situacao,":",1)
		limite_sit2 = instr(1,cd_situacao,"{",1)
		limite_sit3 = instr(1,cd_situacao,"}",1)
		limite_sit4 = instr(1,cd_situacao,"[",1)
	
		'*** Nova Situação ***
		'if cd_nova_situacao = 4 then
		if limite_sitn > 1 then
			str_situacao = mid(cd_nova_situacao,1,limite_sitn)
				str_situacao = replace(str_situacao,":","")
			cd_unidade1 = mid(cd_nova_situacao,limite_sitn,limite_sit2n-limite_sitn)
				cd_unidade1 = replace(cd_unidade1,":","")
			cd_rack1 = mid(cd_nova_situacao,limite_sit2n,limite_sit3n-limite_sit2n)
				cd_rack1 = replace(cd_rack1,"{","")
			cd_unidade2 = mid(cd_situacao,limite_sit3n,limite_sit4n-limite_sit3n)
				cd_unidade2 = replace(cd_unidade2,"}","")
			cd_rack2 = mid(cd_situacao,limite_sit4n,len(cd_situacao))
				cd_rack2 = replace(cd_rack2,"[","")
		
		'*** Situação ***
		'elseif cd_situacao <> "" Then
		elseif limite_sit > 1 Then
			str_situacao = mid(cd_situacao,1,limite_sit)
				str_situacao = replace(str_situacao,":","")
			cd_unidade1 = mid(cd_situacao,limite_sit,limite_sit2-limite_sit)
				cd_unidade1 = replace(cd_unidade1,":","")				
			cd_rack1 = mid(cd_situacao,limite_sit2,limite_sit3-limite_sit2)
				cd_rack1 = replace(cd_rack1,"{","")
			cd_unidade2 = mid(cd_situacao,limite_sit3,limite_sit4-limite_sit3)
				cd_unidade2 = replace(cd_unidade2,"}","")				
			cd_rack2 = mid(cd_situacao,limite_sit4,len(cd_situacao))
				cd_rack2 = replace(cd_rack2,"[","")
		end if
		
		
dia_movimento_obj = request("dia_movimento_obj")
mes_movimento_obj = request("mes_movimento_obj")
ano_movimento_obj = request("ano_movimento_obj")
	dt_movimento_obj = "'"&mes_movimento_obj&"/"&dia_movimento_obj&"/"&ano_movimento_obj&"'"

und_destino_obj = request("und_destino_obj")
nm_solicitante = request("nm_solicitante")
nm_observacoes = request("nm_observacoes")
rack_destino_obj = request("rack_destino_obj")

'*** Movimentações do patrimônio ***
	if und_destino_obj <> "" Then
		cd_unidade_destino = und_destino_obj
		'cd_unidade_comodato = und_destino_obj
		'response.write("unidade:"&cd_unidade_comodato&":"&und_destino_obj)
	end if
	
	if 	rack_destino_obj <> "" Then
		cd_rack_destino = rack_destino_obj
		'cd_rack_comodato = rack_destino_obj
	end if
	
'*** devolução 
	if str_situacao = 8 AND xpto = 1 then
		cd_unidade_comodato = und_movimento_atual
		cd_rack_comodato = rack_movimento_atual
	end if
'*** Campos desativados *********************		

cd_unidade_des = request("cd_unidade_des")
cd_rack_des = request("cd_rack_des")

nm_solicitante = request("nm_solicitante")
nm_empr_obs = request("nm_empr_obs")

dia_emprestimo = request("dia_emprestimo")
mes_emprestimo = request("mes_emprestimo")
ano_emprestimo = request("ano_emprestimo")
	IF dia_emprestimo <> "" AND mes_emprestimo <> "" AND ano_emprestimo <> "" THEN
		dt_emprestimo = "'"&mes_emprestimo&"/"&dia_emprestimo&"/"&ano_emprestimo&"'"
	ELSE
		dt_emprestimo = "NULL"
	END IF
	
	
dia_estoque = request("dia_estoque")
mes_estoque = request("mes_estoque")
ano_estoque = request("ano_estoque")
	dt_estoque = mes_estoque&"/"&dia_estoque&"/"&ano_estoque
	
dia_descarte = request("dia_descarte")
mes_descarte = request("mes_descarte")
ano_descarte = request("ano_descarte")
	dt_descarte = mes_descarte&"/"&dia_descarte&"/"&ano_descarte
	

cd_unidade_des_novo = request("cd_unidade_des_novo")
cd_rack_des_novo = request("cd_rack_des_novo")
dia_emprestimo_novo = request("dia_emprestimo_novo")
mes_emprestimo_novo = request("mes_emprestimo_novo")
ano_emprestimo_novo = request("ano_emprestimo_novo")
	dt_emprestimo_novo = mes_emprestimo_novo&"/"&dia_emprestimo_novo&"/"&ano_emprestimo_novo
nm_empr_obs_novo = request("nm_empr_obs_novo")


dia_empr_devolucao = request("dia_empr_devolucao")
mes_empr_devolucao = request("mes_empr_devolucao")
ano_empr_devolucao = request("ano_empr_devolucao")
	dt_empr_devolucao = mes_empr_devolucao&"/"&dia_empr_devolucao&"/"&ano_empr_devolucao
	response.write("<br>\\\\\\\\\\\\\\\"&dt_empr_devolucao&"////////<br>")
		'if dia_empr_devolucao = "" Then
		'	dia_empr_devolucao = 0
		'end if

dia_efetivacao = request("dia_efetivacao")
mes_efetivacao = request("mes_efetivacao")
ano_efetivacao = request("ano_efetivacao")
	dt_efetivacao = mes_efetivacao&"/"&dia_efetivacao&"/"&ano_efetivacao
		'if dia_efetivacao = "" Then
		'	dia_efetivacao = 0
		'end if
		
cd_cancela = request("cd_cancela")
cd_mov_atual = request("cd_mov_atual")

'***************************************************************************************************************************
	'******************************
	'***	Unidades e racks.	***
	'******************************
	
	if str_situacao = "1" AND cd_unidade <> "" AND cd_unidade_des <> "" AND dia_empr_devolucao = "" then ' Descarta
		cd_unidade_comodato = 33' request("cd_unidade_des")
		cd_rack_comodato = 11' request("cd_rack_des")
		erro = "Descarta"
		
	elseif str_situacao = "2" then ' Manutenção
		
	
	elseif str_situacao = "3" then 'Aquisição
		
	
	elseif str_situacao = "4" then 'Emprestimo
		
	
	elseif str_situacao = "5" then ' inicia comodato
		cd_unidade_comodato = und_destino_obj 'request("cd_unidade_des")
		cd_rack_comodato = rack_destino_obj 'request("cd_rack_des")
		erro = "inclusao"
	
	elseif str_situacao = "6" then ' Estoque
		cd_unidade_des = "29" '= cod estoque
		cd_rack_des = "11" '= sem rack
		cd_unidade_comodato = und_destino_obj'cd_unidade_des
		cd_rack_comodato = rack_destino_obj 'cd_rack_des
		erro = "Estoque"
	
	elseif str_situacao = "7" then ' Troca o Rack
		cd_unidade_comodato = und_destino_obj 'request("cd_unidade_des")
		cd_rack_comodato = rack_destino_obj 'request("cd_rack_des")
		erro = "Rack"
	
	elseif str_situacao = "9" then ' Efetiva o Comodato
		erro = "Rack"
	
	Else
		Erro = "ok"
	
	end if
	
	
'*****************************
'*** INSERE AS OCORRENCIAS ***
'*****************************%>
<!--#include file="../ocorrencia/acoes/ocorrencias_acao.asp"-->

<%'="up_patrimonio_update<br>@cd_pat_codigo="&cd_pat_codigo&", <br>@cd_equipamento='"&cd_equipamento&"', <br>@cd_marca='"&cd_marca&"', <br>@cd_pn='"&cd_pn&"', <br>@cd_ns='"&cd_ns&"',<br>@no_patrimonio='"&no_patrimonio&"',<br>@nm_tipo='"&nm_tipo&"', <br>@cd_especialidade='"&cd_especialidade&"', <br>@cd_unidade='"&cd_unidade&"', <br>@cd_rack='"&cd_rack&"', <br>@cd_unidade_comodato='"&cd_unidade_comodato&"', <br>@cd_rack_comodato='"&cd_rack_comodato&"', <br>@num_hospital='"&num_hospital&"', <br>@cd_categoria='"&cd_categoria&"' ,<br>@dt_data='"&dt_data&"', <br>@cd_nf="&cd_nf&",<br>@cd_preventiva="&cd_preventiva&", <br>@dt_periodicidade_prev="&dt_periodicidade_prev&", <br>@cd_calibracao="&cd_calibracao&", <br>@dt_periodicidade_cal="&dt_periodicidade_cal&", <br>@cd_seg_elet="&cd_seg_elet&", <br>@dt_periodicidade_elet="&dt_periodicidade_elet&", <br>@nao_constar="&nao_constar&""%><br>
<br>

<%x = ""
IF x = "" then
' ****  AÇÕES  ****
	'Patrimonio
	existe = 0
	
'******************************************
'***		INSERE O PATRIMÔNIO			***
'******************************************
	If acao = "inserir" AND cd_tipo = "patrimonio" Then
		'*** Verifica se o patrimonio já existe ***
		strsql ="SELECT * FROM View_patrimonio_lista WHERE no_patrimonio = "&no_patrimonio&" AND nm_tipo='"&nm_tipo&"'"
		Set rs_patrimonio = dbconn.execute(strsql)
		while not rs_patrimonio.EOF
			'aviso = "Este patrimonio já existe"
			existe = 1
		rs_patrimonio.movenext
		wend
			if existe > 0 and cd_ciente = 1 then
				existe = 0
			end if
	
		IF existe = 0  THEN
			'Insere dados
			xsql = "up_patrimonio_insert @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @no_patrimonio='"&no_patrimonio&"', @nm_tipo='"&nm_tipo&"', @cd_especialidade='"&cd_especialidade&"', @cd_unidade='"&cd_unidade&"', @cd_rack='"&cd_rack&"', @num_hospital='"&num_hospital&"', @cd_categoria='"&cd_categoria&"', @dt_data='"&dt_data&"', @cd_nf="&cd_nf&", @cd_preventiva="&cd_preventiva&", @dt_periodicidade_prev="&dt_periodicidade_prev&", @cd_fornecedor_prev="&cd_fornecedor_prev&", @cd_calibracao="&cd_calibracao&", @dt_periodicidade_cal="&dt_periodicidade_cal&", @cd_fornecedor_cal="&cd_fornecedor_cal&", @cd_seg_elet="&cd_seg_elet&", @dt_periodicidade_elet="&dt_periodicidade_elet&",@cd_fornecedor_elet="&cd_fornecedor_elet&", @nao_constar="&nao_constar&""
			dbconn.execute(xsql)
			response.write("patrimonio insere OK")
				'*** Recupera o codigo recem criado e agenda as manutenções***
				xsql = "select * from TBL_patrimonio where no_patrimonio='"&no_patrimonio&"' AND  cd_equipamento='"&cd_equipamento&"' AND nm_tipo='"&nm_tipo&"'"
				Set rs_patr = dbconn.execute(xsql)
				
				cd_pat_codigo = rs_patr("cd_codigo")
				
					'*** PREVENTIVA ***
					'*** Calcula a data da primeira manutenção preventiva a partir da data de aquisição ***
					if dt_periodicidade_prev <> "" AND cd_preventiva  <> "0" Then
						dt_plan_prev_dia = cd_dia'day(dt_data)
						dt_plan_prev_mes = cd_mes'month(dt_data)
						dt_plan_prev_ano = cd_ano'year(dt_data)
							mes_prev = int(dt_periodicidade_prev) + int(dt_plan_prev_mes)
							ano_prev = dt_plan_prev_ano
							if int(mes_prev) > 12 then 'Virada de ano
								while not mes_prev <= 12
									mes_prev = mes_prev - 12
									ano_prev = ano_prev + 1
								wend
							else
								mes_prev = 12 - mes_prev
								ano_prev = dt_plan_prev_ano + 1
							end if
								dt_plan_prev = "'"&mes_prev&"/"&dt_plan_prev_dia&"/"&ano_prev&"'"
								dt_patrimonio_garantia_prev = "'"&dt_plan_prev_mes&"/"&dt_plan_prev_dia&"/"&dt_plan_prev_ano&"'"
							xsql = "up_patrimonio_preventiva_insert @tipo_manut=1, @cd_patrimonio="&cd_pat_codigo&", @cd_suspenso='0', @dt_plan_prev="&dt_plan_prev&",@dt_garantia="&dt_patrimonio_garantia_prev&""
							dbconn.execute(xsql)
					end if
					
					'*** CALIBRAÇÃO ***
					'*** Calcula a data da primeira calibração a partir da data de aquisição ***
					if dt_periodicidade_cal <> "" AND cd_calibracao  <> "0" Then
						dt_plan_cal_dia = cd_dia
						dt_plan_cal_mes = cd_mes
						dt_plan_cal_ano = cd_ano
							mes_cal = int(dt_periodicidade_cal) + int(dt_plan_cal_mes)
							ano_cal = dt_plan_cal_ano
							if int(mes_cal) > 12 then 'Virada de ano
								do while not mes_cal <= 12
									mes_cal = mes_cal - 12
									ano_cal = ano_cal + 1
								loop
							else 'Mesmo ano
								mes_cal = 12 - mes_cal
								ano_cal = dt_plan_cal_ano + 1
							end if
								dt_plan_cal = "'"&mes_cal&"/"&dt_plan_cal_dia&"/"&ano_cal&"'"
								dt_patrimonio_garantia_cal = "'"&dt_plan_cal_mes&"/"&dt_plan_cal_dia&"/"&dt_plan_cal_ano&"'"
							xsql = "up_patrimonio_preventiva_insert @tipo_manut=2, @cd_patrimonio="&cd_pat_codigo&", @cd_suspenso='0', @dt_plan_prev="&dt_plan_cal&",@dt_garantia="&dt_patrimonio_garantia_cal&""
							dbconn.execute(xsql)
					end if
					
					'*** SEG. ELETRICA ***
					'*** Calcula a data da primeira calibração a partir da data de aquisição ***
					if dt_periodicidade_elet <> "" AND cd_seg_elet  <> "0" Then
						dt_plan_elet_dia = cd_dia
						dt_plan_elet_mes = cd_mes
						dt_plan_elet_ano = cd_ano
							mes_elet = int(dt_periodicidade_elet) + int(dt_plan_elet_mes)
							ano_elet = dt_plan_elet_ano
							if int(mes_elet) > 12 then 'Virada de ano
								do while not mes_elet <= 12
									mes_elet = mes_elet - 12
									ano_elet = ano_elet + 1
								loop
							else 'Mesmo ano
								mes_elet = 12 - mes_elet
								ano_elet = dt_plan_elet_ano + 1
							end if
								dt_plan_elet = "'"&mes_elet&"/"&dt_plan_elet_dia&"/"&ano_elet&"'"
								dt_patrimonio_garantia_elet = "'"&dt_plan_elet_mes&"/"&dt_plan_elet_dia&"/"&dt_plan_elet_ano&"'"
							xsql = "up_patrimonio_preventiva_insert @tipo_manut=3, @cd_patrimonio="&cd_pat_codigo&", @cd_suspenso='0', @dt_plan_prev="&dt_plan_elet&",@dt_garantia="&dt_patrimonio_garantia_elet&""
							dbconn.execute(xsql)
					end if
					
					'*** MOVIMENTAÇÃO ***
					'*** Grava a empresa da Holding vinculada à unidade de negócio ***
					'XXX Seleciona a empresa da holding vinculada à unidade de negócio XXX
					'*** busca a empresa fornecedora de serviços ***
					strsql ="SELECT * FROM TBL_unidades WHERE cd_codigo = "&cd_unidade&""
					Set rs_und = dbconn.execute(strsql)
						if not rs_und.EOF then
							cd_empresa = rs_und("cd_cliente_empresa")
						end if
					
					id_movimento = 1 '*** indica o primeiro movimento do recém cadastrado patrimônio ***
					'xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_empresa&"', @cd_unidade_destino='"&cd_unidade&"', @cd_tipo_movimentacao=3, @dt_ida='"&dt_data&"', @dt_retorno='"&dt_data&"', @id_movimento="&id_movimento&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
					xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='1', @cd_rack_origem=NULL, @cd_unidade_destino='"&cd_unidade&"', @cd_rack_destino='"&cd_rack&"', @cd_tipo_movimentacao=NULL, @dt_ida=null, @dt_retorno=null, @id_movimento='1', @cd_natureza_mov='1', @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
					dbconn.execute(xsql)
					
			if jan <> 1 then
				response.redirect("../../patrimonio.asp?tipo=cadastro&mensagem=Inserido com sucesso")
			end if
		ELSE
			if jan <> 1 then
				response.redirect("../../patrimonio.asp?tipo=cadastro&caminho=patrimonio&jan=0&existe=1&no_patrimonio="&no_patrimonio&"&nm_tipo="&nm_tipo&"&cd_equipamento="&cd_equipamento&"&cd_marca="&cd_marca&"&cd_pn="&cd_pn&"&cd_ns="&cd_ns&"&cd_especialidade="&cd_especialidade&"&cd_unidade="&cd_unidade&"&cd_rack="&cd_rack&"&num_hospital="&num_hospital&"&cd_categoria="&cd_categoria&"&dt_data="&month(dt_data)&"/"&day(dt_data)&"/"&year(dt_data)&"&cd_nf="&cd_nf&"&cd_preventiva="&cd_preventiva&"&dt_periodicidade_prev="&dt_periodicidade_prev&"&cd_calibracao="&cd_calibracao&"&dt_periodicidade_cal="&dt_periodicidade_cal&"&cd_seg_elet="&cd_seg_elet&"&dt_periodicidade_elet="&dt_periodicidade_elet&"&nao_constar="&nao_constar&"")
			else
				response.redirect("../patrimonio_cad.asp?tipo=cadastro&caminho=patrimonio&jan=1&existe=1&no_patrimonio="&no_patrimonio&"&nm_tipo="&nm_tipo&"&cd_equipamento="&cd_equipamento&"&cd_marca="&cd_marca&"&cd_pn="&cd_pn&"&cd_ns="&cd_ns&"&cd_especialidade="&cd_especialidade&"&cd_unidade="&cd_unidade&"&cd_rack="&cd_rack&"&num_hospital="&num_hospital&"&cd_categoria="&cd_categoria&"&dt_data="&month(dt_data)&"/"&day(dt_data)&"/"&year(dt_data)&"&cd_nf="&cd_nf&"&cd_preventiva="&cd_preventiva&"&dt_periodicidade_prev="&dt_periodicidade_prev&"&cd_calibracao="&cd_calibracao&"&dt_periodicidade_cal="&dt_periodicidade_cal&"&cd_seg_elet="&cd_seg_elet&"&dt_periodicidade_elet="&dt_periodicidade_elet&"&nao_constar="&nao_constar&"")
			end if
		END IF
		
		
'******************************
'***		EDITA			***
'******************************		
	ElseIf acao = "editar" AND cd_tipo = "patrimonio" Then
		'Modifica os dados
		'xsql = "up_patrimonio_update @cd_pat_codigo="&cd_pat_codigo&", @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @no_patrimonio='"&no_patrimonio&"', @nm_tipo='"&nm_tipo&"', @cd_especialidade='"&cd_especialidade&"', @cd_unidade='"&cd_unidade&"', @cd_rack='"&cd_rack&"', @num_hospital='"&num_hospital&"', @cd_categoria='"&cd_categoria&"' ,@dt_data='"&dt_data&"', @cd_preventiva="&cd_preventiva&", @dt_periodicidade_prev="&dt_periodicidade_prev&", @cd_calibracao="&cd_calibracao&", @dt_periodicidade_cal="&dt_periodicidade_cal&", @cd_seg_elet="&cd_seg_elet&", @dt_periodicidade_elet="&dt_periodicidade_elet&", @nao_constar="&nao_constar&""
		xsql = "up_patrimonio_update @cd_pat_codigo="&cd_pat_codigo&", @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @no_patrimonio='"&no_patrimonio&"', @nm_tipo='"&nm_tipo&"', @cd_especialidade='"&cd_especialidade&"', @cd_unidade='"&cd_unidade&"', @cd_rack='"&cd_rack&"', @cd_unidade_comodato='"&cd_unidade_comodato&"', @cd_rack_comodato='"&cd_rack_comodato&"', @num_hospital='"&num_hospital&"', @cd_categoria='"&cd_categoria&"' ,@dt_data='"&dt_data&"', @cd_nf="&cd_nf&",@cd_preventiva="&cd_preventiva&", @dt_periodicidade_prev="&dt_periodicidade_prev&", @cd_fornecedor_prev="&cd_fornecedor_prev&", @cd_calibracao="&cd_calibracao&", @dt_periodicidade_cal="&dt_periodicidade_cal&", @cd_fornecedor_cal="&cd_fornecedor_cal&", @cd_seg_elet="&cd_seg_elet&", @dt_periodicidade_elet="&dt_periodicidade_elet&", @cd_fornecedor_elet="&cd_fornecedor_elet&", @nao_constar="&nao_constar&""
		dbconn.execute(xsql)
		response.write("patrimonio edita OK<br>")
		response.write(cd_unidade_comodato&"<br>")
		
			'*** Planejamento Preventiva ***
			if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
				xsql = "up_patrimonio_preventiva_insert @tipo_manut=1, @cd_patrimonio='"&cd_pat_codigo&"', @cd_suspenso='"&cd_suspenso&"', @dt_plan_prev="&dt_plan_prev&",@dt_garantia="&dt_patrimonio_garantia&""
				dbconn.execute(xsql)
				response.write("/////<br>preventiva, Hooooo!<br>\\\\\\\")
			end if
			
			'*** Planejamento Calibração ***
			if dt_plan_mes_cal <> "" AND dt_plan_ano_cal <> "" AND cd_calibracao <> "0" Then
				xsql = "up_patrimonio_preventiva_insert @tipo_manut=2, @cd_patrimonio='"&cd_pat_codigo&"', @cd_suspenso='"&cd_suspenso&"', @dt_plan_prev="&dt_plan_cal&",@dt_garantia="&dt_patrimonio_garantia&""
				dbconn.execute(xsql)
			end if
			
			'*** Planejamento Segurança Eletrica ***
			if dt_plan_mes_elet <> "" AND dt_plan_ano_elet <> "" AND cd_seg_elet <> "0" Then
				xsql = "up_patrimonio_preventiva_insert @tipo_manut=3, @cd_patrimonio='"&cd_pat_codigo&"', @cd_suspenso='"&cd_suspenso&"', @dt_plan_prev="&dt_plan_elet&", @dt_garantia="&dt_patrimonio_garantia&""
				dbconn.execute(xsql)
			end if
			
			
			'*** Verifica se há o primeiro movimento do patrimonio ***
			'strsql = "SELECT * FROM TBL_patrimonio_movimentacao where cd_patrimonio='"&cd_pat_codigo&"' AND id_movimento = 1"
			'set rs = dbconn.execute(strsql)
			'	if not rs.EOF then
			'		existe_mov = 1
			'	end if
			'		if existe_mov < 1 then '*** Grava o primeiro movimento = Empresa vinculada à unidade ***
			'			'*** MOVIMENTAÇÃO ***
			'			'*** Grava a empresa da Holding vinculada à unidade de negócio ***
			'			cd_empresa = 1'XXX -> Seleciona a empresa da holding vinculada à unidade de negócio XXX - *** NECESSITA ALTERAÇÃO ***
			'			id_movimento = 1 '*** indica o primeiro movimento do recém cadastrado patrimônio ***
			'			xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_empresa&"', @cd_unidade_destino='"&cd_empresa&"', @cd_tipo_movimentacao=3, @dt_ida='"&dt_data&"', @dt_retorno='"&dt_data&"', @id_movimento="&id_movimento&", @cd_user="&cd_user&", @dt_cadastro='"&dt_hoje&"'"
			'			dbconn.execute(xsql)
			'		end if
			'strsql = "UPDATE TBL_patrimonio_movimentacao SET cd_unidade_origem="&cd_comprador&"  cd_unidade_destino="&cd_unidade_compra&" WHERE (cd_patrimonio = "&cd_pat_codigo&") AND (id_movimento = 1)"
			'response.write("cd_unidade_compra:"&cd_unidade_compra&"<br> cd_unidade_origem:"&cd_unidade_origem&"<br> cd_unidade_destino:"&cd_unidade_destino)
			
			if cd_unidade_origem <> "" AND cd_unidade_destino <> "" AND str_situacao <> 4 Then
				strsql = "UPDATE up_patrimonio_movimentacao_update @cd_unidade_origem="&cd_comprador&",  @cd_unidade_destino="&cd_unidade_compra&", @cd_patrimonio = "&cd_pat_codigo&""
				Set rs = dbconn.execute(strsql)
				response.write("<br>** movimenta a unidade.**<br>")
			end if
			
			if str_situacao = 7 Then '*** 7 - Movimentação Rack para Rack***
				'strsql = "UPDATE up_patrimonio_movimentacao_update @cd_unidade_origem="&cd_comprador&",  @cd_unidade_destino="&cd_unidade_compra&", @cd_patrimonio = "&cd_pat_codigo&""
				'Set rs = dbconn.execute(strsql)
				action = "altera Rack"
			end if
			
		if jan <> 1 then
			'response.redirect("../../patrimonio.asp?tipo=cadastro&mensagem=Alterado com sucesso")
		end if
		
		
'******************************
'***	APAGA PATRIMONIO	***
'******************************
	Elseif acao = "apaga" AND cd_tipo = "patrimonio" Then
		'Apaga os dados
		xsql = "up_patrimonio_delete @cd_apaga='"&cd_apaga&"'"
		dbconn.execute(xsql)
		response.write("patrimonio apaga OK")
		
			'*** Apaga todas as manutenções referentes ao patrimonio que está sendo apagado ***
			xsql = "up_patrimonio_preventiva_delete_todas @cd_patrimonio='"&cd_apaga&"'"
			dbconn.execute(xsql)
			'*** Apaga toda movimentação referente ao patrimonio que está sendo apagado ***
			xsql = "up_patrimonio_emprestimo_delete_todas @cd_patrimonio='"&cd_apaga&"'"
			dbconn.execute(xsql)
				'*** Apaga o histórico de movimentação ***
				xsql = "up_patrimonio_movimentacao_delete_todas @cd_patrimonio='"&cd_apaga&"'"
			dbconn.execute(xsql)
			'*************************************************
			reload = "1"
		if jan <> 1 then
			response.redirect("../../patrimonio.asp?tipo=cadastro&mensagem=Apagado com sucesso")
		end if
		
		
'**********************************
'***	MANUTENÇÃO PREVENTIVA	***
'**********************************
	Elseif acao = "inserir" and cd_tipo = "manutencao" then
	
		'if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
			xsql = "up_patrimonio_preventiva_insert @cd_patrimonio="&cd_plan_prev&", @dt_plan_prev="&dt_plan_manut&", @dt_garantia="&dt_patrimonio_garantia&", @tipo_manut="&tipo_manut&",@cd_suspenso='0'"
			'response.write("<br> aqui editamos a manutenção dos equip.<br>")
			dbconn.execute(xsql)
		'end if
	
	Elseif acao = "editar" and cd_tipo = "manutencao" then
	
		'if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
			xsql = "up_patrimonio_preventiva_update @cd_patrimonio_preventiva='"&cd_plan_prev&"',  @dt_plan_prev="&dt_plan_manut&", @dt_garantia="&dt_patrimonio_garantia&""
			dbconn.execute(xsql)
		'end if
	
	Elseif acao = "apagar" and cd_tipo = "manutencao" then
	
		'if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
			'xsql = "up_patrimonio_preventiva_update @cd_patrimonio_preventiva='"&cd_plan_prev&"',  @dt_plan_prev="&dt_plan_manut&""
			xsql = "DELETE FROM TBL_patrimonio_manutencoes WHERE cd_codigo = '"&cd_plan_prev&"'"
			dbconn.execute(xsql)
		'end if
		action = "erro"
	
	Elseif acao = "inserir" and cd_tipo = "manut_obs" then
				'if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
					xsql = "up_patrimonio_manut_obs_insert @dt_obs='"&month(dt_obs)&"/"&day(dt_obs)&"/"&year(dt_obs)&"', @nm_obs='"&nm_obs&"', @tipo_manut="&tipo_manut&", @cd_unidade='"&cd_unidade&"', @cd_user="&cd_user&", @dt_cadastro='"&dt_hoje&"'"
					dbconn.execute(xsql)
				'end if
	Elseif acao = "editar" and cd_tipo = "manut_obs" then
					'if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
					xsql = "up_patrimonio_manut_obs_update @cd_obs='"&cd_codigo&"', @nm_obs='"&nm_obs&"', @cd_user="&cd_user&", @dt_atualiza='"&dt_hoje&"'"
					dbconn.execute(xsql)
				'end if
	Elseif acao = "apagar" and cd_tipo = "manut_obs" then
				'if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
					xsql = "DELETE FROM TBL_patrimonio_manut_obs WHERE cd_codigo = '"&cd_codigo&"'"
					dbconn.execute(xsql)
				'end if
	
''*************************************
'***	??? CANCELA EMPRESTIMO ???	***
'**************************************
	
	Elseif acao = "cancela" and cd_tipo = "emprestimo" then
		xsql = "up_patrimonio_emprestimo_cancela @cd_emprestimo='"&cd_cancela&"'"
		dbconn.execute(xsql)
		erro = erro &"/ emprestimo cancelado"
		'response.redirect("../patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo="&cd_pat_codigo&"&acao="&acao&"&caminho="&caminho&"")
		'response.redirect("../patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo="&cd_pat_codigo&"&acao="&acao&"&jan=1&caminho="&caminho&"")
	Else 
	action = "erro"
	End If
	
'***********************
'*** D E S C A R T E ***
'***********************
	If str_situacao = 1 Then
		'Descarta o patrimonio
		xsql = "up_patrimonio_descarte @cd_pat_codigo='"&cd_pat_codigo&"', @dt_descarte='"&dt_descarte&"', @nm_motivo='"&nm_obs&"'"
		dbconn.execute(xsql)
		response.write("Descarte OK")
		
		'********************
		'*** Movimentação ***
		'********************
		'*** Grava a empresa da Holding vinculada à unidade de negócio ***
		cd_empresa = 1'XXX -> Seleciona a empresa da holding vinculada à unidade de negócio XXX - *** NECESSITA ALTERAÇÃO ***
		id_movimento = 1 '*** indica o primeiro movimento do recém cadastrado patrimônio ***
		'*** busca o ultimo número de movimentação do patrimônio ***
		strsql = "SELECT MAX(id_movimento) AS ultimo FROM TBL_patrimonio_movimentacao WHERE cd_patrimonio='"&cd_pat_codigo&"'"
		set rs = dbconn.execute(strsql)
			ultimo_mov = int(rs("ultimo"))+1
				xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_unidade1&"', @cd_rack_origem='"&cd_rack&"', @cd_unidade_destino='"&cd_unidade_comodato&"', @cd_rack_destino='"&cd_rack_comodato&"', @cd_tipo_movimentacao="&str_situacao&", @dt_ida='"&dt_descarte&"', @dt_retorno=null, @id_movimento="&ultimo_mov&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
				response.write("<br><br>xsql = up_patrimonio_movimentacao_insert <br>@cd_patrimonio='"&cd_pat_codigo&"', <br>@cd_unidade_origem='"&cd_unidade1&"', <br>@cd_rack_origem='"&cd_rack&"', <br>@cd_unidade_destino='"&cd_unidade_comodato&"', <br>@cd_rack_destino='"&cd_rack_comodato&"', <br>@cd_tipo_movimentacao="&str_situacao&", <br>@dt_ida='"&dt_descarte&"', <br>@dt_retorno=null, <br>@id_movimento="&ultimo_mov&", <br>@cd_user="&cd_user&",<br>@dt_cadastro='"&dt_hoje&"'<br><br>")
				dbconn.execute(xsql)
				'**********************************************************************************************************************
		'response.redirect("../../patrimonio.asp?tipo=descarte&mensagem=Alterado com sucesso")
	
	
'**********************************
'***		MOVIMENTAÇÃO		***
'**********************************
	Elseif str_situacao = 5 AND cd_cancela > 0 Then
		'*** Não faz nada, apenas evita que dê erro ao cancelar uma movimentação

	'***************************
	'*** Emprestimo/comodato ***
	'***************************
	Elseif str_situacao = 4  AND cd_cancela = "" OR str_situacao = 5  AND cd_cancela = "" Then
	'*** 4 - empréstimo / 5 - Comodato ***
		
			if str_situacao = "5" then
				cd_unidade1 = cd_unidade2
			else
				dt_retorno = NULL
			end if
		
		'xsql = "up_patrimonio_empresta @cd_patrimonio='"&cd_pat_codigo&"', @cd_situacao='"&str_situacao&"', @nm_solicitante='"&nm_solicitante&"', @cd_unidade_ori='"&cd_unidade1&"', @cd_unidade_des='"&und_destino_obj&"', @dt_emprestimo="&dt_movimento_obj&", @dt_retorno='"&dt_retorno&"', @nm_empr_obs='"&nm_observacoes&"'"
		xsql = "up_patrimonio_empresta @cd_patrimonio='"&cd_pat_codigo&"', @cd_situacao='"&str_situacao&"', @nm_solicitante='"&nm_solicitante&"', @cd_unidade_ori='"&cd_unidade1&"', @cd_unidade_des='"&und_destino_obj&"', @dt_emprestimo="&dt_movimento_obj&", @dt_retorno='"&dt_retorno&"', @nm_empr_obs='"&nm_observacoes&"'"
		dbconn.execute(xsql)
		'response.write("<br>up_patrimonio_empresta @cd_patrimonio='"&cd_pat_codigo&"', <br>@cd_situacao='"&str_situacao&"', <br>@nm_solicitante='"&nm_solicitante&"', <br>@cd_unidade_ori='"&cd_unidade1&"', <br>@cd_unidade_des='"&cd_unidade_des&"', <br>@dt_emprestimo='"&dt_emprestimo&"', <br>@dt_retorno="&dt_retorno&", <br>@nm_empr_obs='"&nm_empr_obs&"'")
				
			'*** Altera o rack se for comodato *** '///// e movimentação de rack ***
			'if str_situacao = "5" OR str_situacao = "7" then '*** Altera a unidade de comodato ***
			'if str_situacao = "5" then '*** Altera a unidade de comodato ***
			'	xsql = "up_patrimonio_comodato @cd_pat_codigo="&cd_pat_codigo&", @cd_unidade_comodato='"&cd_unidade_des&"', @cd_rack_comodato='"&cd_rack_des&"'"
				'dbconn.execute(xsql)
			'end if
				
				'********************
				'*** MOVIMENTAÇÃO ***
				'********************
				'*** Grava a empresa da Holding vinculada à unidade de negócio ***
				cd_empresa = 1'XXX -> Seleciona a empresa da holding vinculada à unidade de negócio XXX - *** NECESSITA ALTERAÇÃO ***
				id_movimento = 1 '*** indica o primeiro movimento do recém cadastrado patrimônio ***
				'*** busca o ultimo número de movimentação do patrimônio ***
				strsql = "SELECT MAX(id_movimento) AS ultimo FROM TBL_patrimonio_movimentacao WHERE cd_patrimonio='"&cd_pat_codigo&"'"
				set rs = dbconn.execute(strsql)
					ultimo_mov = int(rs("ultimo"))+1
						'xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_unidade1&"', @cd_rack_origem='"&cd_rack&"', @cd_unidade_destino='"&cd_unidade_des&"', @cd_rack_destino='"&cd_rack_destino&"', @cd_tipo_movimentacao="&str_situacao&", @dt_ida="&dt_emprestimo&", @dt_retorno="&dt_retorno&", @id_movimento="&ultimo_mov&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
						'xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_unidade1&"', @cd_rack_origem='"&cd_rack&"', @cd_unidade_destino='"&und_destino_obj&"', @cd_rack_destino='"&rack_destino_obj&"', @cd_tipo_movimentacao="&str_situacao&", @dt_ida="&dt_movimento_obj&", @dt_retorno='"&dt_retorno&"', @id_movimento="&ultimo_mov&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
						xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_unidade1&"', @cd_rack_origem='"&cd_rack&"', @cd_unidade_destino='"&und_destino_obj&"', @cd_rack_destino='"&rack_destino_obj&"', @cd_tipo_movimentacao="&str_situacao&", @dt_ida="&dt_movimento_obj&", @dt_retorno=NULL, @id_movimento="&ultimo_mov&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"',@cd_natureza_mov='"&natureza_mov&"'"
						dbconn.execute(xsql)
						
						xsql = "UPDATE TBL_patrimonio SET cd_status=1 WHERE cd_codigo='"&cd_pat_codigo&"'"
						dbconn.execute(xsql)
				'**********************************************************************************************************************
				
		'response.redirect("../patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo="&cd_pat_codigo&"&acao="&acao&"&jan=1&caminho="&caminho&"")
		
	'*** Estoque ***
	elseif str_situacao = 6  AND cd_cancela = "" Then
		dt_retorno = "'"&dt_emprestimo&"'"
		'*** MOVIMENTAÇÃO ***
			'*** Grava a empresa da Holding vinculada à unidade de negócio ***
			cd_empresa = 1'XXX -> Seleciona a empresa da holding vinculada à unidade de negócio XXX - *** NECESSITA ALTERAÇÃO ***
			id_movimento = 1 '*** indica o primeiro movimento do recém cadastrado patrimônio ***
			'*** busca o ultimo número de movimentação do patrimônio ***
				strsql = "SELECT MAX(id_movimento) AS ultimo FROM TBL_patrimonio_movimentacao WHERE cd_patrimonio='"&cd_pat_codigo&"'"
				set rs = dbconn.execute(strsql)
					ultimo_mov = int(rs("ultimo"))+1
						xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_unidade1&"', @cd_rack_origem='"&cd_rack1&"', @cd_unidade_destino='"&cd_unidade_des&"', @cd_rack_destino='"&rack_destino_obj&"', @cd_tipo_movimentacao="&str_situacao&", @dt_ida='"&dt_estoque&"', @dt_retorno='"&dt_estoque&"', @id_movimento="&ultimo_mov&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
						dbconn.execute(xsql)
						'response.write("<br>---<br>xsql = up_patrimonio_movimentacao_insert <br>@cd_patrimonio='"&cd_pat_codigo&"', <br>@cd_unidade_origem='"&cd_unidade1&"', <br>@cd_rack_origem='"&cd_rack1&"', <br>@cd_unidade_destino='"&cd_unidade_des&"', <br>@cd_rack_destino='"&cd_rack_des&"', <br>@cd_tipo_movimentacao="&str_situacao&", <br>@dt_ida='"&dt_estoque&"', <br>@dt_retorno="&dt_estoque&", <br>@id_movimento="&ultimo_mov&", <br>@cd_user="&cd_user&",<br>@dt_cadastro='"&dt_hoje&"'<br>---<br>")
	
	'*** Rack ***
	elseif str_situacao = 7  AND cd_cancela = "" Then
		dt_retorno = "'"&dt_emprestimo&"'"
		'*** MOVIMENTAÇÃO ***
			'*** Grava a empresa da Holding vinculada à unidade de negócio ***
			cd_empresa = 1'XXX -> Seleciona a empresa da holding vinculada à unidade de negócio XXX - *** NECESSITA ALTERAÇÃO ***
			id_movimento = 1 '*** indica o primeiro movimento do recém cadastrado patrimônio ***
			'*** busca o ultimo número de movimentação do patrimônio ***
				strsql = "SELECT MAX(id_movimento) AS ultimo FROM TBL_patrimonio_movimentacao WHERE cd_patrimonio='"&cd_pat_codigo&"'"
				set rs = dbconn.execute(strsql)
					ultimo_mov = int(rs("ultimo"))+1
						'xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_unidade1&"', @cd_rack_origem='"&cd_rack&"', @cd_unidade_destino='"&und_destino_obj&"', @cd_rack_destino='"&rack_destino_obj&"', @cd_tipo_movimentacao="&str_situacao&", @dt_ida='"&dt_emprestimo&"', @dt_retorno="&dt_retorno&", @id_movimento="&ultimo_mov&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
						xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_unidade1&"', @cd_rack_origem='"&cd_rack&"', @cd_unidade_destino='"&und_destino_obj&"', @cd_rack_destino='"&rack_destino_obj&"', @cd_tipo_movimentacao="&str_situacao&", @dt_ida="&dt_movimento_obj&", @dt_retorno=NULL, @id_movimento="&ultimo_mov&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
						dbconn.execute(xsql)
	
	'*** Devolve ***
	elseif str_situacao = 8  AND cd_cancela = ""  Then
		xsql = "up_patrimonio_devolucao @cd_emprestimo='"&cd_emprestimo&"', @dt_pat_devolucao="&dt_movimento_obj&""
		'dbconn.execute(xsql)
			'*** altera o movimento para fechar o ciclo *** Altera a informação sobre a movimentação ***
			xsql = "up_patrimonio_movimentacao_fecha @cd_mov_atual='"&cd_mov_atual&"', @cd_patrimonio='"&cd_pat_codigo&"', @dt_pat_retorno="&dt_movimento_obj&", @cd_user="&cd_user&",@dt_devolucao='"&dt_hoje&"'"
			dbconn.execute(xsql)
	
	'*** Efetiva ***
	elseif str_situacao = 9  AND cd_cancela = "" Then
		'*** altera o movimento para fechar o ciclo ***
		xsql = "up_patrimonio_movimentacao_fecha @cd_mov_atual='"&cd_mov_atual&"', @cd_patrimonio='"&cd_pat_codigo&"', @dt_pat_retorno='"&dt_efetivacao&"', @cd_user="&cd_user&",@dt_devolucao='"&dt_hoje&"'"
		dbconn.execute(xsql)
		'response.redirect("../patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo="&cd_pat_codigo&"&acao="&acao&"&jan=1&caminho="&caminho&"")
	
	end if
	
	'response.redirect("../patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo="&cd_pat_codigo&"&acao="&acao&"&jan="&jan&"&caminho="&caminho&"")
	
END IF

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>

<!--body onLoad=window.close() onUnload="window.opener.location.reload()"-->
<!--body onLoad=window.close() onUnload="window.opener.form.submit()"-->

<%if int(cd_user) <> "46" then
	if jan = 1 then
		'if reload = "1" then%>
			<body onLoad=window.close() onUnload="window.opener.form.submit();">
		<%'else%>
			<!--body onLoad=window.close();-->
		<%'end if%>
	<%else%>
		<body>
	<%end if%>
<%end if%>
<br>*******************<br>
cd_situacao = <%=str_situacao%><br>
dt_emprestimo = <%=dt_emprestimo%><br>
dt_empr_devolucao = <%=dt_empr_devolucao%><br>
dt_efetivacao = <%=dt_efetivacao%><br>
cd_unidade_des = <%=cd_unidade_des%><br>
cd_rack_ori = <%=cd_rack_ori%><br>
cd_rack_des = <%=cd_rack_des%><br>
cd_emprestimo = <%=cd_emprestimo%><br><br>
nm_solicitante = <%=nm_solicitante%><br>
cd_mov_atual = <%=cd_mov_atual%><br>
<br>		

****************************<br>
Erro = <%=erro%><br>
cd_user = <%=cd_user%><br>
<br>
jan = <%=jan%><br>
existe = <%=existe%><br>
cd_cancela = <%=cd_cancela%><br>

Codigo = <%=cd_pat_codigo%><br>
<br>


Mensagem = <%=action%><br>
Ação: <%=acao%><br>
Tipo: <%=cd_tipo%><br><br>
Mes/Ano: <%=cd_mes%>/<%=cd_ano%><br>
Apagar: <%=cd_apaga%><br>
cd_obs: <%=cd_obs%>

<br>
Num fab.: <%=cd_pn%><br>
Serie: <%=cd_ns%><br>
Eqp.: <%=cd_equipamento%><br>
Unidade nota: <%=cd_unidade%><br>
Unidade Como/Empr: <%=cd_unidade_comodato%><br>
Marca: <%=cd_marca%><br>
<br>
Descarte: <%=dt_descarte%><br>
Motivo: <%=nm_motivo%><br>
Tipo: <%=nm_tipo%><br>
Especialidade: <%=cd_especialidade%><br>
cd_categoria>: <%=cd_categoria%><br>
cd_rack:<%=cd_rack%><br>

<br>
<br>
tipo:<%=nm_tipo%><br>
pat:<%=no_patrimonio%><br><br>

dt_plan_manut:<%=dt_plan_manut%><br><br>
//////////////////////////////////////<br>
Preventiva: <%=cd_preventiva%><br>
periodicidade: <%=dt_periodicidade_prev%><br>
dt prev:<%=dt_plan_prev%><br>
.....................................<br>
Calibração: <%=cd_calibracao%><br>
periodicidade: <%=dt_periodicidade_cal%><br>
dt cal: <%=dt_plan_cal%><br>
.....................................<br>
Seg. Eletrica: <%=cd_seg_elet%><br>
periodicidade: <%=dt_periodicidade_elet%><br>
dt elet: <%=dt_plan_elet%><br><br>
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\<br>
<br>

suspenso:<%=cd_suspenso%><br>

garantia: <%=cd_patrimonio_garantia%><br><br>

dia_emprestimo: <%=dia_emprestimo%><br><br>


dia_movimento_obj = <%=dia_movimento_obj%><br>
mes_movimento_obj = <%=mes_movimento_obj%><br>
ano_movimento_obj = <%=ano_movimento_obj%><br>
	dt_movimento_obj = <%=dt_movimento_obj%><br>
und_destino_obj = <%=und_destino_obj%><br>
nm_solicitante = <%=nm_solicitante%><br>
nm_observacoes = <%=nm_observacoes%><br>
nm_rack = <%=rack_destino_obj%><br><br>

xsql = "up_patrimonio_empresta @cd_patrimonio='"&<%=cd_pat_codigo%>&"', @cd_situacao='"&<%=str_situacao%>&"', @nm_solicitante='"&<%=nm_solicitante%>&"', @cd_unidade_ori='"&<%=cd_unidade1%>&"', @cd_unidade_des='"&<%=und_destino_obj%>&"', @dt_emprestimo="&<%=dt_movimento_obj%>&", @dt_retorno="&<%=dt_retorno%>&", @nm_empr_obs='"&<%=nm_observacoes%>&"'"
<br><br>
xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&<%=cd_pat_codigo%>&"', @cd_unidade_origem='"&<%=cd_unidade1%>&"', @cd_rack_origem='"&<%=cd_rack%>&"', @cd_unidade_destino='"&<%=und_destino_obj%>&"', @cd_rack_destino='"&<%=rack_destino_obj%>&"', @cd_tipo_movimentacao="&<%=str_situacao%>&", @dt_ida="&<%=dt_movimento_obj%>&", @dt_retorno="&<%=dt_retorno%>&", @id_movimento="&<%=ultimo_mov%>&", @cd_user="&<%=cd_user%>&",@dt_cadastro='"&<%=dt_hoje%>&"'"
<br><br>
cd_nova_situacao= <%=cd_nova_situacao%>											

</body>
</html>
