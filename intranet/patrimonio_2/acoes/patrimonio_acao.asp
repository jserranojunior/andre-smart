
  <!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
   <!--#include file="../../includes/inc_area_restrita.asp"-->

<%
cd_user = session("cd_codigo")
jan = request("jan")
	if jan = "" Then
		jan = 0
	end if
caminho = request("caminho")

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

cd_comprador = request("cd_comprador")
cd_unidade_compra = request("cd_unidade_compra")

cd_unidade = request("cd_unidade")
cd_unidade_comodato = request("cd_unidade_comodato")
cd_unit = request("cd_unit")
cd_rack = request("cd_rack")
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
cd_situacao = request("cd_situacao")
cd_unidade_des = request("cd_unidade_des")
cd_rack_des = request("cd_rack_des")

nm_solicitante = request("nm_solicitante")
nm_empr_obs = request("nm_empr_obs")

dia_emprestimo = request("dia_emprestimo")
mes_emprestimo = request("mes_emprestimo")
ano_emprestimo = request("ano_emprestimo")
	dt_emprestimo = mes_emprestimo&"/"&dia_emprestimo&"/"&ano_emprestimo

dia_empr_devolucao = request("dia_empr_devolucao")
mes_empr_devolucao = request("mes_empr_devolucao")
ano_empr_devolucao = request("ano_empr_devolucao")
	dt_empr_devolucao = mes_empr_devolucao&"/"&dia_empr_devolucao&"/"&ano_empr_devolucao
		if dia_empr_devolucao = "" Then
			dia_empr_devolucao = 0
		end if
		
cd_cancela = request("cd_cancela")

	'**********************************************
	'***	Comodato: alteração e efetivação.	***
	'**********************************************
	if cd_situacao = "1" AND cd_unidade <> "" AND cd_unidade_des <> "" AND dia_empr_devolucao = "" then ' inicia comodato
		cd_unidade_comodato = request("cd_unidade_des")
		cd_rack_comodato = request("cd_rack_des")
		'cd_rack = request("cd_rack_comodato")
		
		erro = "Descarta"
	elseif cd_situacao = "5" AND cd_unidade = cd_unidade_comodato AND cd_unidade_des <> "" AND dia_empr_devolucao = "" then ' inicia comodato
		cd_unidade_comodato = request("cd_unidade_des")
		cd_rack_comodato = request("cd_rack_des")
		'cd_rack = request("cd_rack_comodato")
		erro = "inclusao"
	elseif cd_situacao = "5" AND cd_unidade = cd_unidade_comodato AND cd_unidade_des = "" AND dia_empr_devolucao >= 0 then'efetiva e termina comodato (ver1)
		cd_unidade_comodato = request("cd_unidade")
		cd_rack_comodato = request("cd_rack")
		dt_empr_devolucao = month(now)&"/"&day(now)&"/"&year(now)
		dia_empr_devolucao = day(now)
		erro = "efetivação1"
	elseif cd_situacao = "5" AND cd_unidade <> cd_unidade_comodato AND cd_unidade_des = "" AND dia_empr_devolucao > 0 then'efetiva e termina comodato (ver2)
		'cd_unidade = request("cd_unidade_comodato")
		'cd_rack = request("cd_rack_comodato")
		cd_unidade_comodato = request("cd_unidade")
		cd_rack_comodato = request("cd_rack")
		dt_empr_devolucao = month(now)&"/"&day(now)&"/"&year(now)
		dia_empr_devolucao = day(now)
		erro = "efetivação2"
	Else
		Erro = "ok"
		'cd_rack_comodato = request("cd_rack_comodato")
		'cd_unidade_comodato = request("cd_unidade_comodato")
		cd_rack_comodato = request("cd_rack_comodato")
	end if
	
	
'*****************************
'*** INSERE AS OCORRENCIAS ***
'*****************************%>
<!--#include file="../../ocorrencia/acoes/ocorrencias_acao.asp"-->


<%IF x = "" then
' ****  AÇÕES  ****
	'Patrimonio
	existe = 0
	
'******************************
'***		INSERE			***
'******************************
	If acao = "inserir" AND cd_tipo = "patrimonio" Then
		'*** Verifica se o patrimonio já existe ***
		strsql ="SELECT * FROM View_patrimonio_lista WHERE no_patrimonio = "&no_patrimonio&" AND nm_tipo='"&nm_tipo&"'"
		Set rs_patrimonio = dbconn.execute(strsql)
		while not rs_patrimonio.EOF
			'aviso = "Este patrimonio já existe"
			existe = 1
		rs_patrimonio.movenext
		wend
	
	
		IF existe = 0  THEN
			'Insere dados
			xsql = "up_patrimonio_insert @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @no_patrimonio='"&no_patrimonio&"', @nm_tipo='"&nm_tipo&"', @cd_especialidade='"&cd_especialidade&"', @cd_unidade='"&cd_unidade&"', @cd_rack='"&cd_rack&"', @num_hospital='"&num_hospital&"', @cd_categoria='"&cd_categoria&"', @dt_data='"&dt_data&"', @cd_preventiva="&cd_preventiva&", @dt_periodicidade_prev="&dt_periodicidade_prev&", @cd_calibracao="&cd_calibracao&", @dt_periodicidade_cal="&dt_periodicidade_cal&",	@cd_seg_elet="&cd_seg_elet&", @dt_periodicidade_elet="&dt_periodicidade_elet&", @nao_constar="&nao_constar&""
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
							'elseif mes_prev = 12 then 'Virada de ano
							'	mes_prev = mes_prev
							'	ano_prev = ano_prev + 1
							else 'Mesmo ano
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
						dt_plan_cal_dia = cd_dia'day(dt_data)
						dt_plan_cal_mes = cd_mes'month(dt_data)
						dt_plan_cal_ano = cd_ano'year(dt_data)
							mes_cal = int(dt_periodicidade_cal) + int(dt_plan_cal_mes)
							ano_cal = dt_plan_cal_ano
							if int(mes_cal) > 12 then 'Virada de ano
								do while not mes_cal <= 12
									mes_cal = mes_cal - 12
									ano_cal = ano_cal + 1
								loop
							'elseif mes_cal = 12 then 'Virada de ano
							'	mes_cal = mes_cal
							'	ano_cal = ano_cal + 1
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
						dt_plan_elet_dia = cd_dia'day(dt_data)
						dt_plan_elet_mes = cd_mes'month(dt_data)
						dt_plan_elet_ano = cd_ano'year(dt_data)
							mes_elet = int(dt_periodicidade_elet) + int(dt_plan_elet_mes)
							ano_elet = dt_plan_elet_ano
							if int(mes_elet) > 12 then 'Virada de ano
								do while not mes_elet <= 12
									mes_elet = mes_elet - 12
									ano_elet = ano_elet + 1
								loop
							'elseif mes_elet = 12 then 'Virada de ano
							'	mes_elet = mes_elet
							'	ano_elet = ano_elet + 1
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
					cd_empresa = 1'XXX -> Seleciona a empresa da holding vinculada à unidade de negócio XXX - *** NECESSITA ALTERAÇÃO ***
					id_movimento = 1 '*** indica o primeiro movimento do recém cadastrado patrimônio ***
					xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_empresa&"', @cd_unidade_destino='"&cd_empresa&"', @cd_tipo_movimentacao=3, @dt_ida='"&dt_data&"', @dt_retorno='"&dt_data&"', @id_movimento="&id_movimento&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
					dbconn.execute(xsql)
					
			if jan <> 1 then
				response.redirect("../../patrimonio.asp?tipo=cadastro&mensagem=Inserido com sucesso")
			end if
		ELSE
			if jan <> 1 then
				response.redirect("../../patrimonio.asp?tipo=cadastro&caminho=patrimonio&jan=0&existe=1&no_patrimonio="&no_patrimonio&"&nm_tipo="&nm_tipo&"&cd_equipamento="&cd_equipamento&"&cd_marca="&cd_marca&"&cd_pn="&cd_pn&"&cd_ns="&cd_ns&"&cd_especialidade="&cd_especialidade&"&cd_unidade="&cd_unidade&"&cd_rack="&cd_rack&"&num_hospital="&num_hospital&"&cd_categoria="&cd_categoria&"&dt_data="&dt_data&"&cd_preventiva="&cd_preventiva&"&dt_periodicidade_prev="&dt_periodicidade_prev&"&cd_calibracao="&cd_calibracao&"&dt_periodicidade_cal="&dt_periodicidade_cal&"&cd_seg_elet="&cd_seg_elet&"&dt_periodicidade_elet="&dt_periodicidade_elet&"&nao_constar="&nao_constar&"")
			else
				response.redirect("../patrimonio_cad.asp?tipo=cadastro&caminho=patrimonio&jan=0&existe=1&no_patrimonio="&no_patrimonio&"&nm_tipo="&nm_tipo&"&cd_equipamento="&cd_equipamento&"&cd_marca="&cd_marca&"&cd_pn="&cd_pn&"&cd_ns="&cd_ns&"&cd_especialidade="&cd_especialidade&"&cd_unidade="&cd_unidade&"&cd_rack="&cd_rack&"&num_hospital="&num_hospital&"&cd_categoria="&cd_categoria&"&dt_data="&dt_data&"&cd_preventiva="&cd_preventiva&"&dt_periodicidade_prev="&dt_periodicidade_prev&"&cd_calibracao="&cd_calibracao&"&dt_periodicidade_cal="&dt_periodicidade_cal&"&cd_seg_elet="&cd_seg_elet&"&dt_periodicidade_elet="&dt_periodicidade_elet&"&nao_constar="&nao_constar&"")
			end if
		END IF
		
		
'******************************
'***		EDITA			***
'******************************		
	ElseIf acao = "editar" AND cd_tipo = "patrimonio" Then
		'Modifica os dados
		'xsql = "up_patrimonio_update @cd_pat_codigo="&cd_pat_codigo&", @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @no_patrimonio='"&no_patrimonio&"', @nm_tipo='"&nm_tipo&"', @cd_especialidade='"&cd_especialidade&"', @cd_unidade='"&cd_unidade&"', @cd_rack='"&cd_rack&"', @num_hospital='"&num_hospital&"', @cd_categoria='"&cd_categoria&"' ,@dt_data='"&dt_data&"', @cd_preventiva="&cd_preventiva&", @dt_periodicidade_prev="&dt_periodicidade_prev&", @cd_calibracao="&cd_calibracao&", @dt_periodicidade_cal="&dt_periodicidade_cal&", @cd_seg_elet="&cd_seg_elet&", @dt_periodicidade_elet="&dt_periodicidade_elet&", @nao_constar="&nao_constar&""
		xsql = "up_patrimonio_update @cd_pat_codigo="&cd_pat_codigo&", @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @no_patrimonio='"&no_patrimonio&"', @nm_tipo='"&nm_tipo&"', @cd_especialidade='"&cd_especialidade&"', @cd_unidade='"&cd_unidade&"', @cd_rack='"&cd_rack&"', @cd_unidade_comodato='"&cd_unidade_comodato&"', @cd_rack_comodato='"&cd_rack_comodato&"', @num_hospital='"&num_hospital&"', @cd_categoria='"&cd_categoria&"' ,@dt_data='"&dt_data&"', @cd_preventiva="&cd_preventiva&", @dt_periodicidade_prev="&dt_periodicidade_prev&", @cd_calibracao="&cd_calibracao&", @dt_periodicidade_cal="&dt_periodicidade_cal&", @cd_seg_elet="&cd_seg_elet&", @dt_periodicidade_elet="&dt_periodicidade_elet&", @nao_constar="&nao_constar&""
		dbconn.execute(xsql)
		response.write("patrimonio edita OK")
		
			'*** Planejamento Preventiva ***
			if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
				xsql = "up_patrimonio_preventiva_insert @tipo_manut=1, @cd_patrimonio='"&cd_pat_codigo&"', @cd_suspenso='"&cd_suspenso&"', @dt_plan_prev="&dt_plan_prev&",@dt_garantia="&dt_patrimonio_garantia&""
				dbconn.execute(xsql)
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
			strsql = "UPDATE TBL_patrimonio_movimentacao SET cd_unidade_origem="&cd_comprador&",cd_unidade_destino="&cd_unidade_compra&" WHERE (cd_patrimonio = "&cd_pat_codigo&") AND (id_movimento = 1)"
			Set rs = dbconn.execute(strsql)
			
		if jan <> 1 then
			response.redirect("../../patrimonio.asp?tipo=cadastro&mensagem=Alterado com sucesso")
		end if
		
		
		
'******************************
'***		DESCARTE		***
'******************************
	'ElseIf acao = "descartar" AND cd_tipo = "patrimonio" Then
		'Modifica os dados
		'xsql = "up_patrimonio_descarte @cd_pat_codigo='"&cd_pat_codigo&"', @dt_descarte='"&dt_data&"', @nm_motivo='"&nm_motivo&"'"
		'dbconn.execute(xsql)
		'response.write("Descarte OK")
		'response.redirect("../../patrimonio.asp?tipo=descarte&mensagem=Alterado com sucesso")
		
		
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
		
'******************************
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
	
	
	Elseif acao = "cancela" and cd_tipo = "emprestimo" then
		xsql = "up_patrimonio_emprestimo_cancela @cd_emprestimo='"&cd_cancela&"'"
		dbconn.execute(xsql)
		erro = "emprestimo cancelado"
		'response.redirect("../patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo="&cd_pat_codigo&"&acao="&acao&"&caminho="&caminho&"")
		response.redirect("../patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo="&cd_pat_codigo&"&acao="&acao&"&jan=1&caminho="&caminho&"")
	Else 
	action = "erro"
	End If
	
	
	
'**********************************
'***		Movimentação		***
'**********************************
	'*** Descarta ***
	if cd_situacao = 1 and cd_unidade_des <> "" and dt_emprestimo <> "" and cd_emprestimo = "" Then
		'Descarta o patrimonio
		xsql = "up_patrimonio_descarte @cd_pat_codigo='"&cd_pat_codigo&"', @dt_descarte='"&dt_data&"', @nm_motivo='"&nm_empr_obs&"'"
		dbconn.execute(xsql)
		response.write("Descarte OK")
		'response.redirect("../../patrimonio.asp?tipo=descarte&mensagem=Alterado com sucesso")
	
	'*** Empresta *** 
	Elseif cd_situacao <> "" and cd_unidade_des <> "" and dt_emprestimo <> "" and cd_emprestimo = "" Then
		xsql = "up_patrimonio_empresta @cd_patrimonio='"&cd_pat_codigo&"', @cd_situacao='"&cd_situacao&"', @nm_solicitante='"&nm_solicitante&"', @cd_unidade_ori='"&cd_unidade&"', @cd_unidade_des='"&cd_unidade_des&"', @dt_emprestimo='"&dt_emprestimo&"',@nm_empr_obs='"&nm_empr_obs&"'"
		dbconn.execute(xsql)
		
			if cd_situacao = "5" then '*** Altera a unidade de comodato ***
				xsql = "up_patrimonio_comodato @cd_pat_codigo="&cd_pat_codigo&", @cd_unidade_comodato='"&cd_unidade_des&"', @cd_rack_comodato='"&cd_rack_des&"'"
				dbconn.execute(xsql)
					
			end if
			
				'*** MOVIMENTAÇÃO ***
				'*** Grava a empresa da Holding vinculada à unidade de negócio ***
				cd_empresa = 1'XXX -> Seleciona a empresa da holding vinculada à unidade de negócio XXX - *** NECESSITA ALTERAÇÃO ***
				id_movimento = 1 '*** indica o primeiro movimento do recém cadastrado patrimônio ***
				'*** busca o ultimo número de movimentação do patrimônio ***
				strsql = "SELECT MAX(id_movimento) AS ultimo FROM TBL_patrimonio_movimentacao WHERE cd_patrimonio='"&cd_pat_codigo&"'"
				set rs = dbconn.execute(strsql)
					ultimo_mov = int(rs("ultimo"))+1
						xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_pat_codigo&"', @cd_unidade_origem='"&cd_unidade&"', @cd_unidade_destino='"&cd_unidade_des&"', @cd_tipo_movimentacao="&cd_situacao&", @dt_ida='"&dt_emprestimo&"', @dt_retorno=NULL, @id_movimento="&ultimo_mov&", @cd_user="&cd_user&",@dt_cadastro='"&dt_hoje&"'"
						dbconn.execute(xsql)
				
		response.redirect("../patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo="&cd_pat_codigo&"&acao="&acao&"&jan=1&caminho="&caminho&"")
		
	'*** Devolve ***
	elseif cd_situacao <> "" and dia_empr_devolucao > 0 and cd_emprestimo <> "" Then
		response.write("<br>\\\***devolve***///<br>")
		xsql = "up_patrimonio_devolucao @cd_emprestimo='"&cd_emprestimo&"', @dt_pat_devolucao='"&dt_empr_devolucao&"'"
		dbconn.execute(xsql)
		
			'*** altera o movimento para fechar o ciclo ***
			xsql = "up_patrimonio_movimentacao_fecha @cd_patrimonio='"&cd_pat_codigo&"', @dt_pat_retorno='"&dt_empr_devolucao&"', @cd_user="&cd_user&",@dt_devolucao='"&dt_hoje&"'"
			dbconn.execute(xsql)
		
		response.redirect("../patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo="&cd_pat_codigo&"&acao="&acao&"&jan=1&caminho="&caminho&"")
	end if
	
END IF

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>

<!--body onLoad=window.close() onUnload="window.opener.location.reload()"-->
<!--body onLoad=window.close() onUnload="window.opener.form.submit()"-->

<%if int(cd_user) <> "460" then
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
cd_situacao = <%=cd_situacao%><br>
dt_empr_devolucao = <%=dt_empr_devolucao%><br>
cd_unidade_des = <%=cd_unidade_des%><br>
cd_emprestimo = <%=cd_emprestimo%><br>
****************************<br>
Erro = <%=erro%><br>
cd_user = <%=cd_user%><br>
<br>
jan = <%=jan%><br>
existe = <%=existe%><br>


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
Unidade: <%=cd_unidade%><br>
Unidade: <%=cd_unidade_comodato%><br>
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
\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\<br>
<br>

suspenso:<%=cd_suspenso%><br>

garantia: <%=cd_patrimonio_garantia%>




</body>
</html>
