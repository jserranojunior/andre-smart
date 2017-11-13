
  <!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
   <!--#include file="../../includes/inc_area_restrita.asp"-->

<%data_atual = Month(now)&"/"&Day(now)&"/"&Year(now)&" "&Hour(now)&":"&Minute(now)&":"&Second(now)
cd_user = session("cd_codigo")

acao = request("acao")
filtro = request("filtro")
jan = request("jan")

sequencia = request("sequencia")
cd_codigo = request("cd_codigo")
campo = request("campo")

'Dados da ocorrencia
	cd_cod_ocorrencia = request("cd_cod_ocorrencia")
	cd_origem = request("cd_origem")
	id_origem = request("id_origem")
	ocorrencia = request("ocorrencia")
	dia_ocorrencia = request("dia_ocorrencia")
	mes_ocorrencia = request("mes_ocorrencia")
	ano_ocorrencia = request("ano_ocorrencia")
		dt_ocorrencia = mes_ocorrencia&"/"&dia_ocorrencia&"/"&ano_ocorrencia
	cd_retorno = request("cd_retorno")



'Dados da O.S.
	cd_codigo_os = request("cd_codigo_os")
	
	cd_unidade = request("cd_unidade")
	num_os = request("num_os")
	cd_natureza_os = request("cd_natureza_os")
	dt_os = request("dt_os")
		dt_os_dia = request("dt_os_dia")
		dt_os_mes = request("dt_os_mes")
		dt_os_ano = request("dt_os_ano")
			'dt_os = Month(dt_os)&"/"&Day(dt_os)&"/"&Year(dt_os)
			dt_os = dt_os_mes&"/"&dt_os_dia&"/"&dt_os_ano
			
		dt_abertura = dt_os
	nm_solicitante  = request("nm_solicitante")
	cd_unidade_os = request("cd_unidade_os")
	nm_uni_select = request("nm_uni_select")
	
	num_qtd = request("num_qtd")
	cd_especialidade = request("cd_especialidade")
	cd_equipamento = request("cd_equipamento")
	cd_marca = request("cd_marca")
	cd_ns = request("cd_ns")
	cd_patrimonio = request("cd_patrimonio")
		'*** Caso o patrimonio seja nulo ***
		if cd_patrimonio = "" Then
			cd_patrimonio = "NULL"
		end if
		
		cd_patrimonio_atual = request("cd_patrimonio_atual")				
		cd_equipamento_atual = request("cd_equipamento_atual")
		cd_marca_atual = request("cd_marca_atual")
		cd_especialidade_atual = request("cd_especialidade_atual")
		cd_ns_atual = request("cd_ns_atual")
		num_qtd_atual = request("num_qtd_atual")
	
	motivo = request("motivo")
	digito = request("digito")
	nm_natureza_defeito = request("nm_natureza_defeito")
	
	dt_entrada = request("dt_entrada")
		if dt_entrada <> "" Then
			dt_entrada = "'"&month(dt_entrada)&"/"&day(dt_entrada)&"/"&year(dt_entrada)&"'"
		else
			'dt_entrada = "NULL"
			dt_entrada = dt_os
		end if
		
	cd_avaliacao = request("cd_avaliacao")
	cd_fornecedor = request("cd_fornecedor")
	cd_liberacao = request("cd_liberacao")
	observacoes = request("observacoes")

'Fechamento da O.S.
	fecha = request("fecha")
	fecha = 1

'Dados da Manutenção
	cd_cod = request("cd_cod")
	cd_os = request.form("cd_os")
	dt_encaminhado = request.form("dt_encaminhado")
	nm_orcamento = request.form("nm_orcamento")
	dt_prev_resposta = request("dt_prev_resposta")
	num_osforn = request("num_osforn")
	dt_previsao = request("dt_previsao")
	nm_processo = request("nm_processo")
	nm_processo_select = request("nm_processo_select")
	comentarios = request("comentarios")
	dt_retorno = request("dt_retorno")
	cd_alerta = request("cd_alerta")



'Dados da ????

	cd_avaliacao = request("cd_avaliacao")
	sequencia = request("sequencia")
	dt_aval_data = request("dt_aval_data")
	nm_liberacao = request("nm_liberacao")
	
	'dt_orcamento_dia = request("dt_orcamento_dia")
	'dt_orcamento_mes = request("dt_orcamento_mes")
	'dt_orcamento_ano = request("dt_orcamento_ano")
	'dt_orcamento = dt_orcamento_dia&"/"&dt_orcamento_mes&"/"&dt_orcamento_ano
	dt_orcamento = request("dt_orcamento")
	'	If dt_orcamento = "" Then
	'		dt_orcamento = "Null"
	'	else
	'		dt_orcamento = "'"&month(dt_orcamento)&"/"&day(dt_orcamento)&"/"&year(dt_orcamento)&"'"
	'	End if
	
	outro = request("outro")
	cd_status = request("cd_status")
	nm_contato = request("nm_contato")
	dt_tempo_resp = request("dt_tempo_resp")
	dt_resposta = request("dt_resposta")
	dt_conclusao = request("dt_conclusao")
	
'Manutencao & Orcamento
	dt_manut_orc = request("dt_manut_orc")
	dt_resposta_orc = request("dt_resposta_orc")
	num_os_assist = request("num_os_assist")
	dt_prev_resposta = request("dt_prev_resposta")
	dt_prev_retorno = request("dt_prev_retorno")
	dt_retorno_un = request("dt_retorno_un")
	cd_situacao = request("cd_situacao")
	cd_gestao = request("cd_gestao")
	
	
	cd_qtd_valor = request("cd_qtd_valor")
	cd_valor_unit = request("cd_valor_unit")
		cd_valor_unit = replace(cd_valor_unit,".",",")
	cd_valor_orc = request("cd_valor_orc")
		cd_valor_orc = replace(cd_valor_orc,".",",")
	
	If dt_manut_orc = "" Then
		dt_manut_orc = "NULL"
	else
		dt_manut_orc = "'"&month(dt_manut_orc)&"/"&day(dt_manut_orc)&"/"&year(dt_manut_orc)&"'"
	End if
	
	If dt_resposta_orc = "" Then
		dt_resposta_orc = "Null"
	else
		dt_resposta_orc = "'"&month(dt_resposta_orc)&"/"&day(dt_resposta_orc)&"/"&year(dt_resposta_orc)&"'"
	End if
	
	If dt_prev_resposta = "" Then
		dt_prev_resposta = "Null"
	else
		dt_prev_resposta = "'"&month(dt_prev_resposta)&"/"&day(dt_prev_resposta)&"/"&year(dt_prev_resposta)&"'"
	End if
	
	If dt_prev_retorno = "" Then
		dt_prev_retorno = "Null"
	else
		dt_prev_retorno = "'"&month(dt_prev_retorno)&"/"&day(dt_prev_retorno)&"/"&year(dt_prev_retorno)&"'"
	End if
	
	If dt_retorno_un = "" Then
		dt_retorno_un = "Null"
	else
		dt_retorno_un = "'"&month(dt_retorno_un)&"/"&day(dt_retorno_un)&"/"&year(dt_retorno_un)&"'"
	End if

' Cadastro de orçamentos 
	cd_orcamento = request("cd_orcamento")
	nr_orcamento = request("nr_orcamento")
	nm_valor = request("nm_valor")
	nr_parcela = request("nr_parcela")
		if nr_parcela = "" Then nr_parcela = 1
	dt_dia = request("dt_dia")
	dt_mes = request("dt_mes")
	dt_ano = request("dt_ano")
		If dt_dia = "" AND dt_mes = "" AND  dt_ano = "" Then
			dt_orcamento_aprov = "Null"
		else
			dt_orcamento_aprov = "'"&dt_mes&"/"&dt_dia&"/"&dt_ano&"'"
		End if

'Condições

'fecha = request("fechado")
If nm_unidade = "outros" Then
nm_unidade = nm_uni_select
End If

If nm_processo = "2" Then
nm_processo = nm_processo_select
End if

If dt_previsao = "" Then
	dt_previsao = "Null"
else
	dt_previsao = "'"&month(dt_previsao)&"/"&day(dt_previsao)&"/"&year(dt_previsao)&"'"
End if

If dt_retorno = "" Then
	dt_retorno = "Null"
else
	dt_retorno = "'"&month(dt_retorno)&"/"&day(dt_retorno)&"/"&year(dt_retorno)&"'"
End if

If dt_tempo_resp = "" Then
	dt_tempo_resp = "Null"
else
	dt_tempo_resp = "'"&month(dt_tempo_resp)&"/"&day(dt_tempo_resp)&"/"&year(dt_tempo_resp)&"'"
End if

If dt_resposta = "" Then
	dt_resposta = "Null"
else
	dt_resposta = "'"&month(dt_resposta)&"/"&day(dt_resposta)&"/"&year(dt_resposta)&"'"
End if

If dt_conclusao = "" Then
	dt_conclusao = "Null"
else
	dt_conclusao = "'"&month(dt_conclusao)&"/"&day(dt_conclusao)&"/"&year(dt_conclusao)&"'"
End if

If comentarios = "" Then
   comentarios = ""
   End if

if cdt <> "" Then
	acao = "3"
	End if

if outro = "" Then
outro = "0"
end if

If nm_orcamento = "0" Then
nm_processo = "0"
End if

if cd_avaliacao = "aguardando" Then
cd_fornecedor = "23"
End if

'Fim das condições

'******************************** Registro de ocorrencias *******************************************
' Registra ocorrências 
if ocorrencia <> "" then
xsql = "up_ocorrencia_insert @cd_origem='"&cd_origem&"', @id_origem='"&id_origem&"', @dt_ocorrencia='"&dt_ocorrencia&"', @ocorrencia='"&ocorrencia&"'"
dbconn.execute(xsql)
	'desativada o redirecionamento da página
	'if cd_codigo <> "" Then
	'response.redirect("../manutencao.asp?tipo=andamento&cd_codigo="&cd_retorno&"&campo=cd_codigo")
	'end if
'continua a operação com a O.S.
end if

' Apaga ocorrências
if cd_cod_ocorrencia <> "" Then
xsql = "up_ocorrencia_delete @cd_cod_ocorrencia='"&cd_cod_ocorrencia&"'"
dbconn.execute(xsql)
'Retorna par a origem
response.redirect("../manutencao.asp?tipo=andamento&cd_codigo="&cd_codigo&"&campo=cd_codigo")
end if

x = ""

if x = "" then

'*******************************************************************************
'    					Cadastro - Ordem de Serviço								
'*******************************************************************************
if acao = "1" And filtro = "cadastra_os" then
	'verifica se a os a ser cadastrada já existe no banco de dados
	strsql = "SELECT * FROM TBL_OrdemServico2 WHERE (cd_unidade = '"&cd_unidade&"') AND (num_os = '"&num_os&"')"
	Set rs = dbconn.execute(strsql)
	while not rs.EOF 
		os_gravada = 1
	rs.movenext
	wend
	
		if os_gravada = "" then 'Se a OS não existir, passa pelo processo de inclusão
		
			if dt_entrada = "Null" Then
				'dt_entrada = Month(dt_os)&"-"&Day(dt_os)&"-"&Year(dt_os)
				dt_entrada = Month(dt_os)&"/"&Day(dt_os)&"/"&Year(dt_os)
			end if
						
			'Cadastra Ordem de servico
			'xsql = "up_os_insert2 @cd_unidade='"&cd_unidade&"', @num_os='"&num_os&"', @dt_os="&dt_os&"', @nm_solicitante='"&nm_solicitante&"', @nm_unidade='"&nm_unidade&"', @num_qtd='"&num_qtd&"', @cd_equipamento='"&cd_equipamento&"', @cd_ns='"&cd_ns&"', @no_patrimonio='"&no_patrimonio&"', @cd_especialidade='"&cd_especialidade&"', @cd_marca='"&cd_marca&"', @motivo='"&motivo&"', @dt_entrada="&dt_entrada&", @cd_avaliacao='"&cd_avaliacao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_liberacao='"&cd_liberacao&"', @observacoes='"&observacoes&"',@sequencia='"&sequencia&"',@data_inc='"&data_atual&"',@cd_user="&cd_user&",@fecha=0"
			xsql = "up_os_insert2 @cd_unidade='"&cd_unidade&"', @num_os='"&num_os&"', @cd_natureza_os='"&cd_natureza_os&"', @dt_os='"&dt_os&"', @nm_solicitante='"&nm_solicitante&"', @cd_unidade_os="&cd_unidade_os&", @num_qtd='"&num_qtd&"', @cd_equipamento='"&cd_equipamento&"', @cd_especialidade='"&cd_especialidade&"', @cd_marca='"&cd_marca&"', @cd_ns='"&cd_ns&"', @cd_patrimonio="&cd_patrimonio&", @motivo='"&motivo&"', @dt_entrada='"&dt_entrada&"', @cd_avaliacao='"&cd_avaliacao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_liberacao='"&cd_liberacao&"', @observacoes='"&observacoes&"', @sequencia='"&sequencia&"', @cd_user="&cd_user&", @data_inc='"&data_atual&"', @fecha=0"
			dbconn.execute(xsql)
			
				'if isnull(cd_patrimonio) Then
				'	strsql = "SELECT MAX(id_movimento) AS ultimo FROM TBL_patrimonio_movimentacao WHERE cd_patrimonio='"&cd_patrimonio&"'"
				'	set rs_mov = dbconn.execute(strsql)
				'		if not rs_mov.EOF then
				'			ultimo_mov = int(rs_mov("ultimo"))+1
				'			xsql = "up_patrimonio_movimentacao_insert @cd_patrimonio='"&cd_patrimonio&"', @cd_unidade_origem='"&cd_unidade_os&"', @cd_unidade_destino=27, @cd_tipo_movimentacao=2, @dt_ida='"&dt_os&"', @dt_retorno=NULL, @id_movimento="&ultimo_mov&", @cd_user="&cd_user&",@dt_cadastro='"&data_atual&"'"
				'			dbconn.execute(xsql)
				'		end if
				'end if
				
				'Verifica o codigo da O.S recem cadastrada e retorna o codigo para cadastro da data de previsão de entraga no andamento
				strsql = "SELECT * FROM TBL_OrdemServico2 WHERE (cd_unidade = "&cd_unidade&") AND (num_os = "&num_os&")"
				Set rs = dbconn.execute(strsql)
				while not rs.EOF 
					ver_os_gravada = 1
					cd_ordemservico = rs("cd_codigo")
					response.write("<br>***** ver_os_gravada"&ver_os_gravada&"********<br>")
				rs.movenext
				wend
				
					if ver_os_gravada = "1" then
						'verifica a data de entrada e o prazo do fornecedor
						'	xsql = "SELECT * FROM TBL_fornecedor where cd_codigo="&cd_fornecedor&""
						'	SET rs = dbconn.execute(xsql)
						'		if not rs.EOF then
						'			cd_prazo = rs("cd_prazo")
						'		end if
							
							cd_prazo = 15
									
									if cd_prazo > 0 then
										dia_estimado = int(dt_os_dia) + int(cd_prazo)
											mes_estimado = dt_os_mes
											ano_estimado = dt_os_ano
											ultimo_dia = ultimodiames(mes_estimado,ano_estimado)
												
												if 	int(dia_estimado) > int(ultimo_dia) then
													dia_estimado = int(dia_estimado) - int(ultimo_dia)
													mes_estimado = int(mes_estimado) + 1
														if mes_estimado > 12 then
															mes_estimado = mes_estimado - 12
															ano_estimado = int(ano_estimado) + 1
														else
															ano_estimado = ano_estimado
														end if
														dt_prev_retorno = "'"&dia_estimado&"/"&mes_estimado&"/"&ano_estimado&"'"
														'dt_prev_retorno = "'"&mes_estimado&"/"&dia_estimado&"/"&ano_estimado&"'"
												else
													'dt_prev_retorno = "'"&dia_estimado&"/"&mes_estimado&"/"&ano_estimado&"'"
													dt_prev_retorno = "'"&mes_estimado&"/"&dia_estimado&"/"&ano_estimado&"'"
												end if
								
									else
										'dt_prev_retorno = "'"&Day(dt_abertura)&"/"&Month(dt_abertura)&"/"&Year(dt_abertura)&"'"
										dt_prev_retorno = "'"&dt_os_mes&"/"&dt_os_dia&"/"&dt_os_ano&"'"
										
									end if
									
									response.write("dt_prev_retorno:"&dt_manut_orc&"<br>dt_prev_resposta:"&dt_prev_resposta&"<br>dt_resposta_orc:"&dt_resposta_orc&"<br>dt_prev_retorno:"&dt_prev_retorno&"<br>dt_retorno:"&dt_retorno&"<br>dt_retorno_un:"&dt_retorno_un&"<br>")
				
						'xsql = "up_andamento_insert2 @cd_os='"&num_os&"', @cd_cod='"&cd_ordemservico&"', @dt_manut_orc='"&Month(dt_manut_orc)&"-"&day(dt_manut_orc)&"-"&year(dt_manut_orc)&"',@cd_fornecedor='"&cd_fornecedor&"', @nm_contato='"&nm_contato&"', @dt_prev_resposta='"&Month(dt_prev_resposta)&"-"&day(dt_prev_resposta)&"-"&year(dt_prev_resposta)&"', @dt_resposta_orc="&Month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"-"&year(dt_resposta_orc)&", @nm_orcamento='"&nm_orcamento&"', @num_os_assist='"&num_os_assist&"', @dt_prev_retorno='"&Month(dt_prev_retorno)&"-"&day(dt_prev_retorno)&"-"&year(dt_prev_retorno)&"', @cd_situacao='"&cd_situacao&"', @cd_gestao='"&cd_gestao&"', @comentarios='"&comentarios&"', @dt_retorno='"&Month(dt_retorno)&"-"&day(dt_retorno)&"-"&year(dt_retorno)&"', @dt_retorno_un='"&Month(dt_retorno_un)&"-"&day(dt_retorno_un)&"-"&year(dt_retorno_un)&"', @cd_qtd_valor='"&cd_qtd_valor&"', @cd_valor_unit='"&cd_valor_unit&"', @cd_valor_orc='"&cd_valor_orc&"', @nm_natureza_defeito='"&nm_natureza_defeito&"',@sequencia='"&sequencia&"', @cd_alerta='"&cd_alerta&"'"
						xsql = "up_andamento_insert2 @cd_os='"&num_os&"', @cd_cod='"&cd_ordemservico&"', @dt_manut_orc="&dt_manut_orc&",@cd_fornecedor='"&cd_fornecedor&"', @nm_contato='"&nm_contato&"', @dt_prev_resposta="&dt_prev_resposta&", @dt_resposta_orc="&dt_resposta_orc&", @nm_orcamento='"&nm_orcamento&"', @num_os_assist='"&num_os_assist&"', @dt_prev_retorno="&dt_prev_retorno&", @cd_situacao='"&cd_situacao&"', @cd_gestao='1', @comentarios='"&comentarios&"', @dt_retorno="&dt_retorno&", @dt_retorno_un="&dt_retorno_un&", @cd_qtd_valor='"&cd_qtd_valor&"', @cd_valor_unit='"&cd_valor_unit&"', @cd_valor_orc='"&cd_valor_orc&"', @nm_natureza_defeito='"&nm_natureza_defeito&"',@sequencia='"&sequencia&"', @cd_alerta='"&cd_alerta&"'"
						dbconn.execute(xsql)
					end if
		else 'Se a OS Existir, abre a OS e permite a inclusão do andamento da OS
			response.redirect("../manutencao2.asp?tipo=andamento&cd_codigo="&cd_ordemservico&"&campo=cd_codigo&ordem=1")
		end if
	
	response.redirect("../../manutencao2.asp?tipo=andamento&cd_codigo="&cd_ordemservico&"&campo=cd_codigo&ordem=1")



'*******************************************************************************
'    					Modifica - Ordem de Serviço								
'*******************************************************************************
elseif acao = "2" And filtro = "modifica_os" then
	xsql = "up_os_update2 @cd_codigo='"&cd_codigo&"', @num_os='"&num_os&"', @cd_natureza_os='"&cd_natureza_os&"', @dt_os='"&dt_os&"', @nm_solicitante='"&nm_solicitante&"', @cd_unidade_os="&cd_unidade_os&", @num_qtd='"&num_qtd&"', @cd_equipamento='"&cd_equipamento&"', @cd_especialidade='"&cd_especialidade&"', @cd_marca='"&cd_marca&"', @cd_ns = '"&cd_ns&"', @cd_patrimonio = "&cd_patrimonio&", @motivo='"&motivo&"', @dt_entrada='"&dt_entrada&"', @cd_avaliacao='"&cd_avaliacao&"', @cd_fornecedor='"&cd_fornecedor&"', @observacoes='"&observacoes&"',@data_inc='"&data_atual&"',@cd_user="&cd_user&",@sequencia='"&sequencia&"'"
	dbconn.execute(xsql)
	
	if jan <> 1 then
		response.redirect("../../manutencao2.asp?tipo=andamento&cd_codigo="&cd_codigo&"&campo=cd_codigo")
		'link = "ordemservico_andamento.asp?cd_codigo="&num_os&"&campo=num_os"
	End if



'*******************************************************************************
'    						Insere - Andamento									
'*******************************************************************************
Elseif acao = "1" AND filtro = "andamento" Then

	'****************************
	'*** Inclui a providência ***
	xsql = "up_os_update_andamento2 @cd_codigo='"&cd_codigo_os&"', @dt_entrada="&dt_entrada&", @cd_avaliacao='"&cd_avaliacao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_liberacao='"&cd_liberacao&"', @observacoes='"&observacoes&"',@data_inc='"&data_atual&"',@cd_user="&cd_user&",@sequencia='"&sequencia&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	
	'***************************************************
	'*** Verifica se já existe um andamento desta OS ***
	strsql = "SELECT * FROM TBL_andamento2 WHERE cd_cod = '"&cd_codigo_os&"'"
	Set rs = dbconn.execute(strsql)
	while not rs.EOF 
		andamento_gravado = 1
		cd_codos = rs("cd_codigo")
	rs.movenext
	wend

		if andamento_gravado = "" then
			xsql = "up_andamento_insert2 @cd_os='"&cd_os&"', @cd_cod='"&cd_cod&"', @dt_manut_orc="&dt_manut_orc&",@cd_fornecedor='"&cd_fornecedor&"', @nm_contato='"&nm_contato&"', @dt_prev_resposta="&dt_prev_resposta&", @dt_resposta_orc="&dt_resposta_orc&", @nm_orcamento='"&nm_orcamento&"', @num_os_assist='"&num_os_assist&"', @dt_prev_retorno="&dt_prev_retorno&", @cd_situacao='"&cd_situacao&"', @cd_gestao='"&cd_gestao&"', @comentarios='"&comentarios&"', @dt_retorno="&dt_retorno&", @dt_retorno_un="&dt_retorno_un&", @cd_qtd_valor='"&cd_qtd_valor&"', @cd_valor_unit='"&cd_valor_unit&"', @cd_valor_orc='"&cd_valor_orc&"', @nm_natureza_defeito='"&nm_natureza_defeito&"',@sequencia='"&sequencia&"', @cd_alerta='"&cd_alerta&"'"
			action = "editado com sucesso"
			dbconn.execute(xsql)
			
						'*************************************************************
						'***************** Fechamento da O.S. ************************
						If dt_retorno_un <> "Null" Then
							'Modifica os dados - Fecha a O.S.
							xsql = "up_os_fechamento2 @cd_os='"&cd_codigo_os&"', @fecha=1"
							dbconn.execute(xsql)
						End if
			
			response.redirect("../../manutencao2.asp?tipo=pendencias&strtop=20&filtro=pendentes")
			'link = "manutencao.asp?cd_codigo="&num_os&"&campo=num_os"
		'elseif andamento_gravado = 1 then
			'response.redirect("manutencao_acao.asp?filtro=andamento&acao=2&cd_codigo='"&cd_codigo&"'&cd_os='"&cd_os&"'&cd_cod='"&cd_cod&"'&dt_manut_orc='"&Month(dt_manut_orc)&"-"&day(dt_manut_orc)&"-"&year(dt_manut_orc)&"'&cd_fornecedor='"&cd_fornecedor&"'&nm_contato='"&nm_contato&"'&dt_prev_resposta='"&Month(dt_prev_resposta)&"-"&day(dt_prev_resposta)&"-"&year(dt_prev_resposta)&"'&dt_resposta_orc='"&Month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"-"&year(dt_resposta_orc)&"'&nm_orcamento='"&nm_orcamento&"'&num_os_assist='"&num_os_assist&"'&dt_prev_retorno='"&Month(dt_prev_retorno)&"-"&day(dt_prev_retorno)&"-"&year(dt_prev_retorno)&"'&cd_situacao='"&cd_situacao&"'&comentarios='"&comentarios&"'&dt_retorno='"&Month(dt_retorno)&"-"&day(dt_retorno)&"-"&year(dt_retorno)&"'&dt_retorno_un='"&Month(dt_retorno_un)&"-"&day(dt_retorno_un)&"-"&year(dt_retorno_un)&"'&cd_qtd_valor='"&cd_qtd_valor&"'&cd_valor_unit='"&cd_valor_unit&"'&cd_valor_orc='"&cd_valor_orc&"'&nm_natureza_defeito='"&nm_natureza_defeito&"'&sequencia='"&sequencia&"'&cd_alerta='"&cd_alerta&"'")
			'response.redirect("manutencao_acao.asp?filtro=andamento&acao=2&cd_codigo="&cd_codos&"&cd_os="&cd_os&"&cd_cod="&cd_cod&"&dt_manut_orc="&day(dt_manut_orc)&"/"&month(dt_manut_orc)&"/"&year(dt_manut_orc)&"&cd_fornecedor="&cd_fornecedor&"&nm_contato="&nm_contato&"&dt_prev_resposta="&day(dt_prev_resposta)&"/"&month(dt_prev_resposta)&"/"&year(dt_prev_resposta)&"&dt_resposta_orc="&day(dt_resposta_orc)&"/"&month(dt_resposta_orc)&"/"&year(dt_resposta_orc)&"&nm_orcamento="&nm_orcamento&"&num_os_assist="&num_os_assist&"&dt_prev_retorno="&day(dt_prev_retorno)&"/"&month(dt_prev_retorno)&"/"&year(dt_prev_retorno)&"&cd_situacao="&cd_situacao&"&cd_gestao="&cd_gestao&"&comentarios="&comentarios&"&dt_retorno="&day(dt_retorno)&"/"&month(dt_retorno)&"/"&year(dt_retorno)&"&dt_retorno_un="&day(dt_retorno_un)&"/"&month(dt_retorno_un)&"/"&year(dt_retorno_un)&"&cd_qtd_valor="&cd_qtd_valor&"&cd_valor_unit="&cd_valor_unit&"&cd_valor_orc="&cd_valor_orc&"&nm_natureza_defeito="&nm_natureza_defeito&"&sequencia="&sequencia&"&cd_alerta="&cd_alerta&"")
		end if


'*******************************************************************************
'    						Altera - Andamento									
'*******************************************************************************
Elseif acao = "2" AND filtro = "andamento" Then
	
	'***************************************************************************
	'*** Ação para preservar o banco de dados e os atualização dos registros ***
		xsql = "up_os_update_andamento2 @cd_codigo='"&cd_codigo_os&"', @dt_entrada="&dt_entrada&", @cd_avaliacao='"&cd_avaliacao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_liberacao='"&cd_liberacao&"', @observacoes='"&observacoes&"',@data_inc='"&data_atual&"',@cd_user="&cd_user&",@sequencia='"&sequencia&"'"
		action = "editado com sucesso"
		dbconn.execute(xsql)
	'***************************************************************************
		xsql = "up_andamento_update2 @cd_codigo='"&cd_codigo&"',@cd_os='"&cd_os&"', @cd_cod='"&cd_cod&"', @dt_manut_orc="&dt_manut_orc&",@cd_fornecedor='"&cd_fornecedor&"', @nm_contato='"&nm_contato&"', @dt_prev_resposta="&dt_prev_resposta&", @dt_resposta_orc="&dt_resposta_orc&", @nm_orcamento='"&nm_orcamento&"', @num_os_assist='"&num_os_assist&"', @dt_prev_retorno="&dt_prev_retorno&", @cd_situacao='"&cd_situacao&"', @cd_gestao='"&cd_gestao&"', @comentarios='"&comentarios&"', @dt_retorno="&dt_retorno&", @dt_retorno_un="&dt_retorno_un&", @cd_qtd_valor='"&cd_qtd_valor&"', @cd_valor_unit='"&cd_valor_unit&"', @cd_valor_orc='"&cd_valor_orc&"', @nm_natureza_defeito='"&nm_natureza_defeito&"', @sequencia='"&sequencia&"', @cd_alerta='"&cd_alerta&"'"
		dbconn.execute(xsql)
		action = "editado com sucesso..."
		
						'*************************************************************
						'************ Gera lista de item para patrimoniar ************
						if int(cd_natureza_os) = 2 and nm_orcamento = "2" and cd_gestao = "2" then
							'*** verifica se já consta na lista ***
							strsql = "SELECT * FROM TBL_patrimonio_compra WHERE cd_os = '"&cd_codigo_os&"'"
							Set rs = dbconn.execute(strsql)
							while not rs.EOF 
								cd_os_ver = rs("cd_os")
							rs.movenext
							wend
								if cd_os_ver = "" then
									'*************** Inclui ********************
									xsql = "up_patrimonio_compra_insert @cd_os="&cd_codigo_os&", @cd_user='"&cd_user&"', @dt_data='"&data_atual&"'"
									dbconn.execute(xsql)
								end if
						end if
						
						'*************************************************************
						'***************** Fechamento da O.S. ************************
						If dt_retorno_un <> "Null" Then
							'Modifica os dados - Fecha a O.S.
							xsql = "up_os_fechamento2 @cd_os='"&cd_codigo_os&"', @fecha=1"
							dbconn.execute(xsql)
							
							'strsql = "SELECT * FROM TBL_patrimonio_movimentacao WHERE (cd_patrimonio = 575) AND (cd_tipo_movimentacao = 2) AND (dt_retorno IS NULL)
							
							'xsql = "up_patrimonio_movimentacao_fecha @cd_mov_atual='"&cd_mov_atual&"', @cd_patrimonio='"&cd_pat_codigo&"', @dt_pat_retorno='"&dt_empr_devolucao&"', @cd_user="&cd_user&",@dt_devolucao='"&dt_hoje&"'"
							'dbconn.execute(xsql)
							
						Elseif dt_retorno = "Null" Then 
							'Modifica os dados - Reabre a O.S.
							xsql = "up_os_fechamento2 @cd_os='"&cd_codigo_os&"', @fecha=0"
							dbconn.execute(xsql)
						End if
	
	if jan <> 1 then
		response.redirect("../../manutencao2.asp?tipo=pendencias&strtop=20&filtro=pendentes")
	end if
	link = "../../manutencao2.asp?cd_codigo="&num_os&"&campo=num_os&teste="&teste&""


'*******************************************************************************
'    						Insere - Orçamemto									
'*******************************************************************************
Elseif acao = "1" AND filtro = "orc_cad" Then

	xsql = "up_manutencao_orcamento_insert	@cd_fornecedor="&cd_fornecedor&", @nr_orcamento='"&nr_orcamento&"', @dt_orcamento="&dt_orcamento_aprov&", @nm_valor='"&nm_valor&"', @nr_parcela="&nr_parcela&", @cd_user="&cd_user&""
	Action = "editado com sucesso"
	dbconn.execute(xsql)
	'response.redirect("../../manutencao2.asp?tipo=cad_orc_aprov")

'*******************************************************************************
'    						Altera - Orçamemto									
'*******************************************************************************
Elseif acao = "1" AND filtro = "orc_cad_upd" Then
	xsql = "up_manutencao_orcamento_update @cd_orcamento="&cd_orcamento&",	@cd_fornecedor="&cd_fornecedor&", @nr_orcamento='"&nr_orcamento&"', @dt_orcamento="&dt_orcamento_aprov&", @nm_valor='"&nm_valor&"', @nr_parcela="&nr_parcela&", @cd_user="&cd_user&""
	dbconn.execute(xsql)
	'response.redirect("../../manutencao2.asp?tipo=cad_orc_aprov&cd_orcamento=5")

'*******************************************************************************
'    						Apaga - Orçamento									
'*******************************************************************************

Elseif acao = "3" AND filtro = "orc_cad_del" Then
	xsql = "up_manutencao_orcamento_delete @cd_orcamento='"&cd_orcamento&"'"
	dbconn.execute(xsql)
	
	'DELETE FROM TBL_andamento2 WHERE     (cd_cod = 8464)
	'DELETE FROM TBL_OrdemServico2 WHERE     (cd_codigo = 8470)
	response.redirect("../../manutencao2.asp?tipo=cad_orc_aprov")

'*******************************************************************************
'    						Apaga - O.S.										
'*******************************************************************************

Elseif acao = "3" AND filtro = "apaga" Then
	xsql = "up_os_delete2 @cd_os='"&cd_codigo_os&"'"
	dbconn.execute(xsql)
	'DELETE FROM TBL_andamento2 WHERE     (cd_cod = 8464)
	'DELETE FROM TBL_OrdemServico2 WHERE     (cd_codigo = 8470)
	'response.redirect("../../manutencao2.asp?tipo=pendencias&strtop=20&filtro=pendentes")

Else
action = "erro - teste"
End If

END IF%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>

<body onload="window.close();" onunload="window.opener.location.reload(true);">
<br>
*** Movimentação ***<br>
jan: <%=jan%><br>
<br><br>


xsql = "up_manutencao_orcamento_update<br>
@cd_orcamento=<%=cd_orcamento%><br>
@cd_fornecedor=<%=cd_fornecedor%><br>
@nr_orcamento=<%=nr_orcamento%><br>
@dt_orcamento=<%=dt_orcamento_aprov%><br>
@nm_valor=<%=nm_valor%><br>
@nr_parcela=<%=nr_parcela%><br>
@cd_user=<%=cd_user%><br><br>
<br>





Resposta1: <%=action%><br><br>
os_gravada = <%=os_gravada%><br>
andamento_gravado = <%=andamento_gravado%><br>
<br>
***********************************<br>
***** Fase 1 - cadastramento da O.S. *****<br>
***********************************<br>

Unidade = <%=cd_unidade%><br>
num_os = <%=num_os%><br>
dt_os = <%=dt_os%><br>
Soliciante = <%=nm_solicitante%><br>
Unidade - OS = <%=cd_unidade_os%><br><br>
**********************************************<br>
QTD = <%=num_qtd%><br>
equipamento = <%=cd_equipamento%><br>
Especialidade = <%=cd_especialidade%><br>
Marca = <%=cd_marca%><br>
N/S = <%=cd_ns%><br>
Pat = <%=cd_patrimonio%><br>
************************<br>

num_qtd_atual = <%=num_qtd_atual%><br>
cd_equipamento_atual = <%=cd_equipamento_atual%><br>
cd_especialidade_atual = <%=cd_especialidade_atual%><br>
cd_marca_atual = <%=cd_marca_atual%><br>
cd_ns_atual = <%=cd_ns_atual%><br>
cd_patrimonio_atual = <%=cd_patrimonio_atual%><br>
**********************************************
<br>
<br>
<br>

Motivo = <%=motivo%><br>
<br>
<br>

cd_codigo = <%=cd_codigo%><br>

<br>

dt_envio = <%=dt_entrada%><br>
cd_avaliacao = <%=cd_avaliacao%><br>
cd_fornecedor = <%=cd_fornecedor%><br>
Obs = <%=observacoes%><br>

<br>
<br>

Açao - <%=acao%><br>
filtro - <%=filtro%><br><br>

sequencia - <%=sequencia%><br><br>

Codigo - <%=cd_codigo%><br>
cd_codigo_os = <%=cd_codigo_os%><br>
Fim? - <%=cdt%><br>
Campo - <%=campo%><br>
<br>

OS - <%=num_os%><br>
(OS) - <%=cd_os%><br>
Orcamento - <%=nm_orcamento%><br>
Processo - <%=nm_processo%><br>
Coment - <%=comentarios%><br>
data - <%=dt_orcamento%><br>
dt_aval_data - <%=dt_orcamento_mes%><br>
obs - <%=observacoes%><br>
liberado por - <%=cd_liberacao%><br>
Prev retorno - <%=dt_prev_retorno%><br>
dt_abertura - <%=dt_abertura%><br>
alerta - <%=alerta%><br>

<br>

fecha - <%=fecha%><br>
contato - <%=nm_contato%><br>
obs - <%=observacoes%><br>
conclusao - <%=dt_conclusao%><br>
Tempo Resp - <%=dt_tempo_resp%><br>
equip_filtro <%=cd_equip_filtro%><br>
outro - <%=outro%><br><br>
comentarios = <%=comentarios%><br><br>

cod = <%=cd_cod_ocorrencia%><br>
dt_ocorrencia = <%=dt_ocorrencia%><br>
Ocorrencia = <%=ocorrencia%><br>
<br>
<br>
cd_qtd_valor = <%=cd_qtd_valor%><br>
cd_valor_unit = <%=cd_valor_unit%><br>
cd_valor_orc = <%=cd_valor_orc%><br>
<br>



cd_situacao = <%=cd_situacao%><br>
cd_gestao = <%=cd_gestao%><br><br>
<br>
cd_natureza_os = <%=cd_natureza_os%><br>
nm_orcamento = <%=nm_orcamento%><br>
cd_gestao = <%=cd_gestao%><br>
<br>
teste = <%=teste%><br><br>

cd_user = <%=cd_user%><br>
data = <%=data_atual%><br>
cd_os = <%=cd_codigo_os%><br>


<br>

link: <a href ="<%=link%>">Redireciona</a>

</body>
</html>
