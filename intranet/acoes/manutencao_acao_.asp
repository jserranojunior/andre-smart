
  <!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
   <!--#include file="../includes/inc_area_restrita.asp"-->

<%

acao = request("acao")
filtro = request("filtro")

sequencia = request("sequencia")
cd_codigo = request("cd_codigo")
campo = request("campo")

data_atual = Month(now)&"/"&Day(now)&"/"&Year(now)&" "&Hour(now)&":"&Minute(now)&":"&Second(now)
cd_user = session("cd_codigo")


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
digito = request("digito")
dt_os = request("dt_os")
nm_solicitante  = request("nm_solicitante")
nm_unidade = request("nm_unidade")
nm_uni_select = request("nm_uni_select")
num_qtd = request("num_qtd")
cd_especialidade = request("cd_especialidade")
cd_equipamento = request("cd_equipamento")
cd_marca = request("cd_marca")
cd_ns = request("cd_ns")
cd_patrimonio = request("cd_patrimonio")
motivo = request("motivo")
digito = request("digito")
nm_natureza_defeito = request("nm_natureza_defeito")

dt_entrada = request("dt_entrada")
	if dt_entrada <> "" Then
		dt_entrada = "'"&month(dt_entrada)&"/"&day(dt_entrada)&"/"&year(dt_entrada)&"'"
	else
		'dt_entrada = "NULL"
		dt_entrada = "'"&month(dt_os)&"/"&day(dt_os)&"/"&year(dt_os)&"'"
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

cd_qtd_valor = request("cd_qtd_valor")
cd_valor_unit = request("cd_valor_unit")
	cd_valor_unit = replace(cd_valor_unit,".",",")
cd_valor_orc = request("cd_valor_orc")
	cd_valor_orc = replace(cd_valor_orc,".",",")

If dt_manut_orc = "" Then
dt_manut_orc = "01/01/1950"
End if
If dt_resposta_orc = "" Then
dt_resposta_orc = "01/01/1950"
End if
If dt_prev_resposta = "" Then
dt_prev_resposta = "01/01/1950"
End if
If dt_prev_retorno = "" Then
dt_prev_retorno = "01/01/1950"
'dt_prev_retorno = ""
End if
If dt_retorno_un = "" Then
dt_retorno_un = "01/01/1950"
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
   dt_previsao = "01/01/1950"
End if

If dt_retorno = "" Then
   dt_retorno = "01/01/1950"
End if

If dt_tempo_resp = "" Then
	dt_tempo_resp = "01/01/1950"
End if

If dt_resposta = "" Then
	dt_resposta = "01/01/1950"
End if

If dt_conclusao = "" Then
	dt_conclusao = "01/01/1950"
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

'****************************** Ordem de serviço *********************************************************
'Apagado!!!

'******************************* Os - Fecha *************************************************

if acao = "2" AND filtro = "os" Then
	'Modifica os dados - Fecha a O.S.
	xsql = "up_os_fechamento @num_os='"&num_os&"', @fecha='"&fecha&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	
	response.redirect("../manutencao.asp?tipo=pendencias&cd_codigo="&cd_os&"&campo=num_os")


'******************************* Os - Modifica ??? **********************************************

Elseif acao = "3" AND filtro = "os" Then
	'Modifica os dados
	xsql = "up_os_update @num_os='"&num_os&"', @dt_os='"&Month(dt_os)&"-"&Day(dt_os)&"-"&Year(dt_os)&"', @nm_solicitante='"&nm_solicitante&"', @nm_unidade='"&nm_unidade&"', @num_qtd='"&num_qtd&"', @cd_equipamento='"&cd_equipamento&"', @cd_especialidade='"&cd_especialidade&"', @cd_marca='"&cd_marca&"', @cd_ns='"&cd_ns&"', @cd_patrimonio='"&cd_patrimonio&"', @motivo='"&motivo&"', @dt_entrada='"&Month(dt_entrada)&"-"&Day(dt_entrada)&"-"&Year(dt_entrada)&"', @cd_avaliacao='"&cd_avaliacao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_liberacao='"&cd_liberacao&"', @observacoes='"&observacoes&"',@sequencia='"&sequencia&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	
	response.redirect("../manutencao.asp?tipo=pendencias&cd_codigo="&cd_os&"&campo=num_os")
	link = "manutencao2.asp?cd_codigo="&num_os&"&campo=num_os"


'******************************* Manutençao - Insere ************************************************

Elseif acao = "1" AND filtro = "manutencao" Then
	'Insere dados
	xsql = "up_manutencao_insert @cd_unidade='"&cd_unidade&"',@cd_os='"&cd_os&"', @num_osforn='"&num_osforn&"', @dt_encaminhado='"&Month(dt_encaminhado)&"-"&day(dt_encaminhado)&"-"&year(dt_encaminhado)&"',@cd_fornecedor='"&cd_fornecedor&"', @nm_orcamento='"&nm_orcamento&"', @dt_prev_resposta='"&Month(dt_prev_resposta)&"-"&day(dt_prev_resposta)&"-"&year(dt_prev_resposta)&"', @nm_processo='"&nm_processo&"', @dt_previsao='"&Month(dt_previsao)&"-"&day(dt_previsao)&"-"&year(dt_previsao)&"', @comentarios='"&comentarios&"', @sequencia='"&sequencia&"',@dt_retorno='"&Month(dt_retorno)&"-"&day(dt_retorno)&"-"&year(dt_retorno)&"', @outro=0"
	action = "inserido com sucesso"
	dbconn.execute(xsql)
	
	'response.redirect("manutencao.asp?tipo=pendencias&cd_codigo="&cd_os&"&campo=num_os")
	response.redirect("manutencao.asp?tipo=andamento&cd_codigo="&cd_os&"&campo=num_os&ordem=1")

'******************************* Manutençao - Modifica **********************************************


Elseif acao = "2" AND filtro = "manutencao" Then
	'Modifica os dados
	xsql = "up_manutencao_update @cd_codigo='"&cd_codigo&"', @nm_orcamento='"&nm_orcamento&"', @num_osforn='"&num_osforn&"', @dt_previsao='"&Month(dt_previsao)&"-"&day(dt_previsao)&"-"&year(dt_previsao)&"', @nm_processo='"&nm_processo&"', @comentarios='"&comentarios&"', @dt_retorno='"&Month(dt_retorno)&"-"&day(dt_retorno)&"-"&year(dt_retorno)&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	
	response.redirect("../manutencao.asp?tipo=pendencias&cd_codigo="&cd_os&"&campo=num_os")
	link = "manutencao2.asp?cd_codigo="&cd_os&"&campo=num_os"
	
	'******************************* Manutençao - Fecha *************************************************
	' Apagado, porém já estava inativo !!!
	
	'******************************* Avaliacao - Insere ************************************************


Elseif acao = "1" AND filtro = "avaliacao" Then
	'Insere dados
	'xsql = "up_avaliacao_insert @cd_os='"&cd_os&"', @num_osforn='"&num_osforn&"', @dt_encaminhado='"&Month(dt_encaminhado)&"-"&day(dt_encaminhado)&"-"&year(dt_encaminhado)&"',@cd_fornecedor='"&cd_fornecedor&"', @nm_orcamento='"&nm_orcamento&"', @dt_previsao='"&Month(dt_previsao)&"-"&day(dt_previsao)&"-"&year(dt_previsao)&"', @nm_processo='"&nm_processo&"', @comentarios='"&comentarios&"', @sequencia='"&sequencia&"',@dt_retorno='"&Month(dt_retorno)&"-"&day(dt_retorno)&"-"&year(dt_retorno)&"'"
	xsql = "up_avaliacao_insert @num_os='"&cd_os&"', @dt_aval_data='"&Month(dt_aval_data)&"-"&day(dt_aval_data)&"-"&year(dt_aval_data)&"',@cd_avaliacao='"&cd_avaliacao&"', @cd_fornecedor='"&cd_fornecedor&"', @nm_liberacao='"&nm_liberacao&"', @observacoes='"&observacoes&"',@sequencia='"&sequencia&"'"
	action = "inserido com sucesso"
	dbconn.execute(xsql)
	
	response.redirect("manutencao.asp?tipo=pendencias&cd_codigo="&cd_os&"&campo=num_os")
	link = "manutencao2.asp?cd_codigo="&cd_os&"&campo=num_os"

'******************************* Avaliacao - Insere ************************************************


Elseif acao = "4" AND filtro = "encerramento" Then
	'Modifica os dados (fecha a OS.)
	xsql = "up_os_fecha @num_os='"&num_os&"', @fecha=1"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	
	response.redirect("manutencao.asp?tipo=pendencias&cd_codigo="&cd_os&"&campo=num_os")
	
	
	'******************************* Orçamento - Insere ************************************************
	'????????????????????????????????????????????????????????????????????????????????????????????????????
Elseif acao = "1" AND filtro = "orcamento" Then
	'Insere dados do Orçamento
	xsql = "up_orcamento_insert @cd_os='"&cd_os&"',@dt_orcamento='"&Month(dt_orcamento)&"-"&day(dt_orcamento)&"-"&year(dt_orcamento)&"',@cd_fornecedor='"&cd_fornecedor&"', @nm_contato='"&nm_contato&"', @cd_status='"&cd_status&"', @dt_tempo_resp='"&Month(dt_tempo_resp)&"-"&day(dt_tempo_resp)&"-"&year(dt_tempo_resp)&"',@dt_resposta='"&Month(dt_resposta)&"-"&day(dt_resposta)&"-"&year(dt_resposta)&"', @observacoes='"&observacoes&"', @dt_conclusao='"&Month(dt_conclusao)&"-"&day(dt_conclusao)&"-"&year(dt_conclusao)&"', @outro='"&outro&"', @sequencia='"&sequencia&"'"
	action = "inserido com sucesso"
	dbconn.execute(xsql)
	
	response.redirect("manutencao.asp?tipo=pendencias&cd_codigo="&cd_os&"&campo=num_os")

'******************************* Orçamento - Modifica ***********************************************
'????????????????????????????????????????????????????????????????????????????????????????????????????
Elseif acao = "2" AND filtro = "orcamento" Then
	'Modifica os dados
	xsql = "up_orcamento_update @cd_codigo='"&cd_codigo&"', @cd_status='"&cd_status&"', @dt_resposta='"&Month(dt_resposta)&"-"&day(dt_resposta)&"-"&year(dt_resposta)&"', @observacoes='"&observacoes&"',@dt_conclusao='"&Month(dt_conclusao)&"-"&day(dt_conclusao)&"-"&year(dt_conclusao)&"', @outro='"&outro&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	
	response.redirect("../manutencao.asp?tipo=pendencias&cd_codigo="&cd_os&"&campo=num_os&obs="&observacoes&"")




'****************************************************************************************************
'    Nova fase     
'******************************* cadastro - Ordem de Serviço ****************************************


Elseif acao = "1" And filtro = "cadastra_os" then
	'verifica se a os a ser cadastrada já existe no banco de dados
	'if 	num_os AND cd_unidade
	strsql = "SELECT * FROM TBL_OrdemServico WHERE (cd_unidade = '"&cd_unidade&"') AND (num_os = '"&num_os&"')"
	Set rs = dbconn.execute(strsql)
	while not rs.EOF 
		os_gravada = 1
	rs.movenext
	wend
	
		if os_gravada = "" then
			'Cadastra Ordem de servico
			xsql = "up_os_insert @cd_unidade='"&cd_unidade&"', @num_os='"&num_os&"', @dt_os='"&Month(dt_os)&"-"&Day(dt_os)&"-"&Year(dt_os)&"', @nm_solicitante='"&nm_solicitante&"', @nm_unidade='"&nm_unidade&"', @num_qtd='"&num_qtd&"', @cd_equipamento='"&cd_equipamento&"', @cd_ns='"&cd_ns&"', @cd_patrimonio='"&cd_patrimonio&"', @cd_especialidade='"&cd_especialidade&"', @cd_marca='"&cd_marca&"', @motivo='"&motivo&"', @dt_entrada="&dt_entrada&", @cd_avaliacao='"&cd_avaliacao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_liberacao='"&cd_liberacao&"', @observacoes='"&observacoes&"',@sequencia='"&sequencia&"',@data_inc='"&data_atual&"',@cd_user="&cd_user&",@fecha=0"
			'action = "inserido com sucesso"
			dbconn.execute(xsql)
			
				'Verifica o codigo da O.S recem cadastrada e retorna o codigo para cadastro da data de previsão de entraga no andamento
				strsql = "SELECT * FROM TBL_OrdemServico WHERE (cd_unidade = '"&cd_unidade&"') AND (num_os = '"&num_os&"')"
				Set rs = dbconn.execute(strsql)
				while not rs.EOF 
					os_gravada = 1
					cd_ordemservico = rs("cd_codigo")
				rs.movenext
				wend
				
				
				
				if os_gravada = 1 then
					'verifica a data de entrada e o prazo do fornecedor
						'xsql = "SELECT * FROM TBL_fornecedor where cd_codigo="&cd_fornecedor&""
						'SET rs = dbconn.execute(xsql)
						'	if not rs.EOF then
						'		cd_prazo = rs("cd_prazo")
							'end if
						
						cd_prazo = 15
						
							if cd_prazo > 0 then
								dia_estimado = int(Day(dt_entrada)) + int(cd_prazo)
									mes_estimado = month(dt_entrada)
									ano_estimado = Year(dt_entrada)
									ultimo_dia = ultimodiames(mes_estimado,ano_estimado)
								if 	int(dia_estimado) > int(ultimo_dia) then
									dia_estimado = int(dia_estimado) - int(ultimo_dia)
									mes_estimado = int(mes_estimado) + 1
										if mes_estimado > 12 then
											ano_estimado = int(ano_estimado) + 1
										else
											ano_estimado = ano_estimado
										end if
										dt_prev_retorno = dia_estimado&"/"&mes_estimado&"/"&ano_estimado
								else
									dt_prev_retorno = dia_estimado&"/"&mes_estimado&"/"&ano_estimado
								end if
							
							else
								dt_prev_retorno = Day(dt_entrada)&"/"&Month(dt_entrada)&"/"&Year(dt_entrada)
							end if
							'dt_prev_retorno = "01/01/1950"
				
					'xsql = "up_andamento_insert @cd_os='"&num_os&"', @cd_cod='"&cd_ordemservico&"', @dt_manut_orc='"&Month(dt_manut_orc)&"-"&day(dt_manut_orc)&"-"&year(dt_manut_orc)&"',@cd_fornecedor='"&cd_fornecedor&"', @nm_contato='"&nm_contato&"', @dt_prev_resposta='"&Month(dt_prev_resposta)&"-"&day(dt_prev_resposta)&"-"&year(dt_prev_resposta)&"', @dt_resposta_orc='"&Month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"-"&year(dt_resposta_orc)&"', @nm_orcamento='"&nm_orcamento&"', @num_os_assist='"&num_os_assist&"', @dt_prev_retorno='"&Month(dt_prev_retorno)&"-"&day(dt_prev_retorno)&"-"&year(dt_prev_retorno)&"', @cd_situacao='"&cd_situacao&"', @comentarios='"&comentarios&"', @dt_retorno='"&Month(dt_retorno)&"-"&day(dt_retorno)&"-"&year(dt_retorno)&"', @dt_retorno_un='"&Month(dt_retorno_un)&"-"&day(dt_retorno_un)&"-"&year(dt_retorno_un)&"', @cd_qtd_valor='"&cd_qtd_valor&"', @cd_valor_unit='"&cd_valor_unit&"', @cd_valor_orc='"&cd_valor_orc&"', @nm_natureza_defeito='"&nm_natureza_defeito&"',@sequencia='"&sequencia&"', @cd_alerta='"&cd_alerta&"'"
					xsql = "up_andamento_insert @cd_os='"&num_os&"', @cd_cod='"&cd_ordemservico&"', @dt_manut_orc='"&Month(dt_manut_orc)&"-"&day(dt_manut_orc)&"-"&year(dt_manut_orc)&"',@cd_fornecedor='"&cd_fornecedor&"', @nm_contato='"&nm_contato&"', @dt_prev_resposta='"&Month(dt_prev_resposta)&"-"&day(dt_prev_resposta)&"-"&year(dt_prev_resposta)&"', @dt_resposta_orc='"&Month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"-"&year(dt_resposta_orc)&"', @nm_orcamento='"&nm_orcamento&"', @num_os_assist='"&num_os_assist&"', @dt_prev_retorno='"&Month(dt_prev_retorno)&"-"&day(dt_prev_retorno)&"-"&year(dt_prev_retorno)&"', @cd_situacao='"&cd_situacao&"', @comentarios='"&comentarios&"', @dt_retorno='"&Month(dt_retorno)&"-"&day(dt_retorno)&"-"&year(dt_retorno)&"', @dt_retorno_un='"&Month(dt_retorno_un)&"-"&day(dt_retorno_un)&"-"&year(dt_retorno_un)&"', @cd_qtd_valor='"&cd_qtd_valor&"', @cd_valor_unit='"&cd_valor_unit&"', @cd_valor_orc='"&cd_valor_orc&"', @nm_natureza_defeito='"&nm_natureza_defeito&"',@sequencia='"&sequencia&"', @cd_alerta='"&cd_alerta&"'"
					dbconn.execute(xsql)
				end if
		else
			response.redirect("../manutencao.asp?tipo=andamento&cd_codigo="&num_os&"&campo=num_os&ordem=1")
			'manutencao.asp?tipo=andamento&cd_codigo=6509&campo=cd_codigo
		end if
	
	'response.redirect("../manutencao.asp?tipo=pendencias&strtop=20&filtro=pendentes")
	response.redirect("../manutencao.asp?tipo=andamento&cd_codigo="&num_os&"&campo=num_os&ordem=1")


Elseif acao = "2" AND filtro = "modifica_os" Then
	'Modifica os dados - OK
	xsql = "up_os_update @cd_codigo='"&cd_codigo&"', @num_os='"&num_os&"', @dt_os='"&Month(dt_os)&"-"&Day(dt_os)&"-"&Year(dt_os)&"', @nm_solicitante='"&nm_solicitante&"', @nm_unidade='"&nm_unidade&"', @num_qtd='"&num_qtd&"', @cd_equipamento='"&cd_equipamento&"', @cd_especialidade='"&cd_especialidade&"', @cd_marca='"&cd_marca&"', @cd_ns = '"&cd_ns&"', @cd_patrimonio = '"&cd_patrimonio&"',@motivo='"&motivo&"', @dt_entrada="&dt_entrada&", @cd_avaliacao='"&cd_avaliacao&"', @cd_fornecedor='"&cd_fornecedor&"', @cd_liberacao='"&cd_liberacao&"', @observacoes='"&observacoes&"',@data_inc='"&data_atual&"',@cd_user="&cd_user&",@sequencia='"&sequencia&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	
	response.redirect("../manutencao.asp?tipo=pendencias&cd_codigo="&num_os&"&campo=num_os")
	link = "ordemservico_andamento.asp?cd_codigo="&num_os&"&campo=num_os"
	
	'******************************* andamento - Inserir ************************************************

Elseif acao = "1" AND filtro = "andamento" Then
	xsql = "up_andamento_insert @cd_os='"&cd_os&"', @cd_cod='"&cd_cod&"', @dt_manut_orc='"&Month(dt_manut_orc)&"-"&day(dt_manut_orc)&"-"&year(dt_manut_orc)&"',@cd_fornecedor='"&cd_fornecedor&"', @nm_contato='"&nm_contato&"', @dt_prev_resposta='"&Month(dt_prev_resposta)&"-"&day(dt_prev_resposta)&"-"&year(dt_prev_resposta)&"', @dt_resposta_orc='"&Month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"-"&year(dt_resposta_orc)&"', @nm_orcamento='"&nm_orcamento&"', @num_os_assist='"&num_os_assist&"', @dt_prev_retorno='"&Month(dt_prev_retorno)&"-"&day(dt_prev_retorno)&"-"&year(dt_prev_retorno)&"', @cd_situacao='"&cd_situacao&"', @comentarios='"&comentarios&"', @dt_retorno='"&Month(dt_retorno)&"-"&day(dt_retorno)&"-"&year(dt_retorno)&"', @dt_retorno_un='"&Month(dt_retorno_un)&"-"&day(dt_retorno_un)&"-"&year(dt_retorno_un)&"', @cd_qtd_valor='"&cd_qtd_valor&"', @cd_valor_unit='"&cd_valor_unit&"', @cd_valor_orc='"&cd_valor_orc&"', @nm_natureza_defeito='"&nm_natureza_defeito&"',@sequencia='"&sequencia&"', @cd_alerta='"&cd_alerta&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	
					If dt_retorno_un <> "01/01/1950" Then
					'Modifica os dados - Fecha a O.S.
					xsql = "up_os_fechamento @num_os='"&cd_os&"', @fecha=1"
					action = "editado com sucesso"
					dbconn.execute(xsql)
					End if
	
	response.redirect("../manutencao.asp?tipo=pendencias&strtop=20&filtro=pendentes")
	link = "manutencao.asp?cd_codigo="&num_os&"&campo=num_os"

'******************************* andamento - Alterar ************************************************

Elseif acao = "2" AND filtro = "andamento" Then
	xsql = "up_andamento_update @cd_codigo='"&cd_codigo&"',@cd_os='"&cd_os&"', @cd_cod='"&cd_cod&"', @dt_manut_orc='"&Month(dt_manut_orc)&"-"&day(dt_manut_orc)&"-"&year(dt_manut_orc)&"',@cd_fornecedor='"&cd_fornecedor&"', @nm_contato='"&nm_contato&"', @dt_prev_resposta='"&Month(dt_prev_resposta)&"-"&day(dt_prev_resposta)&"-"&year(dt_prev_resposta)&"', @dt_resposta_orc='"&Month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"-"&year(dt_resposta_orc)&"', @nm_orcamento='"&nm_orcamento&"', @num_os_assist='"&num_os_assist&"', @dt_prev_retorno='"&Month(dt_prev_retorno)&"-"&day(dt_prev_retorno)&"-"&year(dt_prev_retorno)&"', @cd_situacao='"&cd_situacao&"', @comentarios='"&comentarios&"', @dt_retorno='"&Month(dt_retorno)&"-"&day(dt_retorno)&"-"&year(dt_retorno)&"', @dt_retorno_un='"&Month(dt_retorno_un)&"-"&day(dt_retorno_un)&"-"&year(dt_retorno_un)&"', @cd_qtd_valor='"&cd_qtd_valor&"', @cd_valor_unit='"&cd_valor_unit&"', @cd_valor_orc='"&cd_valor_orc&"', @nm_natureza_defeito='"&nm_natureza_defeito&"', @sequencia='"&sequencia&"', @cd_alerta='"&cd_alerta&"'"
	action = "editado com sucesso"
	dbconn.execute(xsql)
	
					'***************** Fechamento da O.S. ************************
	
					If dt_retorno_un <> "01/01/1950" Then
					'Modifica os dados - Fecha a O.S.
					'xsql = "up_os_fechamento @num_os='"&cd_os&"', @unidade='"&&"', @fecha=1"
					xsql = "up_os_fechamento @cd_os='"&cd_codigo_os&"', @fecha=1"
					dbconn.execute(xsql)
					Elseif dt_retorno = "01/01/1950" Then 
					'Modifica os dados - Reabre a O.S.
					'xsql = "up_os_fechamento @num_os='"&cd_os&"', @unidade='"&&"', @fecha=null"
					xsql = "up_os_fechamento @cd_os='"&cd_codigo_os&"', @fecha=0"
					dbconn.execute(xsql)
					End if
	
	response.redirect("../manutencao.asp?tipo=pendencias&strtop=20&filtro=pendentes")
	link = "../manutencao.asp?cd_codigo="&num_os&"&campo=num_os"

'****************************************************************************************************
Else
action = "erro"
End If



'='"&Month()&"-"&day()&"-"&year()&"'


'response.redirect("manutencao.asp?cd_codigo="&cd_codigo&"&campo="&campo&"")

'response.redirect("manutencao.asp?cd_codigo=&campo=dt_os")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>

<body>



Resposta1: <%=action%><br><br>

Açao - <%=acao%><br>
filtro - <%=filtro%><br><br>

sequencia - <%=sequencia%><br><br>

Codigo - <%=cd_codigo%><br>
cd_codigo_os = <%=cd_codigo_os%><br>
Fim? - <%=cdt%><br>
Campo - <%=campo%><br>
OS - <%=num_os%><br>
(OS) - <%=cd_os%><br>
Orcamento - <%=nm_orcamento%><br>
Processo - <%=nm_processo%><br>
Coment - <%=comentarios%><br>
data - <%=dt_orcamento%><br>
dt_aval_data - <%=dt_orcamento_mes%><br>
obs - <%=observacoes%><br>
liberado por - <%=cd_liberacao%><br>


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



cd_situacao = <%=cd_situacao%>

link: <a href ="<%=link%>">Redireciona</a>

</body>
</html>
