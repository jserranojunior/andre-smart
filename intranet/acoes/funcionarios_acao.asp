<%@ Language=VBScript %>
<%

'Set Upload = Server.CreateObject("Persits.Upload.1")

'Count = Upload.Save("E:\users\vdlap.com.br\website\foto\temp\")
%>

<!--#include file="../includes/inc_open_connection.asp"-->

<%
strcod    = request("cod")
stracao   = request("acao")

strregime_trabalho = request("cd_regime_trabalho")
strcd_matricula = request("cd_matricula")
strnome =  request("nm_nome")
strnm_sobrenome =  request("nm_sobrenome")
strfoto = request("foto_h")
Strdt_contratado =  request("dt_contratado")
Strdt_demissao =  request("dt_demissao")
strdt_nascimento = request("dt_nascimento")
strnm_email =  request("nm_email")
strnm_rg =  request("nm_rg")
strnm_cpf =  request("nm_cpf")
strmae = request("nm_mae")
strpai = request("nm_pai")
strnm_ddd = request("nm_ddd")
strnm_fone =  request("nm_fone")
strnm_ddd_cel = request("nm_ddd_cel")
strnm_cel =  request("nm_cel")
strocontato = request("nm_o_contato")
strnm_endereco = request("nm_endereco")
strnr_numero =  request("nr_numero")
strnm_complemento = request("nm_complemento")
strnm_cidade =  request("nm_cidade")
strnm_estado =  request("nm_estado")
strnm_cep = request("nm_cep")
strcd_funcao = request("cd_funcao")
strstatus =  request("cd_status")
strcd_unidades = request("cd_unidades")
strbairro = request("nm_bairro")
str_nacionalidade = request("nm_nacionalidade")
str_escolaridade = request("nm_escolaridade")
str_pis = request("cd_pis")
str_cd_estado_civil = request("cd_estado_civil")
str_ctps = request("cd_ctps")
str_ctps_serie = request("cd_ctps_serie")
strqtd_dependentes = request("cd_qtd_dependentes")
strqtd_filhos = request("cd_qtd_filhos")
str_estado_civil = request("cd_estado_civil")
str_carga_horaria = request("cd_carga_horaria")
str_carga_diaria = request("cd_carga_diaria")
str_assistencia_medica = request("cd_assistencia_medica")
str_assistencia_medica_matricula = request("cd_assistencia_medica_matricula")
str_assistencia_odonto = request("cd_assistencia_odonto")
str_assistencia_odonto_matricula = request("cd_assistencia_odonto_matricula")
str_vr = request("cd_vr")
str_vale_alimentacao = request("cd_vale_alimentacao")
str_cmt_bom = request("cd_cmt_bom")
str_sptrans = request("cd_sptrans")
str_funcao = request("cd_funcao")
str_status = request("cd_status")


'dt_demissao = request("dt_demissao")
'dt_desliga = request("dt_desliga")
cd_dia = request("cd_dia")
cd_mes = request("cd_mes")
cd_ano = request("cd_ano")
	cd_hora = hour(now)
	cd_minuto = minute(now)
	cd_segundo = second(now)

	dt_desliga = cd_mes&"/"&cd_dia&"/"&cd_ano
	dt_atualizacao = cd_mes&"/"&cd_dia&"/"&cd_ano&" "&cd_hora&":"&cd_minuto&":"&cd_segundo
	
	'dt_desliga = cd_dia&"/"&cd_mes&"/"&cd_ano
	'dt_desliga = "12/09/2008"

str_banco = request("nm_banco")
str_banco_ag = request("cd_banco_ag")
str_banco_conta = request("cd_banco_conta")
str_banco_conta_tipo = request("cd_banco_conta_tipo")
nm_tit_eleitor = request("nm_tit_eleitor")
nm_reservista = request("nm_reservista")
nm_creche = request("nm_creche")

dt_ano = YEAR(Now)
dt_mes = Month(Now)
dt_dia = Day(now)
dt_hora = hour(now)
dt_minuto = minute(now)

'**********************************************************************
'*** 					Insere funcionario						    ***
'**********************************************************************
If stracao = "insert" Then
	xsql = "up_Funcionario_insert @cd_regime_trabalho='"&strregime_trabalho&"',@cd_matricula='"&strcd_matricula&"',@nm_nome='"&strnome&"',@nm_sobrenome='"&strnm_sobrenome&"', @nm_foto='"&strfoto&"', @dt_contratado='"&Month(Strdt_contratado)&"-"&day(Strdt_contratado)&"-"&year(Strdt_contratado)&"', @dt_nascimento='"&Month(strdt_nascimento)&"-"&day(strdt_nascimento)&"-"&year(strdt_nascimento)&"', @nm_email='"&strnm_email&"', @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"', @nm_pai='"&strpai&"', @nm_ddd='"&strnm_ddd&"', @nm_fone='"&strnm_fone&"', @nm_ddd_cel='"&strnm_ddd_cel&"', @nm_o_contato='"&strocontato&"',@nm_cel='"&strnm_cel&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_complemento='"&strnm_complemento&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"',@nm_cep='"&strnm_cep&"', @cd_status='"&strstatus&"', @nm_bairro = '"&strbairro&"', @nm_nacionalidade='"&str_nacionalidade&"',@nm_escolaridade='"&str_escolaridade&"',@cd_pis='"&str_pis&"',@cd_ctps='"&str_ctps&"',@cd_ctps_serie='"&str_ctps_serie&"',@cd_qtd_dependentes='"&strqtd_dependentes&"', @cd_qtd_filhos='"&strqtd_filhos&"', @cd_estado_civil='"&str_estado_civil&"', @cd_carga_horaria='"&str_carga_horaria&"', @cd_carga_diaria='"&str_carga_diaria&"', @cd_assistencia_medica='"&str_assistencia_medica&"', @cd_assistencia_medica_matricula='"&str_assistencia_medica_matricula&"', @cd_assistencia_odonto='"&str_assistencia_odonto&"', @cd_assistencia_odonto_matricula='"&str_assistencia_odonto_matricula&"', @cd_vr='"&str_vr&"',@cd_vale_alimentacao='"&str_vale_alimentacao&"',@cd_sptrans='"&str_sptrans&"',@cd_cmt_bom='"&str_cmt_bom&"', @nm_banco='"&str_banco&"', @cd_banco_ag='"&str_banco_ag&"', @cd_banco_conta='"&str_banco_conta&"',@nm_tit_eleitor='"&nm_tit_eleitor&"',@nm_reservista='"&nm_reservista&"',@nm_creche='"&nm_creche&"'"
	dbconn.execute(xsql)
	
		'Cargo: Verifica o nome do funcionario e resgata seu código
		strsql = "Select * From TBL_funcionario Where nm_nome='"&strnome&"' AND nm_sobrenome='"&strnm_sobrenome&"'"
		Set rs = dbconn.execute(strsql)
			if Not rs.EOF Then
			cd_cod_funcionario = rs("cd_codigo")
			end if
		
		'Insere o cargo do funcionario recem cadastrado
			xsql = "up_funcionario_cargo_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_cargo='"&strcd_funcao&"', @dt_atualizacao='"&Month(Strdt_contratado)&"-"&day(Strdt_contratado)&"-"&year(Strdt_contratado)&" "&dt_hora&":"&dt_minuto&"'"
			dbconn.execute(xsql)
			
		'Insere a unidade do funcionário
			xsql = "up_funcionario_unidade_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_unidade='"&strcd_unidades&"', @dt_atualizacao='"&Month(Strdt_contratado)&"-"&day(Strdt_contratado)&"-"&year(Strdt_contratado)&" "&dt_hora&":"&dt_minuto&"'"
			dbconn.execute(xsql)
					'Atualiza a sigla da unidade na tabela de Funcionarios
					xsql="Select nm_sigla from TBL_unidades Where cd_codigo='"&strcd_unidades&"'"
					Set rs = dbconn.execute(xsql)
					nm_sigla = rs("nm_sigla")
					
					xsql = "UPDATE TBL_funcionario SET nm_sigla='"&nm_sigla&"' WHERE cd_codigo='"&cd_cod_funcionario&"'"
					dbconn.execute(xsql)
					
		'Insere o status do funcionário
			xsql = "up_funcionario_status_insert @cd_funcionario='"&cd_cod_funcionario&"', @cd_status='"&strstatus&"', @dt_atualizacao='"&Month(Strdt_contratado)&"-"&day(Strdt_contratado)&"-"&year(Strdt_contratado)&" "&dt_hora&":"&dt_minuto&"'"
			dbconn.execute(xsql)

			'Registra o mes posterior à sua entrada para pagamento do sindicato
			dt_ano_post = year(Strdt_contratado)
			dt_mes_post = month(Strdt_contratado)+1
				if dt_mes_post = 13 Then
				dt_mes_post = 1
				dt_ano_post = dt_ano_post + 1
				end if
				
			xsql = "up_funcionario_sindicato_insert @cd_funcionario='"&cd_cod_funcionario&"', @dt_mes='"&Month(dt_mes_post)&"', @dt_ano='"&year(dt_ano_post)&"'"
			dbconn.execute(xsql)
			response.redirect "../empresa.asp?tipo=lista&msg="&strmsg&""


'**********************************************************************
'*** 					Altera funcionario						    ***
'**********************************************************************

ElseIf stracao = "altera" Then
	if strstatus <> 4 Then ' Altera
	'xsql = "up_Funcionario_update @cd_regime_trabalho = '"&strregime_trabalho&"',@cd_matricula='"&strcd_matricula&"',@nm_nome='"&strnome&"',@nm_sobrenome='"&strnm_sobrenome&"', @nm_foto='"&strfoto&"', @dt_contratado='"&Month(Strdt_contratado)&"-"&day(Strdt_contratado)&"-"&year(Strdt_contratado)&"', @dt_nascimento='"&Month(strdt_nascimento)&"-"&day(strdt_nascimento)&"-"&year(strdt_nascimento)&"', @nm_email='"&strnm_email&"', @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"', @nm_pai='"&strpai&"',@nm_ddd='"&strnm_ddd&"', @nm_fone='"&strnm_fone&"', @nm_ddd_cel='"&strnm_ddd_cel&"', @nm_cel='"&strnm_cel&"', @nm_o_contato='"&strocontato&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_bairro = '"&strbairro&"', @nm_complemento='"&strnm_complemento&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"',@nm_cep='"&strnm_cep&"', @cd_codigo='"&strcod&"', @nm_nacionalidade='"&str_nacionalidade&"',@nm_escolaridade='"&str_escolaridade&"',@cd_pis='"&str_pis&"',@cd_ctps='"&str_ctps&"',@cd_ctps_serie='"&str_ctps_serie&"',@cd_qtd_dependentes='"&strqtd_dependentes&"', @cd_qtd_filhos='"&strqtd_filhos&"', @cd_estado_civil='"&str_estado_civil&"', @cd_assistencia_medica='"&str_assistencia_medica&"',@cd_carga_horaria='"&str_carga_horaria&"', @cd_carga_diaria='"&str_carga_diaria&"', @cd_assistencia_medica_matricula='"&str_assistencia_medica_matricula&"', @cd_vr='"&str_vr&"',@cd_vale_alimentacao='"&str_vale_alimentacao&"',@cd_sptrans='"&str_sptrans&"',@cd_cmt_bom='"&str_cmt_bom&"', @nm_banco = '"&str_banco&"', @cd_banco_ag = '"&str_banco_ag&"', @cd_banco_conta='"&str_banco_conta&"', @nm_tit_eleitor='"&nm_tit_eleitor&"',@nm_reservista='"&nm_reservista&"',@nm_creche='"&nm_creche&"',@cd_assistencia_odonto='"&str_assistencia_odonto&"',@cd_assistencia_odonto_matricula='"&str_assistencia_odonto_matricula&"'"
	xsql = "up_Funcionario_update @cd_regime_trabalho = '"&strregime_trabalho&"',@cd_matricula='"&strcd_matricula&"',@nm_nome='"&strnome&"',@nm_sobrenome='"&strnm_sobrenome&"', @nm_foto='"&strfoto&"', @dt_contratado='"&Month(Strdt_contratado)&"-"&day(Strdt_contratado)&"-"&year(Strdt_contratado)&"', @dt_nascimento='"&Month(strdt_nascimento)&"-"&day(strdt_nascimento)&"-"&year(strdt_nascimento)&"', @nm_rg='"&strnm_rg&"', @nm_cpf='"&strnm_cpf&"', @nm_mae='"&strmae&"', @nm_pai='"&strpai&"', @nm_endereco='"&strnm_endereco&"', @nr_numero='"&strnr_numero&"', @nm_bairro = '"&strbairro&"', @nm_complemento='"&strnm_complemento&"', @nm_cidade='"&strnm_cidade&"', @nm_estado='"&strnm_estado&"',@nm_cep='"&strnm_cep&"', @cd_codigo='"&strcod&"', @nm_nacionalidade='"&str_nacionalidade&"',@cd_pis='"&str_pis&"',@cd_ctps='"&str_ctps&"',@cd_ctps_serie='"&str_ctps_serie&"', @cd_qtd_filhos='"&strqtd_filhos&"', @cd_estado_civil='"&str_estado_civil&"', @cd_assistencia_medica_matricula='"&str_assistencia_medica_matricula&"', @cd_vr='"&str_vr&"',@cd_vale_alimentacao='"&str_vale_alimentacao&"',@cd_sptrans='"&str_sptrans&"',@cd_cmt_bom='"&str_cmt_bom&"', @nm_banco = '"&str_banco&"', @cd_banco_ag = '"&str_banco_ag&"', @cd_banco_conta='"&str_banco_conta&"', @nm_tit_eleitor='"&nm_tit_eleitor&"',@nm_reservista='"&nm_reservista&"',@nm_creche='"&nm_creche&"'"
	dbconn.execute(xsql)
	strmsg = "Colaborador Alterado com sucesso!"	
	response.redirect "../empresa.asp?tipo=lista&msg="&strmsg&""
	
	Elseif strstatus = 4 Then 'Dispensa
		xsql_demissao = "UPDATE TBL_Funcionario SET dt_desliga = '"&dt_desliga&"' WHERE cd_codigo = "&strcod&""
		dbconn.execute(xsql_demissao)
	strmsg = "Colaborador dispensado! "&dt_desliga&""
				'*** Atualiza o status do funcionário ***'
				xsql = "up_funcionario_status_insert @cd_funcionario='"&strcod&"', @cd_status='4', @dt_atualizacao='"&dt_desliga&" "&dt_hora&":"&dt_minuto&"'"
			dbconn.execute(xsql)
	end if


'**********************************************************************
'*** 					Insere funcionario						    ***
'**********************************************************************

Elseif stracao = "excluir" Then
    xsql= "up_Funcionario_delete @cd_codigo='"&strcod&"'"
	dbconn.execute(xsql)
		'Apaga o histórico de cargos do funcionário
		xsql= "up_Funcionario_cargo_delete @cd_funcionario='"&strcod&"'"
		dbconn.execute(xsql)
		'Apaga o histórico de unidades do funcionário
		xsql= "up_Funcionario_unidade_delete @cd_funcionario='"&strcod&"'"
		dbconn.execute(xsql)
		'Apaga o histórico de status do funcionário
		xsql= "up_Funcionario_status_delete @cd_funcionario='"&strcod&"'"
		dbconn.execute(xsql)
		
	strmsg = "funcionario excluido com sucesso!!!"
	response.write("exluiiu")
	response.redirect "../empresa.asp?tipo=lista&msg="&strmsg&""


'******************************************************************
'*** 					Altera Cargos						    ***
'******************************************************************

Elseif stracao = "cargos" Then
	'Insere cargo
	xsql = "up_funcionario_cargo_insert @cd_funcionario='"&strcod&"', @cd_cargo='"&str_funcao&"', @dt_atualizacao='"&dt_atualizacao&"'"
	dbconn.execute(xsql)
	response.write("cargo OK")

'******************************************************************
'*** 					Altera Unidades						    ***
'******************************************************************

ElseIf stracao = "unidade" Then
	'Insere dados
	xsql = "up_funcionario_unidade_insert @cd_funcionario='"&strcod&"', @cd_unidade='"&strcd_unidades&"', @dt_atualizacao='"&dt_atualizacao&"'"
	dbconn.execute(xsql)
	'response.redirect("assistencia.asp?action="&action&"")
		xsql="Select nm_sigla from TBL_unidades Where cd_codigo='"&strcd_unidades&"'"
		Set rs = dbconn.execute(xsql)
		nm_sigla = rs("nm_sigla")
			'Atualiza a sigla da unidade na tabela de Funcionarios
			xsql = "UPDATE TBL_funcionario SET nm_sigla='"&nm_sigla&"' WHERE cd_codigo='"&strcod&"'"
			dbconn.execute(xsql)


'******************************************************************
'*** 					Altera Staus						    ***
'******************************************************************

ElseIf stracao = "status" Then
	'Insere dados
	xsql = "up_funcionario_status_insert @cd_funcionario='"&strcod&"', @cd_status='"&str_status&"', @dt_atualizacao='"&dt_atualizacao&"'"
	dbconn.execute(xsql)
	'response.redirect("assistencia.asp?action="&action&"")

'*******************************************************************


Else
response.write("erro")
End if

'Set Upload = Nothing 
'response.redirect "../empresa.asp?tipo=lista&msg="&strmsg&""
%>

   

   
<body onload="window.close()">
<!--body-->
<br>
código: <%=strcod%><br>
ação: <%=stracao%><br>
cargo: <%=str_funcao%><br>
atual: <%=dt_atualizacao%><br>
Status: <%=str_status%><br>
<%=strcd_matricula%><br>
<%=strnome%><br>
<%=strnm_sobrenome%><br>
<%=Strdt_contratado%><br>
<%=strmsg%><br>
<%=strfoto%><br>
<%=str_assistencia_medica_tipo%><br>
<%=str_banco%><br>
<%=str_cd_banco_ag%><br>
<%=str_cd_banco_conta%><br>
<%=str_banco_conta_tipo%><br>
status: <%=strstatus%><br>
Demissao: <%=dt_desliga%>(<%=cd_dia%>-<%=cd_mes%>-<%=cd_ano%><br>
Hora: <%=dt_hora%><br>
Minuto: <%=dt_minuto%><br>
<br>
unidade: <%=strcd_unidades%><br>
sigla: <%=nm_sigla%>


</body>