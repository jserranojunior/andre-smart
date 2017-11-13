<!--#include file="../includes/inc_open_connection.asp"-->
   <!--include file="../includes/inc_area_restrita.asp"-->
<%
'***** Variáveis ******

acao = request("acao") 
cd_unidade = request("cd_unidade")
cd_protocolo = request("cd_protocolo")
cd_digito = request("cd_digito")

nm_nome = request("nm_nome")
cd_idade = request("cd_idade")
nm_registro = request("nm_registro")
cd_convenio = request("cd_convenio")
	cd_convenio_1 = request("cd_convenio_1")
nm_sexo = request("nm_sexo")
nm_cirug_realizada = request("nm_cirug_realizada")

	dt_dia_cirurgia = request("dt_dia_cirurgia")
	dt_mes_cirurgia = request("dt_mes_cirurgia")
	dt_ano_cirurgia = request("dt_ano_cirurgia")
dt_cirurgia = dt_mes_cirurgia&"/"&dt_dia_cirurgia&"/"&dt_ano_cirurgia

nm_agenda_cirurgia = request("nm_agenda_cirurgia") 

	dt_hora_agenda = request("dt_hora_agenda")
	dt_minuto_agenda = request("dt_minuto_agenda")
dt_hr_agenda = dt_hora_agenda&":"&dt_minuto_agenda&":00"

	dt_hora_inicio = request("dt_hora_inicio")
	dt_minuto_inicio = request("dt_minuto_inicio")
dt_hr_inicio = dt_hora_inicio&":"&dt_minuto_inicio

	dt_hora_fim = request("dt_hora_fim")
	dt_minuto_fim = request("dt_minuto_fim")
dt_hr_fim = dt_hora_fim&":"&dt_minuto_fim

cd_crm = request("cd_crm")
	cd_crm_1 = request("cd_crm_1")

cd_rack = request("cd_rack")
	cd_rack_1 = request("cd_rack_1")

cd_especialidade = request("cd_especialidade")
	cd_especialidade_1 = request("cd_especialidade_1")

cd_procedimento = request("cd_procedimento")
	cd_procedimento_1 = request("cd_procedimento_1")

procedimentos = request("procedimentos")

'************ Condições ************************
if cd_convenio_1 <> "" Then
	cd_convenio = cd_convenio_1
end if

if cd_crm_1 <> "" Then
	cd_crm = cd_crm_1
end if

if cd_rack_1 <> "" Then
	cd_rack = cd_rack_1
end if

if cd_especialidade_1 <> "" Then
	cd_especialidade = cd_especialidade_1
end if

if cd_procedimento_1 <> "" Then
	cd_procedimento = cd_procedimento_1
end if 

if dt_hr_agenda = ":" Then
	dt_hr_agenda = NULL
End if

if dt_hr_inicio = ":" Then
	dt_hr_inicio = NULL
End if

if dt_hr_fim = ":" Then
	dt_hr_fim = NULL
End if

'***** verifica se já existe no Banco de Dados *****
		'xsql ="SELECT cd_protocolo FROM TBL_PROTOCOLO WHERE cd_protocolo = '"&cd_protocolo&"'"
		xsql ="SELECT A002_numpro FROM TBL_PROTOCOLO WHERE A002_numpro = '"&cd_protocolo&"'"
		Set rs = dbconn.execute(xsql)
			Do while NOT rs.EOF
			cd_prot = rs("A002_numpro")
			rs.movenext
			Loop	

'if cd_prot = "" Then
	
	'***** Ações ******
	If acao = "inserir" Then
	'xsql ="up_protocolo_insert @cd_protocolo='"&int(cd_protocolo)&"', @cd_unidade='"&cd_unidade&"',@nm_nome='"&nm_nome&"', @cd_idade='"&cd_idade&"', @nm_registro='"&nm_registro&"',@cd_convenio='"&cd_convenio&"',@nm_sexo='"&nm_sexo&"', @nm_cirug_realizada='"&nm_cirug_realizada&"',@dt_cirurgia='"&dt_cirurgia&"', @nm_agenda_cirurgia='"&nm_agenda_cirurgia&"',@dt_hr_agenda='"&dt_hr_agenda&"',@dt_hr_inicio='"&dt_hr_inicio&"',@dt_hr_fim='"&dt_hr_fim&"',@cd_crm='"&cd_crm&"',@cd_rack='"&cd_rack&"',@cd_especialidade='"&cd_especialidade&"',@procedimentos='"&procedimentos&"'"
	xsql ="up_protocolo_insert @cd_protocolo='"&int(cd_protocolo)&"', @cd_unidade='"&cd_unidade&"',@nm_nome='"&nm_nome&"', @cd_idade='"&cd_idade&"', @nm_registro='"&nm_registro&"',@cd_convenio='"&cd_convenio&"',@nm_sexo='"&nm_sexo&"', @nm_cirug_realizada='"&nm_cirug_realizada&"',@dt_cirurgia='"&dt_cirurgia&"', @nm_agenda_cirurgia='"&nm_agenda_cirurgia&"',@dt_hr_agenda='"&dt_hr_agenda&"',@dt_hr_inicio='"&dt_hr_inicio&"',@dt_hr_fim='"&dt_hr_fim&"',@cd_crm='"&cd_crm&"',@cd_rack='"&cd_rack&"',@cd_especialidade='"&cd_especialidade&"'"',@procedimentos='"&procedimentos&"'"
	dbconn.execute(xsql)
	response.write("OK, inserido")
	response.redirect "../protocolo.asp"
	
	
	Elseif acao = "editar" Then
	xsql ="up_protocolo_update @cd_protocolo='"&int(cd_protocolo)&"', @cd_unidade='"&cd_unidade&"',@nm_nome='"&nm_nome&"', @cd_idade='"&cd_idade&"', @nm_registro='"&nm_registro&"',@cd_convenio='"&cd_convenio&"',@nm_sexo='"&nm_sexo&"', @nm_cirug_realizada='"&nm_cirug_realizada&"',@dt_cirurgia='"&dt_cirurgia&"', @nm_agenda_cirurgia='"&nm_agenda_cirurgia&"',@dt_hr_agenda='"&dt_hr_agenda&"',@dt_hr_inicio='"&dt_hr_inicio&"',@dt_hr_fim='"&dt_hr_fim&"',@cd_crm='"&cd_crm&"',@cd_rack='"&cd_rack&"',@cd_especialidade='"&cd_especialidade&"',@procedimentos='"&procedimentos&"'"
	dbconn.execute(xsql)
	response.write("OK, alterado")
	response.redirect "../protocolo.asp?tipo=ver&cd_protocolo="&int(cd_protocolo)&""
	
	'ElseIf acao = "excluir" Then
	'xsql ="up_cartao_horas_excluir @cd_codigo='"&cd_codigo&"',@cd_cod_horas='"&cd_cod_horas&"'"
	'dbconn.execute xsql
	'response.write("OK, excluído")
	'response.redirect "../produtividade.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""
	
	Else
	
	response.write("Erro, não passou")
	End if
'Else
'	response.redirect("../protocolo.asp")
'end if
%>

<br>
acao: <%=acao%><br>
Unidade: <%=cd_unidade%><br>
protocolo: <%=cd_protocolo%>-<%=cd_digito%><br><br>


Nome: <%=nm_nome%><br>
Idade: <%=cd_idade%><br>
registro: <%=nm_registro%><br>
cd_convenio: <%=cd_convenio%><br>
sexo: = <%=nm_sexo%><br>
cirug_realizada?: <%=nm_cirug_realizada%><br>
dt_cirurgia: <%=dt_cirurgia%><br>
nm_agenda_cirurgia: <%=nm_agenda_cirurgia%><br>
hr_agenda: <%=dt_hr_agenda%><br>
hora_inicio: <%=dt_hr_inicio%><br>
hora_fim: <%=dt_hr_fim%><br>
CRM: <%=cd_crm%><br>
cd_rack: <%=cd_rack%><br>
Especialidade: <%=cd_especialidade%><br>
procedimentos: <%=procedimentos%><br>

