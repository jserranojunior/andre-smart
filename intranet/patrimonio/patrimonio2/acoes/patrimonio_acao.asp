
  <!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
   <!--#include file="../../includes/inc_area_restrita.asp"-->

<%
acao = request("acao")
cd_tipo = request("cd_tipo")
cd_apaga = request("cd_apaga")

cd_pat_codigo = request("cd_pat_codigo")
cd_equipamento = request("cd_equipamento")
cd_pn = request("cd_pn")
cd_ns = request("cd_ns")
cd_patrimonio = request("cd_patrimonio")
nm_tipo = request("nm_tipo")
cd_especialidade = request("cd_especialidade")

nm_motivo = request("nm_motivo")
cd_unidade = request("cd_unidade")
cd_rack = request("cd_rack")
num_hospital = request("num_hospital")
cd_status = request("cd_status")
cd_marca =request("cd_marca")

cd_dia = request("cd_dia")
cd_mes = request("cd_mes")
cd_ano = request("cd_ano")

cd_preventiva = request("cd_preventiva")
	if cd_preventiva = "" Then
		cd_preventiva = 0
	end if
dt_periodicidade_prev = request("dt_periodicidade_prev")
	if dt_periodicidade_prev = "" Then
		dt_periodicidade_prev = "NULL"
	end if
'dt_plan_mes_cal = request("dt_plan_mes_cal")
'dt_plan_ano_cal = request("dt_plan_ano_cal")

cd_calibracao = request("cd_calibracao")
	if cd_calibracao = "" Then
		cd_calibracao = 0
	end if
dt_periodicidade_cal = request("dt_periodicidade_cal")
	if dt_periodicidade_cal = "" Then
		dt_periodicidade_cal = "NULL"
	end if
'dt_plan_mes_cal = request("dt_plan_mes_cal")
'dt_plan_ano_cal = request("dt_plan_ano_cal")

cd_seg_elet = request("cd_seg_elet")
	if cd_seg_elet = "" Then
		cd_seg_elet = 0
	end if
dt_periodicidade_elet = request("dt_periodicidade_elet")
	if dt_periodicidade_elet = "" Then
		dt_periodicidade_elet = "NULL"
	end if
'dt_plan_mes_cal = request("dt_plan_mes_cal")
'dt_plan_ano_cal = request("dt_plan_ano_cal")

cd_hora = hour(now)
cd_minuto = minute(now)
cd_segundo = second(now)

	dt_atualizacao = cd_mes&"/"&cd_dia&"/"&cd_ano&" "&cd_hora&":"&cd_minuto&":"&cd_segundo
	dt_data = dt_atualizacao
	dt_descarte = dt_atualizacao

voltar = request("voltar")

if cd_apaga <> "" Then
	acao = "apaga"
end if

'Cargo
cd_funcao = request("cd_funcao")

IF x = "" then
' ****  AÇÕES  ****
	'Patrimonio
	If acao = "inserir" AND cd_tipo = "patrimonio" Then
	'Insere dados
	xsql = "up_patrimonio_insert @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @cd_patrimonio='"&cd_patrimonio&"', @nm_tipo='"&nm_tipo&"', @cd_especialidade='"&cd_especialidade&"', @cd_unidade='"&cd_unidade&"', @cd_rack='"&cd_rack&"', @num_hospital='"&num_hospital&"', @dt_data='"&dt_data&"', @cd_preventiva="&cd_preventiva&", @dt_periodicidade_prev="&dt_periodicidade_prev&", @cd_calibracao="&cd_calibracao&", @dt_periodicidade_cal="&dt_periodicidade_cal&",	@cd_seg_elet="&cd_seg_elet&", @dt_periodicidade_elet="&dt_periodicidade_elet&""
	dbconn.execute(xsql)
	'response.write("patrimonio insere OK")
		'*** Recupera o codigo recem criado e agenda as manutenções***
			'*** PREVENTIVA ***
			'*** CALIBRAÇÃO ***
			'*** SEG. ELETRICA ***

	response.redirect("../../patrimonio.asp?tipo=cadastro&mensagem=Inserido com sucesso")
	
	ElseIf acao = "editar" AND cd_tipo = "patrimonio" Then
	'Modifica os dados
	xsql = "up_patrimonio_update @cd_pat_codigo='"&cd_pat_codigo&"', @cd_equipamento='"&cd_equipamento&"', @cd_marca='"&cd_marca&"', @cd_pn='"&cd_pn&"', @cd_ns='"&cd_ns&"', @cd_patrimonio='"&cd_patrimonio&"', @nm_tipo='"&nm_tipo&"', @cd_especialidade='"&cd_especialidade&"', @cd_unidade='"&cd_unidade&"', @cd_rack='"&cd_rack&"', @num_hospital='"&num_hospital&"', @dt_data='"&dt_data&"', @cd_preventiva="&cd_preventiva&", @dt_periodicidade_prev="&dt_periodicidade_prev&", @cd_calibracao="&cd_calibracao&", @dt_periodicidade_cal="&dt_periodicidade_cal&", @cd_seg_elet="&cd_seg_elet&", @dt_periodicidade_elet="&dt_periodicidade_elet&""
	dbconn.execute(xsql)
	response.write("patrimonio edita OK")
	response.redirect("../../patrimonio.asp?tipo=cadastro&mensagem=Alterado com sucesso")
	
			ElseIf acao = "descartar" AND cd_tipo = "patrimonio" Then
			'Modifica os dados
			xsql = "up_patrimonio_descarte @cd_pat_codigo='"&cd_pat_codigo&"', @dt_descarte='"&dt_data&"', @nm_motivo='"&nm_motivo&"'"
			dbconn.execute(xsql)
			response.write("Descarte OK")
			response.redirect("../../patrimonio.asp?tipo=descarte&mensagem=Alterado com sucesso")
	
	
	Elseif acao = "apaga" AND cd_tipo = "patrimonio" Then
	'Apaga os dados
	xsql = "up_patrimonio_delete @cd_apaga='"&cd_apaga&"'"
	dbconn.execute(xsql)
	response.write("patrimonio apaga OK")
	response.redirect("../../patrimonio.asp?tipo=cadastro&mensagem=Apagado com sucesso")
	
	
	Else 
	action = "erro"
	End If
END IF

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>

<body> onLoad=window.close()>
<br>
Mensagem = <%=action%><br>
Ação: <%=acao%><br>
Tipo: <%=cd_tipo%><br><br>
Mes/Ano: <%=cd_mes%>/<%=cd_ano%><br>
Apagar: <%=cd_apaga%><br>

<br>
Num fab.: <%=cd_pn%><br>
Serie: <%=cd_ns%><br>
Eqp.: <%=cd_equipamento%><br>
Unidade: <%=cd_unidade%><br>
Marca: <%=cd_marca%><br>
<br>
Descarte: <%=dt_descarte%><br>
Motivo: <%=nm_motivo%><br>
Tipo: <%=nm_tipo%><br>
Especialidade: <%=cd_especialidade%>





</body>
</html>
