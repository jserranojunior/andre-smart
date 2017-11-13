<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
   <!--include file="../includes/inc_area_restrita.asp"-->
   
<%
'********************* Variáveis **************************************
cd_codigo = request("cd_codigo")
cd_cod_horas = request("cd_cod_horas")
cd_cod_folga = request("cd_cod_folga")
cd_folga = request("cd_folga")

cd_status = request("cd_status")
cd_unidades = request("cd_unidades")
cd_funcao = request("cd_funcao")

cd_carga_horaria = request("cd_carga_horaria")
cd_excedente = request("cd_excedente")
excedente = request("cd_excedente")
cd_minutos_trab = request("cd_minutos_trab")
excedente_anterior = cd_excedente
acao = request("acao") 

diaentrada = request("diaentrada")
mesentrada = request("mesentrada")
anoentrada = request("anoentrada")

'********************* Verifica ultimo dia do mês *********************
ultimo_dia_mes = UltimoDiaMes(mesentrada,anoentrada)
'**********************************************************************

horaentrada_hr = request("horaentrada_hr")
horaentrada_min = request("horaentrada_min")
data_entrada = mesentrada&"/"&diaentrada&"/"&anoentrada&" "&horaentrada_hr&":"&horaentrada_min
dt_data = mesentrada&"/"&diaentrada&"/"&anoentrada

'************** Intervalo - Saída ********************************
diasaida_int = diaentrada
messaida_int = mesentrada
intsaida_hr = request("intsaida_hr")
intsaida_min = request("intsaida_min")

	if intsaida_hr = "" Then
		intsaida = data_entrada
	elseif intsaida_hr < horaentrada_hr then
		if int(diaentrada) = int(ultimo_dia_mes) Then
			diasaida_int = 1
			messaida_int = mesentrada + 1
			response.write("caso 1 <br>")
		Else
			diasaida_int = diaentrada + 1
			response.write("caso 2 <br>")
		end if
		intsaida = messaida_int&"/"&diasaida_int&"/"&anoentrada&" "&intsaida_hr&":"&intsaida_min
	Else
		intsaida = messaida_int&"/"&diasaida_int&"/"&anoentrada&" "&intsaida_hr&":"&intsaida_min
	end if

'************** Intervalo - Entrada ******************************
diaentrada_int = diasaida_int
mesentrada_int = messaida_int
intentrada_hr = request("intentrada_hr")
intentrada_min = request("intentrada_min")
	
	if intentrada_hr = "" Then
		intentrada = data_entrada
	elseif intentrada_hr < intsaida_hr then
		if int(diaentrada) = int(ultimo_dia_mes) Then
			diaentrada_int = 1
			mesentrada_int = messaida_int + 1
			response.write("caso 1 <br>")
		Else
			diaentrada_int = diasaida_int + 1
			response.write("caso 2 <br>")
		end if
		intentrada = mesentrada_int&"/"&diaentrada_int&"/"&anoentrada&" "&intentrada_hr&":"&intentrada_min
	Else
		intentrada = mesentrada_int&"/"&diaentrada_int&"/"&anoentrada&" "&intentrada_hr&":"&intentrada_min
	end if

diasaida = diaentrada_int
messaida = mesentrada_int
horasaida_hr = request("horasaida_hr")
horasaida_min = request("horasaida_min")

	if horasaida_hr < intentrada_hr then
		if int(diaentrada) = int(ultimo_dia_mes) Then
			diasaida = 1
			messaida = mesentrada_int + 1
		Else
			diasaida = diaentrada_int + 1
		end if
		data_saida = messaida&"/"&diasaida&"/"&anoentrada&" "&horasaida_hr&":"&horasaida_min
	Else
		data_saida = messaida&"/"&diasaida&"/"&anoentrada&" "&horasaida_hr&":"&horasaida_min
	end if
	
cd_carga_horaria = request("cd_carga_horaria")
cd_folga = request("cd_folga")
nm_obs = request("nm_obs")
acao = request("acao")



'************ Se Não houver informações do intervalo inserir valores nulos **************
		'if intsaida_hr = "" or intentrada_hr = "" Then
		'	intsaida = empty
		'	intentrada = empty
		'end if

'************ Se a hora de saída for menor que a hora de entrada, então aumenta um dia ***

'************ Folgas *********************************************************************


'************ Carga Horária **************************************************************
	if acao <> "excluir" Then
	cd_carga_horaria = replace(cd_carga_horaria,":",",")
		cd_carga_horaria = int(cd_carga_horaria) *60
	
'************ Verifica se há carga horaria para o mes de trabalho ************************
	xsql = "SELECT * From TBL_carga_horaria where cd_funcionario = '"&cd_codigo&"' AND cd_mes = '"&mesentrada&"' AND cd_ano = '"&anoentrada&"'"
	Set rs = dbconn.execute(xsql)
	Do while not rs.EOF
	
	carga_h_cod = rs("cd_codigo")
	
	rs.movenext
	Loop
		'*** Se ainda não houver valor para o mes de trabalho, insere o valor digitado no campo ********
		If carga_h_cod = "" Then 'insere
			xsql ="up_carga_horaria_insert @cd_funcionario='"&cd_codigo&"',@dt_mes='"&mesentrada&"', @dt_ano = '"&anoentrada&"', @cd_carga_horaria = '"&cd_carga_horaria&"'"
			dbconn.execute(xsql)
		'*** Se já houver carga horária para o mes de trabalho, altera o valor de horas a trabalhar *****
		Elseif carga_h_cod <> "" Then 'Altera
			xsql ="up_carga_horaria_update @cd_funcionario='"&cd_codigo&"',@cd_mes='"&mesentrada&"', @cd_ano = '"&anoentrada&"', @cd_carga_horaria = '"&cd_carga_horaria&"'"
			dbconn.execute(xsql)
		end if
	end if
%>

<%
if  diaentrada <> "" Or acao = "excluir" Then%>
	
	<!--include file="../banco_horas/banco_horas_acao.asp"-->

<%		If acao = "inserir" AND cd_folga <> "1" Then%>
<!--#include file="../banco_horas/banco_horas_acao.asp"--><%
			xsql ="up_cartao_horas_insert @cd_codigo="&cd_codigo&",@data_entrada='"&data_entrada&"', @data_saida_interv = '"&intsaida&"', @data_entr_interv = '"&intentrada&"', @data_saida='"&data_saida&"',@cd_status="&cd_status&",@cd_unidades='"&cd_unidades&"',@cd_funcao='"&cd_funcao&"',@nm_obs='"&nm_obs&"'"
			dbconn.execute(xsql)
			response.write("OK, inserido")
			
			'response.redirect "../../empresa_funcionarios_horas.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""
			response.redirect "../../empresa_funcionarios_obs.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""
			
		Elseif acao = "editar" AND cd_folga <> "1" Then%>
<!--#include file="../banco_horas/banco_horas_acao.asp"--><%
			xsql ="up_cartao_horas_update @cd_funcionario="&cd_codigo&",@cd_cod_horas="&cd_cod_horas&",@data_entrada='"&data_entrada&"', @data_saida_interv = '"&intsaida&"', @data_entr_interv = '"&intentrada&"',@data_saida='"&data_saida&"',@nm_obs='"&nm_obs&"'"
			dbconn.execute(xsql)
			response.write("OK, alterado")
			'response.redirect "../../empresa_funcionarios_horas.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""
			response.redirect "../../empresa_funcionarios_obs.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""
			
		ElseIf acao = "excluir" AND cd_folga <> "1" Then%>
<!--#include file="../banco_horas/banco_horas_acao.asp"--><%
			xsql ="up_cartao_horas_excluir @cd_codigo='"&cd_codigo&"',@cd_cod_horas='"&cd_cod_horas&"'"
			dbconn.execute xsql
			response.write("OK, excluído")
			'response.redirect "../../empresa_funcionarios_horas.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""
			response.redirect "../../empresa_funcionarios_obs.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""

	'********** F O L G A S *************
		
	Elseif acao = "inserir" AND cd_folga = 1 Then
		xsql ="up_Hora_Trabalhada_folga_insert @cd_funcionario='"&cd_codigo&"', @dt_data='"&dt_data&"', @cd_folga='"&cd_folga&"', @txt_obs='"&nm_obs&"'"
		dbconn.execute xsql
		response.write("Grava a Folga")
		response.redirect "../../empresa_funcionarios_horas.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""
		
	Elseif acao = "excluir" AND cd_folga = "1" Then
		xsql ="up_Hora_Trabalhada_folga_delete @cd_cod_folga='"&cd_cod_folga&"', @cd_funcionario='"&cd_funcionario&"'"
		dbconn.execute xsql
		response.write("Exclui a folga de Cód.: "&cd_cod_folga&" <br>")
		response.redirect "../../empresa_funcionarios_horas.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""

	Else
		response.write("Erro, não registrou")
	End if
	
Else 'Retorna para o formulário sem inserir horas, mas modifica a carga horária do mês
	
	response.redirect "../../empresa_funcionarios_horas.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""
end if
%>

<br>
código: <%=cd_codigo%><br>
cd_status: <%=cd_status%><br>
cd_unidades: <%=cd_unidades%><br>
cd_funcao: <%=cd_funcao%><br>
acao: <%=acao%><br><br>
Mensagem = <%=mens1%><br><br>

entrada --- : <%=data_entrada%><br>
intervalo - s: <%=intsaida%><br>
intervalo - e: <%=intentrada%><br>
saida ------ :<%=data_saida%><br><br>

********* Banco de horas *************************<br>
Excedente (X): <%=excedente%><br>
Minutos trabalhados (Y): <%=cd_minutos_trab%><br>
Novo Excedente (Z): <%=cd_excedente%><br><br>

Excedente mensagem: <%=exc_mens%><br><br>




cod_hora: <%=cd_cod_horas%><br>
qtd_minutos: <%=qtd_minutos%> (<%=qtd_minutos/60%> hs)<br><br>

existem horas no BD? <%=cd_existe%><br><br>

Excedente total anterior: <%=excedente_anterior%><br><br>


Diferença atualização: <%=diferenca%><br>
Horas trabalhadas atualizadas: <%=min%><br><br>

Excedente atualizado : <%=novo_excedente%><br>

Obs: <%=nm_obs%><br>
Obs2: <%=txt_obs%><br>
<br>
Folga: <%=cd_folga%><br>
Cod: <%=cd_cod_folga%><br>

ultimo_dia_mes = <%=ultimo_dia_mes%>


