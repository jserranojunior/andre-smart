<%
'if cd_excedente > 0 Then ' Verifica se existe excedente
'	response.write("Há excedente a incluir <br>")

if acao <> "excluir" Then
'*********************** Calcula quantidade de minutos sendo inserido *******************************
		if hour(intsaida)< hour(data_entrada) Then ' **** intervalo após a meia noite ***
			min1 = datediff("n",cdate(data_entrada),cdate(mesentrada&"/"&diaentrada&"/"&anoentrada&" "&"23:59"))
			min2 = datediff("n",cdate(mesentrada&"/"&diaentrada +1 &"/"&anoentrada&" 0:0"),cdate(intsaida))
			min3 = datediff("n",cdate(intentrada),cdate(data_saida))
			min = 1 + min1 + min2 + min3
				mens1 = "(01) "
		Elseif hour(intentrada)< hour(intsaida) Then ' ***** Intervalo em dias diferentes
			min0 = datediff("n",cdate(data_entrada),cdate(intsaida))
			min1 = datediff("n",cdate(intentrada),cdate(data_saida))
			min = min0 + min1
				mens1 = "(02) "
		Elseif hour(data_saida)< hour(data_entrada) Then ' ***** Saída no outro dia
			min0 = datediff("n",cdate(data_entrada),cdate(intsaida))
			min1 = datediff("n",cdate(intentrada),cdate(day(intentrada)&"/"&month(intentrada)&"/"&year(intentrada)&" 23:59"))
			response.write(day(intentrada))
			min2 = datediff("n",cdate(day(data_saida)&"/"&month(data_saida)&"/"&year(data_saida)&" 0:0"),cdate(data_saida))
			min = 1 + min0 + min1 + min2
				mens1 = "(03) "
		Else '***** Processo normal
			min1 = datediff("n",cdate(data_entrada),cdate(intsaida))
			min2 = datediff("n",cdate(intentrada),cdate(data_saida))
			min = min1 + min2
				mens1 = "(00) "
		end if
End if		
		
		'Em caso de inserção no BD
		if acao = "inserir" then 'OK!
			cd_excedente = cd_excedente + min
			
		elseif acao = "editar" Then
			cd_excedente = cd_excedente - cd_minutos_trab
				cd_excedente = cd_excedente + min
			
		elseif acao = "excluir" then
			cd_excedente = cd_excedente - cd_minutos_trab
		end if

'if cd_excedente <> a Then ' Verifica se existe excedente
	response.write("Há excedente a incluir <br>")		

	'***** Verifica se já existe excedente para o FUNCIONÁRIO no MES e no ANO selecionados... *****
				xsql = "select * from TBL_banco_horas Where cd_funcionario='"&cd_codigo&"' AND dt_mes='"&mesentrada&"' AND dt_ano='"&anoentrada&"'"
				set rs = dbconn.execute(xsql) 
				
				do while NOT rs.EOF
				 'response.write ("existe!")
				 	cd_existe = rs("cd_codigo")
				rs.movenext
				loop
		
	'***** Operações no banco de dados... *****
		if cd_existe = "" Then 'Se não houver registros, insere!...
			response.write("Ainda inexistente, inserir "&min&" minutos, <br>")
			exc_mens = "Ainda inexistente, inserir "&min&" minutos"
				xsql = "up_banco_horas_insert @cd_funcionario='"&cd_codigo&"', @cd_horas_excedentes='"&cd_excedente&"', @dt_mes='"&mesentrada&"', @dt_ano='"&anoentrada&"'"
				dbconn.execute(xsql)

		Elseif cd_existe <> "" then 'Se já houver registros, atualiza!...
			response.write("Já existe, incluir "&min&"minutos, <br>")
			exc_mens = "Já existe, incluir "&min&"minutos"
				xsql = "up_banco_horas_update @cd_funcionario='"&cd_codigo&"', @cd_horas_excedentes='"&cd_excedente&"', @cd_existe='"&cd_existe&"'"
				dbconn.execute(xsql)

		Else
			response.write("(Z E R O) <br>")
			exc_mens = "( Z e r o )"
		End if

	if acao = "inserir" then
	qtd_minutos = min
	end if
'end if
'Else
'response.write("Não há excedente a incluir <br>")
'end if
%>