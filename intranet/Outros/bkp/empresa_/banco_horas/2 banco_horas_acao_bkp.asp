<%
'Vari�veis
cd_excedente = request("cd_excedente")




'Verifica se h� excedente a incluir no Banco de Horas...
	'**************** Caso a a��o seja diferente de "Excluir", verifica a quantidade de horas sendo inseridas **********************
	if acao <> "excluir" Then
					'Se saida ou entrada do intervalo for igual a nada, ent�o calcular a diferen�a entre a hora de sa�da e a hora de entrada.
						if intsaida_hr = "" OR intentrada_hr = "" Then
							min = datediff("n",cdate(data_entrada),cdate(data_saida))
							'response.write("1")
						'Ou ent�o, calcula a diferen�a entre a hora de entrada e a hora de sa�da para o intervalo, e a hora de entrada do intervalo e a hora de sa�da.
						Else
							min1 = datediff("n",cdate(data_entrada),cdate(data_sai))
							min2 = datediff("n",cdate(data_entr),cdate(data_saida))
								min = min1 + min2
							'response.write("2")
						end if
	
end if

if cd_excedente > 0 Then
	response.write("H� excedente a incluir "&"<br>")


		' Verifica se j� existe excedente para o FUNCION�RIO no MES e no ANO selecionados...
				xsql = "select * from TBL_banco_horas Where cd_funcionario='"&cd_codigo&"' AND dt_mes='"&mesentrada&"' AND dt_ano='"&anoentrada&"'"
				set rs = dbconn.execute(xsql) 
				
				do while NOT rs.EOF
				 'caso n�o exista, insere o valor excedente no banco de dados...
				 'response.write ("existe!")
				 	cd_existe = rs("cd_codigo")
				rs.movenext
				loop
				if cd_existe = "" Then
				response.write("Ainda inexistente, inserir "&min&" minutos, ")
				Else 
				response.write("J� existe, incluir "&min&"minutos, ")
				End if


Else 
	response.write("N�o h� excedente a incluir "&"<br>")
End if
%>