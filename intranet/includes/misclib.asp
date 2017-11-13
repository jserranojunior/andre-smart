<%

function iniciodasemana(dia,mes,ano)
	data_=DateSerial(ano,mes,dia)
	diadasemana=weekday(data_)
	if diadasemana<>1 then data_=dateadd("d",-diadasemana+1,data_)
	iniciodasemana=data_
end function

function zeroesquerda(i)  '  poe o zero nao significativo a esquerda de um numero, para uso em data e hora e numeros com dois digitos
	'response.write "zero esquerda->" & i & "->0" & i & "<BR>" ' debug
	if len(i) = 1 then zeroesquerda = "0" & i else zeroesquerda=i
end function

function ifchecked(i)
	if i="true" or i="on" or i="1" then ifchecked=1 else ifchecked=0
end function

function datamysql(dia,mes,ano)
	dia=zeroesquerda(dia)
	mes=zeroesquerda(mes)
	'ano=zeroesquerda(ano)
	'response.write "datamysql->" & "dia=" & dia & " mes=" & mes & " ano=" & ano & "<BR>" ' debug
	datamysql=ano &"-"& mes & "-" & dia
end function

function horamysql(hora,minuto,segundo)
	hora=zeroesquerda(hora)
	minito=zeroesquerda(minuto)
	segundo=zeroesquerda(segundo)
	horamysql=hora &":"& minuto & ":" & segundo
end function

function checkedif(i)
	if i="true" or i="on" or i="1" or i=1 or i=true then response.write "checked" else response.write ""
end function

function databrasil(dia,mes,ano)
	databrasil=dia & "/" &mes & "/" &ano
end function
%>