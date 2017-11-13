<!--#include file="inc_open_connection.asp"-->
<!--include file="util.asp"-->

<%'*********************************
'*** 1   - Feriados Nascionais ***
'*********************************
if data_atual = "" Then
	data_atual = request("data_atual")
end if


if data_atual = "" Then
	data_atual = now
end if

if data_atual = "" Then
	ano_sel_feriados = year(now)
else
	ano_sel_feriados = year(data_atual)
end if

'*** A - Cálculo Carnaval, Paixão, Páascoa e Corpus Cristi ***
		'cód. válido de 1900 até 2099 - http://www.inf.ufrgs.br/~cabral/Pascoa.html
		x = 24
		y = 5
		
		a = year(data_atual) MOD 19
		b = year(data_atual) MOD 4
		c = year(data_atual) MOD 7
		d = (19 * a + X) MOD 30
		e = (2 * b + 4 * c + 6 * d + Y) MOD 7
		
		if (d + e) > 9 then
			dia_pascoa = d + e - 9
			mes_pascoa = 4
		elseif int(d + e) = 9 then
			dia_pascoa = 31
			mes_pascoa = 3
		else
			dia_pascoa = d + e + 9
			mes_pascoa = 3
		end if
		carnaval = dateadd("d",-47,dia_pascoa&"/"&mes_pascoa&"/"&year(data_atual))
		paixao = dateadd("d",-2,dia_pascoa&"/"&mes_pascoa&"/"&year(data_atual))
		pascoa = dia_pascoa&"/"&mes_pascoa&"/"&year(data_atual)
		c_cristi = dateadd("d",60,dia_pascoa&"/"&mes_pascoa&"/"&year(data_atual))
		
				if weekday(pascoa) = 1 AND month(data_atual) = month(pascoa) AND year(data_atual) = year(pascoa) then qtd_feriado_dsr = qtd_feriado_dsr + 1
				if weekday(c_cristi) = 1 AND month(data_atual) = month(c_cristi) AND year(data_atual) = year(c_cristi) then qtd_feriado_dsr = qtd_feriado_dsr + 1
		
'*** B - Outros Feriados Nacionais ***
		'xsql ="SELECT * FROM TBL_calendario_feriados where day(dt_data)="&i&" and month(dt_data) = "&month(data_atual)&""
		xsql ="SELECT * FROM TBL_calendario_feriados where month(dt_data) = "&month(data_atual)&""
			Set rs_feriado = dbconn.execute(xsql)
			do while Not rs_feriado.EOF
				feriados = rs_feriado("titulo_dia")
				feriado_data = rs_feriado("dt_data")
					if weekday(month(feriado_data)&"/"&day(feriado_data)&"/"&year(ano_sel_feriados)) = 2 then
						qtd_feriado_dsr = qtd_feriado_dsr + 1
					end if
				
				'response.write(feriados&": "&feriado_data&"<br>")
				qtd_feriados = qtd_feriados + 1
			rs_feriado.movenext
			loop
%>

<%'response.write("carnaval: "&carnaval&"<br>Paixão: "&paixao&"<br> Pascoa: "&pascoa&"<br>Corpus Cristi: "&c_cristi&"<br><br><br>")

if month(carnaval) = month(data_atual) then	
	'response.write("Carnaval: "&carnaval&"<br>")
	qtd_feriados = qtd_feriados + 1
end if
if month(paixao) = month(data_atual) then	
	'response.write("Paixão: "&paixao&"<br>")
	qtd_feriados = qtd_feriados + 1
end if
if month(pascoa) = month(data_atual) then	
	'response.write("Páscoa: "&pascoa&"<br>")
	qtd_feriados = qtd_feriados + 1
end if
if month(c_cristi) = month(data_atual) then	
	'response.write("C.Cristi: "&c_cristi&"<br>")
	qtd_feriados = qtd_feriados + 1
end if

if qtd_feriado_dsr = "" Then
	qtd_feriado_dsr = 0
end if
'response.write(month(carnaval)&"<br>"&month(data_atual))
'response.write("-------------------------------------------<br>total: "&qtd_feriados)
%>