<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%
variavel = request("variavel")

'*** Separa a data e a instrução do objetivo ***
dt_calendario = mid(request("variavel"),1,instr(variavel,";")-1)
tipo = mid(variavel,instr(variavel,";")+1,len(variavel))
'***********************************************
	pri_bar = instr(1,dt_calendario,"/",1)
	seg_bar = instr(pri_bar+1,dt_calendario,"/",1)
	
	if seg_bar > 0 Then
		dt_cal_dia = mid(dt_calendario,1,pri_bar-1)
		dt_cal_mes = mid(dt_calendario,pri_bar+1,(seg_bar-pri_bar)-1)
		dt_cal_ano = mid(dt_calendario,seg_bar+1,4)
	else
		'dt_cal_dia = mid(dt_calendario,1,pri_bar-1)
		dt_cal_mes = mid(dt_calendario,1,pri_bar-1)
		dt_cal_ano = mid(dt_calendario,pri_bar+1,4)
	end if
	'response.write(dt_calendario)


dt_dia = day(now)

'dt_data = dt_dia&"/"&dt_calendario
'dt_data = dt_calendario
ultimo_dia = UltimoDiaMes(dt_cal_mes,dt_cal_ano)
dt_data = dt_cal_mes&"/"&ultimo_dia&"/"&dt_cal_ano



'ultimo_dia = UltimoDiaMes(mes_data,ano_data)
'ultimo_dia = UltimoDiaMes(dt_cal_mes,dt_cal_ano)

dt_dia = day(dt_data)
%>
		<table border="0" width="150" style="font-size:8px;">
			<tr>
				<td colspan="7" align="center" style="color:white; background-color:black;">
					<a href="javascript:void(0);" onclick="fn_calendario_hoje();" title="Retorna ao mês atual" style="color:white;">
						<b><%response.write(ucase(mes_selecionado(month(dt_data)))&" - "&year(dt_data))%></b>
					</a>
				</td>
			</tr>
			<tr>
				<%'********************************
				'*  DEFINE A POSICÇÃO DO DOMINGO  *
				'* 	1 = primeiro dia da semana;   *
				'* 	0 = ultimo dia da semana.     *
				'**********************************
					domingo = 1
				'**********************************
				dias_semana = "border:1px solid black; height:10px; width:7%;"%>
				<%if domingo = 1 then%>
					<td align="center" style="<%=dias_semana%>"><b>D</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>T</b></td><td align="center" style="<%=dias_semana%>"><b>Q</b></td><td align="center" style="<%=dias_semana%>"><b>Q</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td>
				<%else%>
					<td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>T</b></td><td align="center" style="<%=dias_semana%>"><b>Q</b></td><td align="center" style="<%=dias_semana%>"><b>Q</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>D</b></td>
				<%end if%>
			</tr>
			<tr>			
				<%dias = 1
				'*** tipo de Calendário : DSTQQSS=0 / STQQSSD = "-1"***
				if domingo = 1 then
					dom = 2
					sab = 1
					ajusta_dia = weekday("1/"&month(dt_data)&"/"&year(dt_data)) '*** MODELO DSTQQSS ***
				else
					dom = 1
					sab = 7
					ajusta_dia = weekday("1/"&month(dt_data)&"/"&year(dt_data))-1 '* MODELO STQQSSD ***
				end if
					if ajusta_dia < 1 then
						ajusta_dia = 7
					end if
					
					
					'Preenche as lacunas vazias no inicio
					while not dias = ajusta_dia %>						
							<td style="background-color:#eaeaea;;">&nbsp;</td>
						<%dias = dias + 1%>
					<%wend%>
			
					<%'-----------------------------------------------------------------------------------
					for i=1 to ultimo_dia
						
						if int(dias) = 7 then
							separador = "</tr><tr>"'"<br>"
							dias = int(1)
						else
							separador = ""
							dias = dias + 1
						end if
						
							hoje = year(now)&zero(month(now))&zero(day(now))
							if dias = sab or dias = dom then
								bg_cor = "red"
								if hoje = year(dt_data)&zero(month(dt_data))&zero(i) then
									cor_dia = "white"
								else
									cor_dia = "black"
								end if
							else
								bg_cor = "silver"
								if hoje = year(dt_data)&zero(month(dt_data))&zero(i) then
									cor_dia = "white"
									titulo_dia = "Hoje"
									bg_cor = "green"'"#ff6666"
									'borda_cor = "#000000"
								else
									cor_dia = "black"
									titulo_dia = ""
								end if
							end if
								
								borda_cor = "gray"
								
								
								'******************************************
								'*** 1   - 	Feriados Nacionais 			***
								'******************************************
								'*** 1.A - Calculo carnaval, pascoa etc ***
										'cód. válido de 1900 até 2099 - http://www.inf.ufrgs.br/~cabral/Pascoa.html
										x = 24
										y = 5
										
										a = year(dt_data) MOD 19
										b = year(dt_data) MOD 4
										c = year(dt_data) MOD 7
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
										carnaval = dateadd("d",-47,dia_pascoa&"/"&mes_pascoa&"/"&year(dt_data))
										paixao = dateadd("d",-2,dia_pascoa&"/"&mes_pascoa&"/"&year(dt_data))
										pascoa = dia_pascoa&"/"&mes_pascoa&"/"&year(dt_data)
										c_cristi = dateadd("d",60,dia_pascoa&"/"&mes_pascoa&"/"&year(dt_data))
										
											if zero(month(dt_data))&zero(i) = zero(month(carnaval))&zero(day(carnaval)) then
												titulo_dia = "Carnaval"
												borda_cor = "1px solid #000000"
												cor_dia = "white"
												bg_cor = "#cc0000"
											elseif  zero(month(dt_data))&zero(i) = zero(month(paixao))&zero(day(paixao)) then
												titulo_dia = "Paixão"
												borda_cor = "1px solid #000000"
												cor_dia = "white"
												bg_cor = "#cc0000"
											elseif  zero(month(dt_data))&zero(i) = zero(month(pascoa))&zero(day(pascoa)) then
												titulo_dia = "Páscoa"
												borda_cor = "1px solid #000000"
												cor_dia = "white"
												bg_cor = "#cc0000"
											elseif  zero(month(dt_data))&zero(i) = zero(month(c_cristi))&zero(day(c_cristi)) then
												titulo_dia = "Corpus Cristi"
												borda_cor = "1px solid #000000"
												cor_dia = "white"
												bg_cor = "#cc0000"
											end if
								'***************************************
								'*** 1.B - Outros Feriados Nacionais ***
								'***************************************
								xsql ="SELECT * FROM TBL_calendario_feriados where day(dt_data)="&i&" and month(dt_data) = "&month(dt_data)&""
									Set rs_feriado = dbconn.execute(xsql)
									do while Not rs_feriado.EOF
										feriados = rs_feriado("titulo_dia")
										feriado_data = rs_feriado("dt_data")
									rs_feriado.movenext
									loop
										
										if zero(month(dt_data))&zero(i) = zero(month(feriado_data))&zero(day(feriado_data)) then
											'bg_cor = "#ff6666"
											if titulo_dia <> "" Then
												titulo_dia = titulo_dia&vbcrlf&feriados
											else
												titulo_dia = feriados
											end if
											borda_cor = "2px solid #000000"
											cor_dia = "white"
											bg_cor = "#cc0000"
										'else
											'borda_cor = "3px solid #000000"
											'cor_dia = "white"
											'bg_cor = "#cc0000"
										end if%>
								 
					<td align="center" style="border:<%=borda_cor%>; background-color:<%=bg_cor%>; height:20px; width:7%;"><a href="<%=pagina_atual%>?tipo=<%=tipo%>&dia_sel=<%=i%>&mes_sel=<%=month(dt_data)%>&ano_sel=<%=year(dt_data)%>&id=<%=id%>" title="<%=titulo_dia%>"><b><%=zero(i)%></b></a><%'=pascoa%></td><%=separador%>
					<%titulo_dia = ""
					next%>
					<%'-----------------------------------------------------------------------------------
					'Preenche as lacunas vazias no final
					dias = dias - 1 
					if not dias = 0 then
						while not dias = 7 %>
							<td style="background-color:#eaeaea;">&nbsp;</td>
							<!--b id="esp_d_calend"><%dias = dias + 1%>&nbsp;</b-->
						<%wend
					end if%>
				</table>