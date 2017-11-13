<%
tipo = request("tipo")
'id  = request("cd_codigo")

dt_dia = request("dia_sel")
dt_mes = request("mes_sel")
dt_ano = request("ano_sel")
if dt_dia = "" or dt_mes = "" or dt_ano = "" then
	dt_dia = ""'request("dia_sel")
	dt_mes = ""'request("mes_sel")
	dt_ano = ""'request("ano_sel"
end if

'Para calculo de Pascoa, carnava e corpus cristi, acesse o site abaixo:
'http://www.inf.ufrgs.br/~cabral/Pascoa.html

data_selecionada = dt_dia&"/"&dt_mes&"/"&dt_ano
'*** Verifica se há alguma data selecionada... ***
if data_selecionada = "//" Then
	data_atual = day(now)&"/"&month(now)&"/"&year(now) '...se não houver, pega a data de agora 
Else
	data_atual = data_selecionada '...se houver, pega a data selecionada
end if

hoje = day(data_atual)
ultimo_dia = UltimoDiaMes(month(data_atual),year(data_atual))
%>
<!--#include file="../js/ajax.js"-->
<!--#include file="../calendario/js/calendario.js"-->
<style type="text/css">
	@import "css/calendario.css";
</style>
<script type="text/javascript" language="javascript">
<!--	
	function fn_calendario(){
		h_data = new Date();
			h_dia = document.calendario.dt_dia.value;
			h_mes = document.calendario.dt_mes.value;
			h_ano = document.calendario.dt_ano.value;
				
				h_data = h_dia + "/" + h_mes + "/" + h_ano;
				document.calendario.variavel.value=h_mes + "/" + h_ano;					
	}
	
	function fn_calendario_hoje(){
		h_data = new Date();
			h_dia = h_data.getDate();
			h_mes = h_data.getMonth()+1;
			h_ano = h_data.getFullYear();
			
				h_data = h_dia + "/" + h_mes + "/" + h_ano;
				document.calendario.dt_dia.value=h_dia;
				document.calendario.dt_mes.value=h_mes;
				document.calendario.dt_ano.value=h_ano;
				document.calendario.variavel.value=h_mes + "/" + h_ano;
			{
		     var oHTTPRequest_cal = createXMLHTTP(); 
		     oHTTPRequest_cal.open("post", "calendario/ajax/ajax_calendario.asp", true);
		     oHTTPRequest_cal.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		     oHTTPRequest_cal.onreadystatechange=function(){
		      if (oHTTPRequest_cal.readyState==4){
		         document.all.divCalendario.innerHTML = oHTTPRequest_cal.responseText;}}
		       oHTTPRequest_cal.send("variavel=" + calendario.variavel.value+";"+calendario.tipo.value);
 			}
	}
	
	function fn_calendario_menos(){		
		h_mes = parseInt(h_mes) - 1;
			if (h_mes < 1){
				h_ano = parseInt(h_ano) - 1;
				h_mes = 12;
				}
				document.calendario.dt_mes.value=h_mes;
				document.calendario.dt_ano.value=h_ano;
				document.calendario.variavel.value=h_dia + "/" + h_mes + "/" + h_ano;				
			{
		     var oHTTPRequest_cal = createXMLHTTP(); 
		     oHTTPRequest_cal.open("post", "calendario/ajax/ajax_calendario.asp", true);
		     oHTTPRequest_cal.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		     oHTTPRequest_cal.onreadystatechange=function(){
		      if (oHTTPRequest_cal.readyState==4){
		         document.all.divCalendario.innerHTML = oHTTPRequest_cal.responseText;}}
		       oHTTPRequest_cal.send("variavel=" + calendario.variavel.value+";"+calendario.tipo.value);
 			}	
		}
	
	function fn_calendario_mais(){	
		h_mes = parseInt(h_mes) + 1;
			if (h_mes > 12){
				h_ano = parseInt(h_ano) + 1;
				h_mes = 1;
				}
				document.calendario.dt_mes.value=h_mes;
				document.calendario.dt_ano.value=h_ano;			
				document.calendario.variavel.value=h_dia + "/" + h_mes + "/" + h_ano;			
			{
		     var oHTTPRequest_cal = createXMLHTTP(); 
		     oHTTPRequest_cal.open("post", "calendario/ajax/ajax_calendario.asp", true);
		     oHTTPRequest_cal.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		     oHTTPRequest_cal.onreadystatechange=function(){
		      if (oHTTPRequest_cal.readyState==4){
		         document.all.divCalendario.innerHTML = oHTTPRequest_cal.responseText;}}
		       oHTTPRequest_cal.send("variavel=" + calendario.variavel.value+";"+calendario.tipo.value);
 			}
	}
document.onload=fn_calendario;
-->
</script>
	<form action="" name="calendario" value="" >
		<input type="hidden" name="dt_dia" size="2" value="<%=day(data_atual)%>">
		<input type="hidden" name="dt_mes" size="2" value="<%=month(data_atual)%>">
		<input type="hidden" name="dt_ano" size="4" value="<%=year(data_atual)%>">
		<input type="hidden" name="variavel" size="20" value="<%=month(data_atual)&"/"&year(data_atual)%>">
		<input type="hidden" name="tipo" size="20" value="<%=tipo%>">		
	</form>
<!--div style="border:1px solid silver; width:150px; height:180px; position:relative; top:325px; left:12px; z-index:1; background:white;" class="no_print"-->
	<table border="1" width="150">
		<tr>
			<td colspan="7" align="center"><b><a href="javascript:void(0);" onclick="fn_calendario_menos();" title="Mês anterior"><<</a>
				<a href="javascript:void(0);" onclick="fn_calendario_hoje();" title="Retorna ao mês atual"><b>CALENDÁRIO</b></a>
				<a href="javascript:void(0);" onclick="fn_calendario_mais();" title="Próximo mês">>></a></b></td>
		</tr>		
	</table>
	<div id="divCalendario">
		<table border="0" width="150" style="font-size:8px;">
			<tr>
				<td colspan="7" align="center" style="color:white; background-color:black;">
					<a href="#"  onclick="javascript:window.open('empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=cd_funcionario%>&busca=<%=strbusca%>','ficha<%=cd_funcionario%>','width=70,height=200,scrollbars=1');" onmouseover="JavaScript:this.style.cursor='pointer'" style="color:white;">
						<b><%response.write(ucase(mes_selecionado(month(data_atual)))&" - "&year(data_atual))%></b>
					</a>
				</td>
			</tr>
			<tr>
				<%'*******************************
				'*  DEFINE A POSIÇÃO DO DOMINGO  *
				'* 	1 = primeiro dia da semana;  *
				'* 	0 = ultimo dia da semana.    *
				'*********************************
					domingo = 1
				'*********************************
				dias_semana = "border:1px solid black; height:10px; width:7%;"%>
				<%if domingo = 1 then%>
					<td align="center" style="<%=dias_semana%>"><b>D</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>T</b></td><td align="center" style="<%=dias_semana%>"><b>Q</b></td><td align="center" style="<%=dias_semana%>"><b>Q</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td>
				<%else%>
					<td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>T</b></td><td align="center" style="<%=dias_semana%>"><b>Q</b></td><td align="center" style="<%=dias_semana%>"><b>Q</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>S</b></td><td align="center" style="<%=dias_semana%>"><b>D</b></td>
				<%end if%>
			</tr>
			<tr>			
				<%'*** tipo de Calendário : DSTQQSS=0 / STQQSSD = "-1"***
				dias = 1
				if domingo = 1 then
					dom = 2
					sab = 1
					ajusta_dia = weekday("1/"&month(data_atual)&"/"&year(data_atual)) '*** MODELO DSTQQSS ***
				else
					dom = 1
					sab = 7
					ajusta_dia = weekday("1/"&month(data_atual)&"/"&year(data_atual))-1 '* MODELO STQQSSD ***
				end if
					if ajusta_dia < 1 then
						ajusta_dia = 7
					end if
					'Preenche as lacunas vazias no inicio
					while not dias = ajusta_dia %>						
						<td style="background-color:#eaeaea;">&nbsp;</td>
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
								if hoje = year(data_atual)&zero(month(data_atual))&zero(i) then
									cor_dia = "white"
								else
									cor_dia = "black"
								end if
							
							else
								bg_cor = "silver"
								if hoje = year(data_atual)&zero(month(data_atual))&zero(i) then
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
										
											if zero(month(data_atual))&zero(i) = zero(month(carnaval))&zero(day(carnaval)) then
												titulo_dia = "Carnaval"
												borda_cor = "1px solid #000000"
												cor_dia = "white"
												bg_cor = "#cc0000"
											elseif  zero(month(data_atual))&zero(i) = zero(month(paixao))&zero(day(paixao)) then
												titulo_dia = "Paixão"
												borda_cor = "1px solid #000000"
												cor_dia = "white"
												bg_cor = "#cc0000"
											elseif  zero(month(data_atual))&zero(i) = zero(month(pascoa))&zero(day(pascoa)) then
												titulo_dia = "Páscoa"
												borda_cor = "1px solid #000000"
												cor_dia = "white"
												bg_cor = "#cc0000"
											elseif  zero(month(data_atual))&zero(i) = zero(month(c_cristi))&zero(day(c_cristi)) then
												titulo_dia = "Corpus Cristi"
												borda_cor = "1px solid #000000"
												cor_dia = "white"
												bg_cor = "#cc0000"
											end if
								'***************************************
								'*** 1.B - Outros Feriados Nacionais ***
								'***************************************
								xsql ="SELECT * FROM TBL_calendario_feriados where day(dt_data)="&i&" and month(dt_data) = "&month(data_atual)&""
									Set rs_feriado = dbconn.execute(xsql)
									do while Not rs_feriado.EOF
										feriados = rs_feriado("titulo_dia")
										feriado_data = rs_feriado("dt_data")
									rs_feriado.movenext
									loop
										
										if zero(month(data_atual))&zero(i) = zero(month(feriado_data))&zero(day(feriado_data)) then
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
								 
					<td align="center" style="border:<%=borda_cor%>; background-color:<%=bg_cor%>; height:20px; width:1%;"><a href="<%=pagina_atual%>?tipo=<%=tipo%>&dia_sel=<%=i%>&mes_sel=<%=month(data_atual)%>&ano_sel=<%=year(data_atual)%>&id=<%=id%>" title="<%=titulo_dia%>"><b><%=zero(i)%></b></a><%'=pascoa%></td><%=separador%>
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
		</div>
<!--/div-->
