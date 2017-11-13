<!--#include file="../includes/util.asp"-->
<%
nome_arquivo = "selecao_periodo_2.asp"
nome_arquivo = request("arquivo")
dt_inicio_sys = "01/03/2008"

'**** Data Selecionada ****
dia_sel = zero(request("dia_sel"))
mes_sel = zero(request("mes_sel"))
ano_sel = request("ano_sel")
data_atual = dia_sel&"/"&mes_sel&"/"&ano_sel

'**** Data atual ****
dia_prod = zero(day(now))
mes_prod = zero(month(now))
ano_prod = year(now)
data_prod = dia_prod&"/"&mes_prod&"/"&ano_prod

'if year(dt_inicio_sys) = int(ano_prod) then
'	'ano_sel = year(dt_inicio_sys)
'end if

'*********************************************************
'*				      1ª Parte	 					  	'*
'*********************************************************
'* 1.1 - Seleção do ANO *
'************************%>
<table border="1">
<%if mes_sel = "" AND ano_sel = "" AND cd_codigo = "" Then%>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr><td align=center colspan="100%">&nbsp;</td></tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
					    <td colspan="100%" align="center">Selecione o ano abaixo:</td>
					</tr>
					<form action="<%=nome_arquivo%>" name="Ano" method="get">
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">
					<tr>
						<td align="center" colspan="100%"><br>
						<select name="ano_sel">
								<%ano_inicio = year(dt_inicio_sys)
								response.write(ano_inicio)
								do while not ano_inicio = ano_prod + 1%>
								<option value="<%=ano_inicio%>" <%if ano_inicio=ano_prod then%><%response.write("selected")%><%end if%>><%=ano_inicio%></option>
								<%ano_inicio = ano_inicio + 1
								loop%>			
						</select><br>
						<br>
						<input type="submit" value="OK">						
						</td>
					</tr>
					</form>
<%'************************
'* 1.2 - Seleção do MÊS *
'************************%>
<%elseif mes_sel = "" AND ano_sel <> "" AND cd_codigo = "" Then%>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr><td align=center colspan="100%">&nbsp;</td></tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
					    <td colspan="100%" align="center">Selecione o mes abaixo:</td>
					</tr>
					<form action="<%=nome_arquivo%>" name="Ano" method="get">
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">
					<input type="hidden" name="ano_sel" value="<%=ano_sel%>">
					<tr>
						<td align="center" colspan="100%"><br>
						<select name="mes_sel">
								<%'mes_inicio = month(dt_inicio_sys)
								'response.write(mes_inicio)
								mes = 1
								do while not mes = 12%>
								<option value="<%=mes%>" <%'if ano_inicio=ano_prod then%><%'response.write("selected")%><%'end if%>><%=mes_selecionado(mes)%></option>
								<%'ano_inicio = ano_inicio + 1
								mes = mes + 1
								loop%>			
						</select><br>
						<br>
						<input type="submit" value="OK">						
						</td>
					</tr>
					</form>


<%'***********************
'* 1.2 - Corpo da página *
'*************************
'elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo = "" Then%>
<!--tr><td><%'=data_atual%></td></tr-->
<%end if%>
<tr><td><%=data_atual%></td></tr>

</table>