<!--include file="../includes/numero_extenso.asp"-->
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../css/geral.htm"-->

<%dt_ano = request("ano_sel")
dt_mes = request("mes_sel")
	'if dt_mes = "" OR dt_ano = "" then
	if dt_ano = "" then
		dt_mes = month(now)
		dt_ano = year(now)
	end if
		dia_final = ultimodiames(dt_mes,dt_ano)
		
		mes_anterior = dt_mes - 1
		ano_anterior = dt_ano
			if mes_anterior < 1 then
				mes_anterior = 12
				ano_anterior = dt_ano - 1
			end if
			dia_final_anterior = ultimodiames(mes_anterior,ano_anterior)
		
		mes_posterior = dt_mes + 1
		ano_posterior = dt_ano
			if mes_posterior > 12 then
				mes_posterior = 1
				ano_posterior = dt_ano + 1
			end if
			dia_final_posterior = ultimodiames(mes_posterior,ano_posterior)

cd_unidade = request("cd_unidade")
cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)
dt_hoje_numerica = year(now)&zero(month(now))&zero(day(now))%>


<form action="../financeiro/fatura_consolidado3.asp" name="seleciona_periodo" method="post" id="seleciona_período">

		<input type="hidden" name="tipo" value="consolidado">		
	<table align="center" class="no_print" border="0" style="border:2px solid black; border-collapse:collapse;">
		<tr style="background-color:gray; color:white; font-size:14px;">
			<td align="center" colspan="3">Consolidado - <%=dt_ano%></td>
		</tr>
		
		<tr bgcolor="#c0c0c0">
			<td align="center">Ano:&nbsp;<input type="text" name="ano_sel" maxlength="4" size="4" value="<%=dt_ano%>"></td>
		<!--/tr>
		<tr bgcolor="#c0c0c0">
			<td colspan="2">&nbsp;<select name="cd_unidade">
					<option value="0">&nbsp;</option>
					<%xsql = "SELECT * FROM TBL_unidades where cd_status = 1 and cd_hospital = 1"
					Set rs = dbconn.execute(xsql)%>
					<%while not rs.EOF
						strcd_unidade = rs("cd_codigo")
						strnm_unidade = rs("nm_unidade")
						strnm_sigla = rs("nm_sigla")%>
						<option value="<%=strcd_unidade%>" <%if abs(strcd_unidade) = abs(cd_unidade) then response.write("selected")%> ><%=strnm_unidade%></option>
					<%rs.movenext
					wend%>
				</select>
			
			</td>
		</tr>
		<tr align="left" bgcolor="#c0c0c0"-->
			<td colspan="2"><input type="submit" name="OK" value="Ok" width="4" height="5" style="background-color:orange;">&nbsp;&nbsp;&nbsp;</td>
		</tr>
		</form>
	</table>
	<!--br class="no_print"-->

<table border="1" class="nok_print" style="border-collapse:collapse">
	<tr bgcolor="#e9e9e9">
		<%if int(dt_ano) < year(now) then
			abrangencia_mes2 = 12
		else
			abrangencia_mes2 = month(now)+1
				if abrangencia_mes2 > 12 then
					abrangencia_mes2 = 12
				end if
		end if
		total_recebidos_atrasados = 0
		total_atrasados = 0
		
		linha = 0
		for i = 1 to abrangencia_mes2
			dt_mes = i%>
		
			<td align="center" style="font-size:13px; color:red;"><b><%=mes_selecionado(i)%></b>
			
			
			</td>
		</tr>
			<%'if i < abrangencia_mes2 then%><!--td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td--><%'end if%>
		
		<%nr_total_mes = 0
		total_recebidos_atrasados = 0
		nm_valor_recebido_atrasado = 0
		total_atrasados = 0
		nm_valor_atrasado = 0
		next%>
		
</table>