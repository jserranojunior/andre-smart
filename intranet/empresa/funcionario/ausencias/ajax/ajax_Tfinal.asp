<!--#include file="../../../../includes/inc_open_connection.asp"-->
<!--#include file="../../../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

qtd_dias = request("qtd_dias")
dt_inicial = request("dt_inicial")
dt_final = request("dt_final")

if qtd_dias = "" Then
	qtd_dias = 0
end if
		
if isDate(dt_inicial) AND  dt_final="" Then
	if qtd_dias > 0 then
		diferenca = Dateadd("d",qtd_dias,dt_inicial)
			diferenca = Dateadd("d",-1,diferenca)
				per_dia_f = day(diferenca)
				per_mes_f = month(diferenca)
				per_ano_f = year(diferenca)
	end if%>	
		<input type="text" name="ajax_dia_f" id="ajax_dia_f" size="2" maxlength="2" value="<%=zero(per_dia_f)%>"  onKeyup="calcula_dias(document.getElementById('dt_dia').value, document.getElementById('dt_mes').value, document.getElementById('dt_ano').value, this.value, document.getElementById('per_mes_f').value, document.getElementById('per_ano_f').value);">
		<input type="text" name="ajax_mes_f" id="ajax_mes_f" size="2" maxlength="2" value="<%=zero(per_mes_f)%>"  onKeyup="calcula_dias(document.getElementById('dt_dia').value, document.getElementById('dt_mes').value, document.getElementById('dt_ano').value, this.value, document.getElementById('per_mes_f').value, document.getElementById('per_ano_f').value);">
		<input type="text" name="ajax_ano_f" id="ajax_ano_f" size="4" maxlength="4" value="<%=per_ano_f%>" 		  onKeyup="calcula_dias(document.getElementById('dt_dia').value, document.getElementById('dt_mes').value, document.getElementById('dt_ano').value, this.value, document.getElementById('per_mes_f').value, document.getElementById('per_ano_f').value);">
	<img src="../../../../imagens/px.gif" alt="" width="1" height="1" border="0" onLoad="data_final('<%=per_dia_f%>','<%=per_mes_f%>','<%=per_ano_f%>','<%=qtd_dias%>');">
<%
elseif isDate(dt_inicial) AND isDate(dt_final) Then
	diferenca = Datediff("d",dt_inicial&" 00:00",dt_final&" 23:59")
		'diferenca = Dateadd("d",-1,diferenca)
			per_dia_f = day(dt_final)
			per_mes_f = month(dt_final)
			per_ano_f = year(dt_final)
			qtd_dias = diferenca+1%>
	<input type="text" name="qtd_dias_x" size="3" maxlength="4" value="<%=qtd_dias%>" onKeyup="calcula_data(this.value, document.getElementById('dt_dia').value, document.getElementById('dt_mes').value, document.getElementById('dt_ano').value);"> dia(s)
	<img src="../../../../imagens/px.gif" alt="" width="5" height="1" border="0" onLoad="data_final2('<%=qtd_dias%>');">
<%else%>

<%end if%>	
<%'rs_cid.close
'set rs_cid = nothing%>
<br>
<%'=data_inicial%>
