<!--#include file="../../../../includes/inc_open_connection.asp"-->
<!--#include file="../../../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_tfinal = request("cd_tfinal")
		qtd_dias = mid(cd_tFinal,1,instr(cd_tfinal,";")-1)
			if qtd_dias = "" Then
				qtd_dias = 0
			end if
		
		data_inicial = mid(cd_tfinal,instr(cd_tfinal,";")+1,len(cd_tfinal))
			dia_inicial = int(day(data_inicial))
			per_mes_f = month(data_inicial)
			per_ano_f = year(data_inicial)
			
			ver_ultimo_dia = UltimoDiaMes(month(data_inicial),year(data_inicial))
				
				for i = 1 to qtd_dias
					
					if per_dia_f = 1 Then 'Quando mudar o mes, verifica qual o ultimo dia novamente
						ver_ultimo_dia = UltimoDiaMes(per_mes_f,year(data_inicial))
					end if
					
					if i = 1 then 'Soma + 1 dia ao dia inicial
						per_dia_f = dia_inicial + i
					else
						per_dia_f = per_dia_f + 1
					end if
					
						if int(per_dia_f) > ver_ultimo_dia then 'Reinicia a contagem de dias, caso seja igual ao ultimo dia do mes e passa ao mês seguinte.
							per_dia_f = 1
							per_mes_f = per_mes_f + 1
								if per_mes_f > 12 then'Reinicia a contagem de meses, caso seja igual a 12 e passa ao ano seguinte.
									per_mes_f = 1
									per_ano_f = per_ano_f + 1
								end if
						end if
					
				next
				per_hor_f = "00:00"
				per_min_f = "00:00"

'If cd_cid = "" Then '*** Mostra campo normal ***%>

Até: <input type="text" name="per_dia_f" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(per_dia_f)%>">/
	<input type="text" name="per_mes_f" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(per_mes_f)%>">/
	<input type="text" name="per_ano_f" size="4" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="<%=per_ano_f%>">&nbsp;
	<input type="text" name="per_hor_f" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">:
	<input type="text" name="per_min_f" size="2" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="">
										
<%'rs_cid.close
'set rs_cid = nothing%>