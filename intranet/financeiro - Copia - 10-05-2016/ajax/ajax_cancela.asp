<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../css/geral.htm"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->

<%str_cancela = request("cod_cancela")
%>
<%div_str = instr(1,str_cancela,"_",1)
dt_cancela = mid(str_cancela,1,div_str)
	dt_cancela = replace(dt_cancela,"_","")
dt_selecionada =  mid(str_cancela,div_str,len(str_cancela))
	dt_selecionada = replace(dt_selecionada,"_","")

if int(dt_selecionada) > int(dt_cancela) then
	ano_sel_canc = mid(dt_selecionada,1,4)
	mes_sel_canc = mid(dt_selecionada,5,2)%>
<!--input type="text" name="cd_cancela" id="cd_cancela" value="" onclick="cancela_pagamento('2','2014')";-->

<a href="javascript:void(0);" onclick="cancela_pagamento('<%=mes_sel_canc%>','<%=ano_sel_canc%>');" style="">
<img src="../../imagens/check_inativo.gif" alt="" width="25" height="12" border="0" id="check_cancela"> << Clique para cancelar a partir de <b>01/<%=mesdoano(int(mes_sel_canc))&"/"&ano_sel_canc%></b></a>.<%'=str_cancela%>
<%else%>
Não é possível cancelar, pois a data<br>inicial é igual a data de cancelamento.
<%end if%>



