<!--#include file="../../includes/inc_open_connection.asp"-->
<%
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")


acao = request("acao")



If acao = "encerra" Then

   xsql = "up_Fechamento_mes @cd_mes='"&mes_sel&"', @cd_ano='"&ano_sel&"', @nm_fechamento=true"
	dbconn.execute(xsql)
	strmsg = "Benefício cadastrado com sucesso!"
	response.redirect("../../coop_fecha_mes.asp")
 End If%>