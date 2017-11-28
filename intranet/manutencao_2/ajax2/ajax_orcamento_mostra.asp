<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'xsql = "SELECT * FROM TBL_marca order by nm_marca"
num_os_assist  =request("num_os_assist")

if num_os_assist <> "" Then
xsql = "SELECT * FROM VIEW_manutencao_orcamento where nr_orcamento='"&num_os_assist&"'"
Set rs = dbconn.execute(xsql)%>
	<b>
	<%while not rs.EOF
		nm_valor = rs("nm_valor")
		dt_orcamento = rs("dt_orcamento")%>
		<input type="hidden" name="dt_manut_orc" value="<%=dt_orcamento%>">
		<%=dt_orcamento%><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0">
		<%if nm_valor <> "" Then response.write("R$"&nm_valor)%>
	<%rs.MoveNext
	wend
	
	rs.close
	set rs = nothing%>
	</b>
<%Else%>
	<input type="text" name="dt_manut_orc" size="11" maxlength="50" class="inputs" value="<%=dt_manut_orc%>">
<%end if%>