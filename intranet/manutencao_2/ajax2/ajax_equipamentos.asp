<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%xsql = "SELECT * FROM TBL_equipamento Where cd_status = 1 order by nm_equipamento"
Set rs = dbconn.execute(xsql)%>
<select size="1" id="lista" name="cd_equipamento">
	<option value="" class="inputs">Selecione</option>
	<%while not rs.EOF%>
	<option value="<%=rs("cd_codigo")%>" class="inputs"><%=rs("nm_equipamento")%></option>
	<%rs.MoveNext
	  wend
	
	  rs.close
	  set rs = nothing%>
</select>&nbsp;