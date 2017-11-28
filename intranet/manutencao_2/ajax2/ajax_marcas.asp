<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%xsql = "SELECT * FROM TBL_marca order by nm_marca"
Set rs_marca = dbconn.execute(xsql)%>
<select size="1" id="lista" name="cd_marca">
	<option value="" class="inputs">Selecione</option>
	<%while not rs_marca.EOF%>
	<option value="<%=rs_marca("cd_codigo")%>" class="inputs"><%=rs_marca("nm_marca")%></option>
	<%rs_marca.MoveNext
	  wend
	
	  rs_marca.close
	  set rs_marca = nothing%>
</select>