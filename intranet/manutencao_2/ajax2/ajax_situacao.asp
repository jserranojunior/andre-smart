<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

'xsql = "SELECT * FROM TBL_equipamento WHERE nm_equipamento = '" & txtBusca & "' "
'xsql = "SELECT * FROM TBL_equipamento WHERE nm_equipamento LIKE '" & txtBusca & "%' "
xsql = "SELECT * FROM TBL_situacao order by nm_situacao"

Set rs_sit = dbconn.execute(xsql)%>

<select size="1" id="lista" name="cd_situacao">
<option value="" class="inputs"></option>
<%while not rs_sit.EOF%>
<option value="<%=rs_sit("cd_codigo")%>" class="inputs"><%=rs_sit("nm_situacao")%></option>
<%rs_sit.MoveNext
  wend

  rs_sit.close
  set rs_sit = nothing%>
</select>
