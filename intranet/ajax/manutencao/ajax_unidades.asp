<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

'xsql = "SELECT * FROM TBL_equipamento WHERE nm_equipamento = '" & txtBusca & "' "
'xsql = "SELECT * FROM TBL_equipamento WHERE nm_equipamento LIKE '" & txtBusca & "%' "
xsql = "SELECT * FROM TBL_unidades order by nm_unidade"

Set rs_uni = dbconn.execute(xsql)%>

<select size="1" id="lista" name="nm_unidade">
<option value="" class="inputs"></option>
<%while not rs_uni.EOF%>
<option value="<%=rs_uni("nm_sigla")%>" class="inputs"><%=rs_uni("nm_sigla")%> - <%=rs_uni("nm_unidade")%></option>
<%rs_uni.MoveNext
  wend

  rs_uni.close
  set rs_uni = nothing%>
</select>