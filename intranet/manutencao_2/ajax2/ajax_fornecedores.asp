<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

'xsql = "SELECT * FROM TBL_fornecedor WHERE nm_fornecedor = '" & txtBusca & "' "
'xsql = "SELECT * FROM TBL_fornecedor WHERE nm_fornecedor LIKE '" & txtBusca & "%' "
xsql = "SELECT * FROM TBL_fornecedor order by nm_fornecedor"

Set rs_forn = dbconn.execute(xsql)%>

<select size="1" id="lista" name="cd_fornecedor">
<option value="" class="inputs"></option>
<%while not rs_forn.EOF%>
<option value="<%=rs_forn("cd_codigo")%>" class="inputs"><%=rs_forn("nm_fornecedor")%></option>
<%rs_forn.MoveNext
  wend

  rs_forn.close
  set rs_forn = nothing%>
</select>


