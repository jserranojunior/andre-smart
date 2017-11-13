<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

'xsql = "SELECT * FROM TBL_equipamento WHERE nm_equipamento = '" & txtBusca & "' "
'xsql = "SELECT * FROM TBL_equipamento WHERE nm_equipamento LIKE '" & txtBusca & "%' "
xsql = "SELECT * FROM TBL_equipamento order by nm_equipamento"

Set rs_equip = dbconn.execute(xsql)%>

<select size="1" id="lista" name="cd_equipamento">
<option value="" class="inputs"></option>
<%while not rs_equip.EOF%>
<option value="<%=rs_equip("cd_codigo")%>" class="inputs"><%=rs_equip("nm_equipamento")%></option>
<%rs_equip.MoveNext
  wend

  rs_equip.close
  set rs_equip = nothing%>
</select>