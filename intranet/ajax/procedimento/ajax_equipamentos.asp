<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

procedimentos = request("procedimentos")
	if right(procedimentos,1) = "," Then
	proc_len = len(procedimentos)
	proc_len = proc_len - 1
			procedimentos = Left(procedimentos,proc_len)
	end if

proc_array = split(procedimentos,",")
		For Each proc_item In proc_array
		
		
xsql = "SELECT * FROM TBL_protocolo_procedimento order by nm_equipamento"

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