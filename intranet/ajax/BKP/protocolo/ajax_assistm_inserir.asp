<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

'procedimentos = request("procedimentos")
'	if right(procedimentos,1) = "," Then
'	proc_len = len(procedimentos)
'	proc_len = proc_len - 1
'			procedimentos = Left(procedimentos,proc_len)
'	end if

'proc_array = split(procedimentos,",")
'		For Each proc_item In proc_array
		
		
xsql = "SELECT * FROM TBL_convenio order by nm_convenio"

Set rs_equip = dbconn.execute(xsql)%>




<select size="1" id="lista" name="cd_convenio">
<option value="" class="inputs"></option>
<%while not rs_equip.EOF%>
<option value="<%=rs_equip("cd_codigo")%>" class="inputs"><%=rs_equip("nm_equipamento")%></option>
<%rs_equip.MoveNext
  wend

  rs_equip.close
  set rs_equip = nothing%>
</select>


****
<%strsql ="Select * from TBL_convenio order by nm_convenio"
						  		Set rs_conv = dbconn.execute(strsql)%>
							<select name="cd_convenio_1" class="inputs" onFocus="nextfield ='nm_sexo';"><option value="">
								<%Do while Not rs_conv.EOF%><%cd_conv = rs_conv("cd_codigo")%>
								<%if cd_convenio=cd_conv Then%><%ck_conv="selected"%><%else ck_conv=""%><%end if%>
								<option value="<%=rs_conv("cd_codigo")%>" <%=ck_conv%>><%=rs_conv("nm_convenio")%></option><%rs_conv.movenext
						 		Loop%>
							</select>