<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

materiais = request("materiais")
	if right(materiais,1) = "," Then
	mat_len = len(materiais)
	mat_len = proc_len - 1
			procedimentos = Left(materiais,mat_len)
	end if

mat_array = split(materiais,",")
		For Each mat_item In mat_array
		
		
xsql = "SELECT * FROM TBL_protocolo_material order by nm_material"

Set rs_mat = dbconn.execute(xsql)%>

<select size="1" id="lista" name="cd_material">
<option value="" class="inputs"></option>
<%while not rs_mat.EOF%>
<option value="<%=rs_mat("cd_codigo")%>" class="inputs"><%=rs_mat("nm_equipamento")%></option>
<%rs_mat.MoveNext
  wend

  rs_mat.close
  set rs_mat = nothing%>
</select>