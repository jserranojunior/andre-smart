<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_convenio = request("cd_convenio")

If cd_convenio = "" Then '*** Mostra campo normal ***

	xsql = "SELECT * FROM TBL_convenios Order by nm_convenio"
	Set rs_conv = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_convenio_1">
	<option value="" class="inputs"></option>
	<%while not rs_conv.EOF%>
	<%cd_cod_conv = rs_conv("cd_codigo")
	nm_convenio = rs_conv("nm_convenio")%>
	
	<option value="<%=cd_cod_conv%>" class="inputs"><%=nm_convenio%> </option>
	<%rs_conv.MoveNext
	conv_sel = ""
	  wend
	
	  rs_conv.close
	  set rs_conv = nothing%>
	</select>
<%Elseif not IsNumeric (cd_convenio) then ' *** Mostra Campo não numérico ***

	xsql = "SELECT * FROM TBL_convenios where nm_convenio like '"&cd_convenio&"%'"
	Set rs_conv = dbconn.execute(xsql)%>
	
	<select size="1" id="lista" name="cd_convenio_1">
	<%while not rs_conv.EOF%>
	<%
	cd_cod_conv = rs_conv("cd_codigo")
	nm_convenio = rs_conv("nm_convenio")%>
	
	<option value="<%=cd_cod_conv%>" class="inputs"><%=nm_convenio%></option>
	<%rs_conv.MoveNext
	wend
	
	  rs_conv.close
	  set rs_conv = nothing%>

<%Else%>
	<%xsql = "SELECT * FROM TBL_convenios Order by nm_convenio"
	Set rs_conv = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_convenio_1">
	<%while not rs_conv.EOF%>
	<%cd_cod_conv = rs_conv("cd_codigo")
	nm_convenio = rs_conv("nm_convenio")
	if int(cd_convenio) = int(cd_cod_conv) then%><%conv_sel = "selected"%><%end if%>
	<option value="<%=cd_cod_conv%>" class="inputs" <%=conv_sel%>><%=nm_convenio%></option>
	<%rs_conv.MoveNext
	conv_sel = ""
	  wend
	
	  rs_conv.close
	  set rs_conv = nothing%>
	 </select>

<%end if%>