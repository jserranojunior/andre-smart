<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_rack = request("cd_rack")

If cd_rack = "" Then '*** Mostra campo normal ***%>
	<%xsql = "SELECT * FROM TBL_rack Order by a056_desrac"
	Set rs_rack = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_rack_1">
	<%while not rs_rack.EOF%>
	<%cd_cod_rack = rs_rack("a056_codrac")
	nm_rack = rs_rack("a056_desrac")%>	
	<option value="<%=cd_cod_rack%>" class="inputs"><%=nm_rack%></option>
	<%rs_rack.MoveNext
	wend%>
	</select>
	
<%Elseif IsNumeric (cd_rack) then ' *** Mostra Campo numérico ***%>
	<%xsql = "SELECT * FROM TBL_rack Order by a056_codrac" '*** Mostra pelo nome
	Set rs_rack = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_rack_1">
	<%while not rs_rack.EOF%>
	<%cd_cod_rack = rs_rack("a056_codrac")
	nm_rack = rs_rack("a056_desrac")
	if int(cd_cod_rack) = int(cd_rack) then%><%rack_sel = "selected"%><%end if%>
	<option value="<%=cd_cod_rack%>" class="inputs" <%=rack_sel%>><%=nm_rack%></option>
	<%rs_rack.MoveNext
	rack_sel = ""
	wend%>
	 </select>

<%Elseif not IsNumeric (cd_rack) then%>
	<%xsql = "SELECT * FROM TBL_rack where a056_desrac like '"&cd_rack&"%' order by a056_desrac"
	Set rs_rack = dbconn.execute(xsql)%>	
	<select size="1" id="lista" name="cd_rack_1">
	<%while not rs_rack.EOF%>
	<%cd_cod_rack = rs_rack("a056_codrac")
	nm_rack = rs_rack("a056_desrac")%>
	<option value="<%=cd_cod_rack%>" class="inputs"><%=nm_rack%> - <%=cd_rack%></option>
	<%rs_rack.MoveNext
	wend%>
	</select>
<%end if%>


<%rs_rack.close
set rs_rack = nothing%>