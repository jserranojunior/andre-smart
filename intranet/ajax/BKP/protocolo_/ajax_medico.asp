<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentua��o dos ajax-->
<%'txtBusca = request("txtBusca")

cd_crm = request("cd_crm")

If cd_crm = "" Then '*** Mostra campo normal ***%>
	<%xsql = "SELECT * FROM TBL_medicos Order by a055_desmed"
	Set rs_crm = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_crm_1">
	<option value="" class="inputs"></option>
	<%while not rs_crm.EOF%>
	<%cd_cod_med = rs_crm("a055_codmed")
	nm_medico = rs_crm("a055_desmed")%>
	
	<option value="<%=cd_cod_med%>" class="inputs"><%=nm_medico%></option>
	<%rs_crm.MoveNext
	wend
	
	  rs_crm.close
	  set rs_crm = nothing%>
	</select>
<%Elseif IsNumeric (cd_crm) then ' *** Mostra Campo n�o num�rico ***%>
	<%xsql = "SELECT * FROM TBL_medicos Order by a055_desmed" '*** Mostra pelo nome
	Set rs_crm = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_crm_1">
	<option value="" class="inputs"></option>
	<%while not rs_crm.EOF%>
	<%cd_cod_crm = rs_crm("a055_codmed")
	num_crm = rs_crm("a055_crmmed")
	nm_medico = rs_crm("a055_desmed")
	if int(cd_crm) = int(num_crm) then%><%med_sel = "selected"%><%end if%>
		<option value="<%=cd_cod_crm%>" class="inputs" <%=med_sel%>><%=nm_medico%></option>
	<%rs_crm.MoveNext
	med_sel = ""
	  wend
	
	  rs_crm.close
	  set rs_crm = nothing%>
	 </select>

	
<%Elseif not IsNumeric (cd_crm) then%>
	<%xsql = "SELECT * FROM TBL_medicos where a055_desmed like '"&cd_crm& "%' order by a055_desmed"
	Set rs_crm = dbconn.execute(xsql)%>
	
	<select size="1" id="lista" name="cd_crm_1">
	<%while not rs_crm.EOF%>
	<%cd_cod_crm = rs_crm("a055_codmed")
	cd_crmmed = rs_crm("a055_crmmed")
	nm_medico = rs_crm("a055_desmed")%>
	
	<option value="<%=cd_cod_crm%>" class="inputs"><%=cd_crmmed%> - <%=nm_medico%></option>
	<%rs_crm.MoveNext
	wend
	
	  rs_crm.close
	  set rs_crm = nothing%>
	</select>
<%end if%>