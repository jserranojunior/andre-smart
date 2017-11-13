<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_crm = request("cd_crm")

If cd_crm = "" Then '*** Mostra campo normal ***%>
	<%xsql = "SELECT * FROM TBL_medicos where a055_status = 1 Order by a055_desmed"
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
	
	
<%Elseif IsNumeric (cd_crm) AND session("cd_codigo") = 460 then ' *** Mostra Campo numérico ***%>
	<%xsql = "SELECT * FROM TBL_medicos where a055_status = 1 Order by a055_desmed" '*** Mostra pelo nome
	Set rs_crm = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_crm_1" onChange="alimenta_medico(this.value);" onBlur="alimenta_medico(this.value);">
		<option value="" class="inputs" style="color:red;">*** Médico não encontrado ***</option>
		<%while not rs_crm.EOF%>
		<%cd_cod_crm = rs_crm("a055_codmed")
		num_crm = rs_crm("a055_crmmed")
		nm_medico = rs_crm("a055_desmed")
		if int(cd_crm) = int(num_crm) then%><%med_sel = "selected"%><%end if%>
			<option value="<%=cd_cod_crm%>" class="inputs" <%=med_sel%>><%=cd_cod_crm%> x <%'=num_crm%><%=nm_medico%>%></option>
		<%rs_crm.MoveNext
		med_sel = ""
		  wend
		
		  rs_crm.close
		  set rs_crm = nothing%>
	 </select>
	 
<%Elseif IsNumeric (cd_crm) then  ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_medicos where a055_crmmed = '"&busca_inteligente(cd_crm)&"' AND a055_status = 1 order by a055_desmed"
	Set rs_crm = dbconn.execute(xsql)%>
	
	<select size="1" id="lista" name="cd_crm_1" onChange="alimenta_medico(this.value);" onBlur="alimenta_medico(this.value);">
	<%if not rs_crm.EOF Then%>
		<%while not rs_crm.EOF%>
		<%cd_cod_crm = rs_crm("a055_codmed")
		num_crm = rs_crm("a055_crmmed")
		nm_medico = rs_crm("a055_desmed")%>		
		<option value="<%=cd_cod_crm%>" class="inputs"><%=num_crm%> - <%=nm_medico%>nnn</option>
		<%rs_crm.MoveNext
		wend
	else%>
		<option value="" class="inputs" style="color:red;">*** Médico não encontrado ***</option>
		response.write("")
	<%End if
	  rs_crm.close
	  set rs_crm = nothing%>
	</select>
	<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="alimenta_medico(<%=cd_cod_crm%>);">
		
	 
<%Elseif not IsNumeric (cd_crm) then  ' *** Mostra Campo não numérico ***%>
	<%xsql = "SELECT * FROM TBL_medicos where a055_desmed like '%"&busca_inteligente(cd_crm)&"%' AND a055_status = 1 order by a055_desmed"
	Set rs_crm = dbconn.execute(xsql)%>
	
	<select size="1" id="lista" name="cd_crm_1" onChange="alimenta_medico(this.value);" onBlur="alimenta_medico(this.value);">
	<%if not rs_crm.EOF Then%>
		<%while not rs_crm.EOF%>
		<%cd_cod_crm = rs_crm("a055_codmed")
		num_crm = rs_crm("a055_crmmed")
		nm_medico = rs_crm("a055_desmed")%>		
		<option value="<%=cd_cod_crm%>" class="inputs"><%=num_crm%> - <%=nm_medico%></option>
		<%rs_crm.MoveNext
		wend
	else%>
		<option value="" class="inputs" style="color:red;">*** Médico não encontrado ***</option>
		response.write("")
	<%End if
	  rs_crm.close
	  set rs_crm = nothing%>
	</select>
	<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="alimenta_medico(<%=cd_cod_crm%>);">
	
<%end if%>