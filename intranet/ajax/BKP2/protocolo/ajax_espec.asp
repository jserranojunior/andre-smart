<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_especialidade = request("cd_especialidade")

If cd_especialidade = "" Then '*** Mostra campo normal ***%>
	<%xsql = "SELECT * FROM TBL_espec Order by nm_especialidade"
	Set rs_espec = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_especialidade_1">
	<option value="" class="inputs"></option>
	<%while not rs_espec.EOF%>
	<%cd_cod_espec = rs_espec("cd_codigo")
	nm_especialidade = rs_espec("nm_especialidade")%>	
	<option value="<%=cd_cod_espec%>" class="inputs"><%=nm_especialidade%>*</option>
	<%rs_espec.MoveNext
	wend%>
	</select>
	
<%Elseif IsNumeric(cd_especialidade) then ' *** Mostra Campo numérico ***%>
	<%xsql = "SELECT * FROM TBL_espec Order by cd_codigo" '*** Mostra pelo nome
	Set rs_espec = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_especialidade_1">
	<%while not rs_espec.EOF%>
	<%cd_cod_espec = rs_espec("cd_codigo")
	nm_especialidade = rs_espec("nm_especialidade")%>
	<%if int(cd_cod_espec) = int(cd_especialidade) then%><%espec_sel = "selected"%><%end if%>
	<option value="<%=cd_cod_espec%>" class="inputs" <%=espec_sel%>><%=nm_especialidade%></option>
	<%rs_espec.MoveNext
	espec_sel = ""
	wend%>
	 </select>

<%Elseif not IsNumeric(cd_especialidade) then%>
		<%xsql = "SELECT * FROM TBL_espec where nm_especialidade like '"&cd_especialidade&"%' order by nm_especialidade"
		Set rs_espec = dbconn.execute(xsql)%>	
		<select size="1" id="lista" name="cd_especialidade_1">
	<%if not rs_espec.EOF Then%>
		<%while not rs_espec.EOF%>
		<%cd_cod_espec = rs_espec("cd_codigo")
		nm_especialidade = rs_espec("nm_especialidade")%>
		<option value="<%=cd_cod_espec%>" class="inputs"><%=nm_especialidade%></option>
		<%rs_espec.MoveNext
		wend%>
	<%else%>
		<option value="" class="inputs">* Especialidade não encontrada *</option>
	<%end if%>
		</select>
	
	
<%end if%>


<%rs_espec.close
set rs_espec = nothing%>