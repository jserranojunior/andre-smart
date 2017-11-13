<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_especialidade = request("cd_especialidade")

If cd_especialidade = "" Then '*** Mostra campo normal ***%>
	<%xsql = "SELECT * FROM TBL_espec where a057_status = 1 Order by nm_especialidade"
	Set rs_espec = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_especialidade_1" onChange="alimenta_especialidade(this.value);" onBlur="alimenta_especialidade(this.value);">
	<option value="" class="inputs"></option>
	<%while not rs_espec.EOF%>
	<%cd_cod_espec = rs_espec("cd_codigo")
	nm_especialidade = rs_espec("nm_especialidade")%>	
	<option value="<%=cd_cod_espec%>" class="inputs"><%=nm_especialidade%>*</option>
	<%rs_espec.MoveNext
	wend%>
	</select>
	
<%Elseif IsNumeric(cd_especialidade) then ' *** Mostra Campo numérico ***%>
	
	
	<%xsql = "SELECT * FROM TBL_espec where cd_codigo = '"&cd_especialidade&"' AND a057_status = 1 order by nm_especialidade"
		Set rs_espec = dbconn.execute(xsql)%>	
		<select size="1" id="lista" name="cd_especialidade_1" onChange="alimenta_especialidade(<%=cd_cod_conv%>);" onBlur="alimenta_especialidade(this.value<%'=cd_cod_conv%>);">
	<%if not rs_espec.EOF Then%>
		<%while not rs_espec.EOF%>
		<%cd_cod_espec = rs_espec("cd_codigo")
		nm_especialidade = rs_espec("nm_especialidade")%>
		<option value="<%=cd_cod_espec%>" class="inputs"><%=cd_cod_espec%> - <%=nm_especialidade%></option>
		<%rs_espec.MoveNext
		wend%>
		
	<%else%>
		<option value="" class="inputs" style="color:red;">* Especialidade não encontrada *</option>
	<%end if%>
	
		</select>
	<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="alimenta_especialidade(<%=cd_cod_espec%>);">
<%Elseif not IsNumeric(cd_especialidade) then%>
		<%xsql = "SELECT * FROM TBL_espec where nm_especialidade like '"&busca_inteligente(cd_especialidade)&"%' AND a057_status = 1 order by nm_especialidade"
		Set rs_espec = dbconn.execute(xsql)%>	
		<select size="1" id="lista" name="cd_especialidade_1" onChange="alimenta_especialidade(<%=cd_cod_conv%>);" onBlurs="alimenta_especialidade(this.value<%'=cd_cod_conv%>);">
	<%if not rs_espec.EOF Then%>
		<%while not rs_espec.EOF%>
		<%cd_cod_espec = rs_espec("cd_codigo")
		nm_especialidade = rs_espec("nm_especialidade")%>
		<option value="<%=cd_cod_espec%>" class="inputs"><%=cd_cod_espec%> - <%=nm_especialidade%></option>
		<%rs_espec.MoveNext
		wend%>
		
	<%else%>
		<option value="" class="inputs" style="color:red;">* Especialidade não encontrada *</option>
	<%end if%>
	
		</select>
	<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="alimenta_especialidade(<%=cd_cod_espec%>);">
	
<%end if%>


<%rs_espec.close
set rs_espec = nothing%>