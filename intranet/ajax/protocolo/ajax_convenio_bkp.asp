<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_convenio = request("cd_convenio")
'response.write("teste")

If cd_convenio = "" Then '*** Mostra campo normal ***

	xsql = "SELECT * FROM TBL_convenios where cd_status = 1 Order by nm_convenio"
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

	xsql = "SELECT * FROM TBL_convenios where nm_convenio like '%"&busca_inteligente(cd_convenio)&"%' AND cd_status = 1"
	Set rs_conv = dbconn.execute(xsql)
	if not rs_conv.EOF then
		cd_cod_conv = rs_conv("cd_codigo")
		nm_convenio = rs_conv("nm_convenio")
		
		response.write(cd_cod_conv&" - "&nm_convenio)%>
		<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="alimenta_convenio(<%=cd_cod_conv%>);">
	<%else%>
		<b style="color:red;">Convênio Médico Não encontrado</b>
		<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="alimenta_convenio();">
	<%end if
	
	
	
	
	
	'xsql = "SELECT * FROM TBL_convenios where nm_convenio like '%"&busca_inteligente(cd_convenio)&"%' AND cd_status = 1"
	'Set rs_conv = dbconn.execute(xsql)%>
	
	<!--select size="1" id="lista" name="cd_convenio_1"-->
	<%'if not rs_conv.EOF Then
		'while not rs_conv.EOF	
		'cd_cod_conv = rs_conv("cd_codigo")
		'nm_convenio = rs_conv("nm_convenio")%>
		<!--option value="<%'=cd_cod_conv%>" class="inputs"><%'=nm_convenio%></option-->
		<%'rs_conv.MoveNext
		'wend
	'Else%>
		<!--option value="" class="inputs">* Convenio não encontrado *</option-->
	<%'end if
	
	'rs_conv.close
	' set rs_conv = nothing%>

<%Else%>
	<%xsql = "SELECT * FROM TBL_convenios where cd_status = 1 Order by nm_convenio"
	Set rs_conv = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_convenio_1">
	<option value="" class="inputs"></option>
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