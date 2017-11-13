<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_convenio = request("cd_convenio")

	'if not IsNumeric(cd_convenio) then
		'xsql_teste = "SELECT * FROM TBL_convenios where nm_convenio Like '%"&cd_convenio&"'"
		'Set rs_conv_teste = dbconn.execute(xsql_teste)
		'while not rs_conv_teste.EOF
		'rs_conv_teste("
		'xsql = "SELECT * FROM TBL_convenios where nm_convenio Like '%"&cd_convenio&"'"
		'xsql = "SELECT * FROM TBL_convenios where nm_convenio Like '%pad%'"
	'else
		'xsql = "SELECT * FROM TBL_convenios Order by nm_convenio" 'where cd_codigo = '"&cd_convenio&"'"
		
		
		'cd_convenio = codconvenio
	'end if

xsql = "SELECT * FROM TBL_convenios Order by nm_convenio" 'where cd_codigo = '"&cd_convenio&"'"

Set rs_conv = dbconn.execute(xsql)%>

<select size="1" id="lista" name="cd_convenio_1">
<option value="" class="inputs"></option>
<%while not rs_conv.EOF%>
<%
cd_cod_conv = rs_conv("cd_codigo")
nm_convenio = rs_conv("nm_convenio")

if int(cd_convenio) = int(cd_cod_conv) then%><%conv_sel = "selected"%><%end if%>
<option value="<%=cd_cod_conv%>" class="inputs" <%=conv_sel%>><%=nm_convenio%> <%'=cd_convenio&"..."&cd_cod_conv&"-"&conv_sel%></option>
<%rs_conv.MoveNext
conv_sel = ""
  wend

  rs_conv.close
  set rs_conv = nothing%>
</select>