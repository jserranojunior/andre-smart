<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

verifica_empresa = busca_ajax(request("verifica_empresa"))

num_lista = 1	%>


<%if verifica_empresa <> "" then%>
	<%xsql = "SELECT * FROM TBL_empresa where nm_empresa = '"&verifica_empresa&"' and cd_status = 1 ORDER BY nm_empresa"
	Set rs_empresa = dbconn.execute(xsql)
	if rs_empresa.EOF Then%>
		<input type="submit" name="envia" value="Cadastra">
		<input type="hidden" name="existe" value="1">	
	<%else%>
		<span style="color:red;"><b>&nbsp;A empresa informada já existe</b></span>
		<input type="hidden" name="existe" value="">
	<%End if
	  rs_empresa.close
	  set rs_empresa = nothing
Else%>
	<input type="hidden" name="existe" value="">
<%end if%>