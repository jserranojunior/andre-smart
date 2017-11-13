<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

verifica_nome = busca_ajax(request("verifica_nome"))

num_lista = 1	%>


<%if verifica_nome <> "" then%>
	<%xsql = "SELECT * FROM View_equipamento where nm_equipamento_novo like '"&verifica_nome&"' and cd_status = 1 ORDER BY nm_equipamento_novo"
	Set rs_equipamento = dbconn.execute(xsql)
	if rs_equipamento.EOF Then%>
		<!--input type="submit" name="envia" value="Cadastra"-->
		<input type="button" name="envia" value="Cadastra" onclick="document.getElementById('Form').submit">
		<input type="hidden" name="existe" value="1">	
	<%else%>
		<span style="color:red;"><b>&nbsp;O nome informado já existe</b></span>
		<input type="hidden" name="existe" value="">
	<%End if
	  rs_equipamento.close
	  set rs_equipamento = nothing
Else%>
	<input type="hidden" name="existe" value="">
<%end if%>