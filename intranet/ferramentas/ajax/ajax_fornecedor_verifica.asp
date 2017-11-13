<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")

verifica_nome = busca_ajax(request("verifica_nome"))

num_lista = 1	%>


<%if verifica_nome <> "" then%>
	<%xsql = "SELECT * FROM View_fornecedores where nm_fornecedor like '"&verifica_nome&"' and cd_status = 1 ORDER BY nm_fornecedor"
	Set rs_fornecedor = dbconn.execute(xsql)
	if rs_fornecedor.EOF Then%>
		<input type="submit" name="envia" value="Cadastra">
		<input type="hidden" name="existe" value="1">	
	<%else%>
		<span style="color:red;"><b>&nbsp;O nome informado já existe</b></span>
		<input type="hidden" name="existe" value="">
	<%End if
	  rs_fornecedor.close
	  set rs_fornecedor = nothing
Else%>
	<input type="hidden" name="existe" value="">
<%end if%>