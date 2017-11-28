<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->

<%nr_orcamento = request("nr_orcamento")


if nr_orcamento <> "" Then

	'xsql = "SELECT * FROM TBL_equipamento WHERE nm_equipamento = '" & txtBusca & "' "
	'xsql = "SELECT * FROM TBL_equipamento WHERE nm_equipamento LIKE '" & txtBusca & "%' "
	xsql = "SELECT * FROM TBL_manutencao_orcamento WHERE nr_orcamento = '"&nr_orcamento&"'"
	Set rs_sit = dbconn.execute(xsql)
	
	if not rs_sit.EOF Then%>
		<input type="button" disabled="disabled" value="Gravar" ><br>
		<span style="color:red;">N° Orçamento existente</span>
	<%else%>
		<input type="submit" value="Gravar">
	<%end if%>
<%else%>
		<input type="button" disabled="disabled" value="Gravar" >
<%end if

response.write(cd_orcamento)%>