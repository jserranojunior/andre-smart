<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'xsql = "SELECT * FROM TBL_marca order by nm_marca"
cd_fornecedor  = request("cd_fornecedor")
num_os_assist = request("num_os_assist")

if len(num_os_assist) >= 2 Then
'xsql = "SELECT * FROM VIEW_manutencao_orcamento where nr_orcamento='"&num_os_assist&"'"
xsql = "sp_orcamento_s01 @funcao = 'OF', @cd_fornecedor='"&cd_fornecedor&"', @nr_orcamento='"&num_os_assist&"'"
Set rs = dbconn.execute(xsql)%>
	<b>
	<%while not rs.EOF
		cd_codigo = rs("cd_codigo")
		nr_orcamento = rs("nr_orcamento")
		nm_valor = rs("nm_valor")
		dt_orcamento = rs("dt_orcamento")
		cd_fornecedor = rs("cd_fornecedor")
		dt_aprov_orc = rs("dt_aprov_orc")
		%>
		<input type="hidden" name="dt_manut_orcx" value="<%=dt_orcamento%>">
		<%=dt_orcamento%><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0">
		<%if isnumeric(nm_valor) Then response.write("R$ "&formatnumber(nm_valor,2))%>
	<%rs.MoveNext
	wend
	
	rs.close
	set rs = nothing
	if nr_orcamento = "" Then
		response.write("O orçamento não foi localizado")
	end if%>
	</b>
	<img src="../../imagens/px.gif" alt="" width="1" height="1" border="0" id="myImage" onload="alimenta_info_orcamento('<%=cd_codigo%>','<%=dt_orcamento%>','<%=dt_aprov_orc%>');">
<%Else%>
	<!--input type="text" name="dt_manut_orc" size="11" maxlength="50" class="inputs" value="<%=dt_manut_orc%>"-->
	
<%end if%>