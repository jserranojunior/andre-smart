<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%strsql = "Select * From TBL_Aparelhos_comunicacao order by nr_numero"
	Set rs_fornec = dbconn.execute(strsql)%>
	<select name="cd_aparelho" class="inputs">
		<option value=""></option>
		<%While Not rs.eof
			cd_aparel = rs("cd_codigo")
			nr_numero = rs("nm_fornecedor")%>	
			<option value="<%=cd_aparel%>" <%if int(strcd_aparelho) = cd_aparelthen%>Selected<%end if%>><%=nr_numero%></option>
			<%rs.movenext
		Wend
			rs.close
			Set rs = nothing%>
				</select>
				
				
				asd