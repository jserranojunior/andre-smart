<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%strsql = "Select * From TBL_fornecedor order by nm_fornecedor"
	Set rs_fornec = dbconn.execute(strsql)%>
	<select name="cd_fornecedor" class="inputs">
		<option value=""></option>
		<%While Not rs_fornec.eof
			cd_fornec = rs_fornec("cd_codigo")
			nm_fornec = rs_fornec("nm_fornecedor")%>	
			<option value="<%=cd_fornec%>" <%if int(strcd_fornecedor) = cd_fornec then%>Selected<%end if%>><%=nm_fornec%></option>
			<%rs_fornec.movenext
		Wend
			rs_fornec.close
			Set rs_fornec = nothing%>
				</select>