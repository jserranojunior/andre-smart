<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_procedimento = request("cd_procedimento")
'cd_procedimentos = "100"

		xsql = "SELECT * FROM TBL_procedimento where cd_codigo="&cd_procedimentos&""
			Set rs_proc1 = dbconn.execute(xsql)
			while not rs_proc1.EOF
			cd_cod_proced =  rs_proc1("cd_codigo")
			nm_procedimento = rs_proc1("nm_procedimento")%>
			(
			<%=cd_cod_proced&" - "&nm_procedimento%>
			)
			<%rs_proc1.MoveNext
			 wend
			
			  rs_proc1.close
			  set rs_proc1 = nothing
%>