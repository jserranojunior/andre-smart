<!--#include file="../../includes/inc_open_connection.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%sel_patrimonio = request("no_patrimonio")
		xsql = "SELECT * FROM View_patrimonio_lista Where cd_codigo = "&sel_patrimonio&""
		Set rs = dbconn.execute(xsql)%>
		
			<%'if not rs.EOF Then%>
				<%cd_pat = rs("cd_codigo")
				no_patrimonio = rs("no_patrimonio")
				cd_equipamento = rs("cd_equipamento")
				nm_equipamento = rs("nm_equipamento")
				cd_marca = rs("cd_marca")
				nm_marca = rs("nm_marca")
				cd_especialidade = rs("cd_especialidade")
				nm_especialidade = rs("nm_especialidade")
				cd_ns = rs("cd_ns")%>
							<input type="hidden" name="cd_patrimonio" value="<%=cd_pat%>">
							<input type="hidden" name="no_patrimonio" value="<%=no_patrimonio%>">
							<input type="hidden" name="cd_equipamento" value="<%=cd_equipamento%>">
							<input type="hidden" name="cd_marca" value="<%=cd_marca%>">
							<input type="hidden" name="cd_especialidade" value="<%=cd_especialidade%>">
							<input type="hidden" name="cd_ns" value="<%=cd_ns%>">
							<input type="hidden" name="num_qtd" value="1">
			<%
			'rs.MoveNext
			'end if
			
			  rs.close
			  set rs = nothing%>