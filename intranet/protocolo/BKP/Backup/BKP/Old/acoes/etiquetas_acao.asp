<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% 
acao = request("acao")
tipo_print = request("tipo_print")
cd_unidade = request("cd_unidade")
print_ok = request("print_ok")

if tipo_print = 2 then
	cd_unidade = 0
end if

	if cd_unidade = 0 then
		cd_unidade = 999
	end if
		
cd_num_inicio = request("cd_num_inicio")

qtd_paginas = request("qtd_paginas")
	cd_num_fim = int(cd_num_inicio) + int(qtd_paginas-1)
cd_user_inc = session("cd_codigo")

dt_dia = day(now)
dt_mes = month(now)
dt_ano = year(now)
dt_hr = hour(now)
dt_min = minute(now)
	dt_data_print = dt_mes&"/"&dt_dia&"/"&dt_ano&" "&dt_hr&":"&dt_min

'if print_ok = 0 OR print_ok = "" Then
if print_ok = "" Then%>

<table align="center" id="no_print" style="position:relative; top:20px; left:0px;">
	<tr>
		<td align="center" bgcolor="#00cc00" colspan="2" style="color:white";><b>CONFIRMAÇÃO</b></td>
	</tr>
	<%strsql ="Select * from TBL_unidades Where cd_codigo = "&cd_unidade&""
					Set rs_und = dbconn.execute(strsql)
					if not rs_und.EOF then
					nm_unidade = rs_und("nm_unidade")
					end if%>
	<tr><%if qtd_paginas > 1 Then 
		p1="ão" 
		p2="s" 
		Else
		p1="á"
		p2=""
		end if%>
		<td align="center"> A impressão da<%=p2%> &nbsp;<b><%=qtd_paginas%> página<%=p2%></b> do <br><b>Relatório Técnico de Cirúrgia</b><br><%if tipo_print <> 2 Then%>para o <b><%=nm_unidade%></b><%else%>em branco<%end if%>, <br>está OK?</td>
	</tr>
	<tr>
		<td align="center">&nbsp;<br><a href="etiquetas_acao.asp?cd_unidade=<%=cd_unidade%>&cd_num_inicio=<%=cd_num_inicio%>&qtd_paginas=<%=qtd_paginas%>&print_ok=1"><b>Sim</b></a> &nbsp; &nbsp; &nbsp; <a href="../../protocolo.asp?tipo=etiquetas&mensagem=2"><b>Não</b></a><p style="color:red;"></p></td>
	</tr>
</table>
<%
'response.redirect("etiquetas_acao.asp?cd_unidade=22&cd_num_inicio=66&qtd_paginas=100&tipo_print=1&session_cd_usuario")
'response.redirect("../../protocolo.asp?tipo=etiquetas&mensagem=1")cd_unidade=22&cd_num_inicio=66&qtd_paginas=100&tipo_print=1&session_cd_usuario


Elseif print_ok = 1 Then

	if cd_num_inicio <> "" then
	xsql = "up_protocolo_form_insert @cd_unidade='"&cd_unidade&"', @cd_num_inicio='"&cd_num_inicio&"', @cd_num_fim='"&cd_num_fim&"', @dt_data='"&dt_data_print&"', @cd_user_inc='"&cd_user_inc&"'"
		dbconn.execute(xsql)
		redirecionamento = "../../protocolo.asp?tipo=etiquetas&mensagem=1"
		response.redirect(redirecionamento)
		response.write("Gravação OK")
	end if
	
End if

'Fim da manipulação do banco de dados
%>
<!--Und: <%response.write(cd_unidade&"<br>")%>
Inicio: <%response.write(cd_num_inicio&"<br>")%>
Fim: <%response.write(cd_num_fim&"<br>")%>
Data: <%response.write(dt_data_print&"<br>")%>
user: <%response.write(cd_user_inc&"<br>")%>
Páginas: <%response.write(qtd_paginas&"<br>")%>
Print_ok?: <%response.write(print_ok)%>-->

<%'response.redirect(redirecionamento)%>