<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% 
acao = request("acao")
tipo_print = request("tipo_print")
janela = request("janela")
cd_unidade = request("cd_unidade")
print_ok = request("print_ok")
cd_codigo = request("cd_codigo")
cd_confirma = request("cd_confirma")
	if cd_confirma = 3 Then
		acao = "delete"
	end if

if tipo_print = 2 then
	cd_unidade = 0
end if

	if cd_unidade = 0 then
		cd_unidade = 999
	end if
		
qtd_paginas = request("qtd_paginas")
qtd_etiquetas = request("qtd_etiquetas")
	if qtd_etiquetas <> "" Then
		qtd_paginas = 65
	end if

cd_num_inicio = request("cd_num_inicio")
cd_num_fim = request("cd_num_fim")
	if cd_num_inicio <> "" AND cd_num_fim = ""Then
		cd_num_fim = int(cd_num_inicio) + int(qtd_paginas-1)
	elseif cd_num_inicio <> "" AND cd_num_fim <> "" Then
		cd_num_fim = cd_num_fim
	end if

cd_user_inc = session("cd_codigo")

dt_dia = day(now)
dt_mes = month(now)
dt_ano = year(now)
dt_hr = hour(now)
dt_min = minute(now)
	dt_data_print = dt_mes&"/"&dt_dia&"/"&dt_ano&" "&dt_hr&":"&dt_min


If acao = "insert" Then

	if cd_num_inicio <> ""  and novo_interv <> 1 then 'registra a impressão
	xsql = "up_protocolo_form_insert @cd_unidade='"&cd_unidade&"', @cd_num_inicio='"&cd_num_inicio&"', @cd_num_fim='"&cd_num_fim&"', @dt_data='"&dt_data_print&"', @cd_user_inc='"&cd_user_inc&"'"
		dbconn.execute(xsql)
		redirecionamento = "../../protocolo.asp?tipo=etiquetas&mensagem=1"
		response.redirect(redirecionamento)
		response.write("Gravação OK <br>")
	elseif cd_num_inicio <> "" and novo_interv = 1 then
		
	end if

Elseif acao = "update" then
	if cd_confirma = 1 then		
		xsql = "up_protocolo_form_confirma_update @cd_codigo = '"&cd_codigo&"'"
		dbconn.execute(xsql)
		response.write("!atualiza!")
	elseif cd_confirma = 2 Then
		If cd_num_inicio <> "" AND cd_num_fim <> "" Then
			xsql = "up_protocolo_form_altera_update @cd_codigo='"&cd_codigo&"', @cd_num_inicio='"&cd_num_inicio&"', @cd_num_fim='"&cd_num_fim&"'"		
		dbconn.execute(xsql)
		else
			response.write("!Campos em branco!")
		End if
	Else
		response.write("!alteração erro!")
	end if

Elseif acao = "delete" then
	xsql= "up_protocolo_form_delete @cd_codigo='"&cd_codigo&"'"
	dbconn.execute(xsql)
	response.write("<br>!apaga!<br>")
else
	response.write("erro")
End if


	if janela = 1 then%>
	<body onload="window.close()"></body>
<%end if


'Fim da manipulação do banco de dados
%>
<Und: <%response.write(cd_unidade&"<br>")%>
Inicio: <%response.write(cd_num_inicio&"<br>")%>
Fim: <%response.write(cd_num_fim&"<br>")%>
Data: <%response.write(dt_data_print&"<br>")%>
user: <%response.write(cd_user_inc&"<br>")%>
Páginas: <%response.write(qtd_paginas&"<br>")%>
Etiquetas: <%response.write(qtd_etiquetas&"<br>")%>
Print_ok?: <%response.write(print_ok&"<br>")%>
cd_codigo: <%=cd_codigo%><br>
-->

<%response.redirect(redirecionamento)%>