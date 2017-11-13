<!--#include file="../../includes/inc_open_connection.asp"-->

<%'Dados da ocorrencia
cd_cod_ocorrencia = request("cd_cod_ocorrencia")
cd_origem = request("cd_origem")
id_origem = request("id_origem")
ocorrencia = request("ocorrencia")
dia_ocorrencia = request("dia_ocorrencia")
mes_ocorrencia = request("mes_ocorrencia")
ano_ocorrencia = request("ano_ocorrencia")
	dt_ocorrencia = mes_ocorrencia&"/"&dia_ocorrencia&"/"&ano_ocorrencia
cd_retorno = request("cd_retorno")
pag_retorno = request("pag_retorno")
tipo = request("tipo")
variaveis = request("variaveis")
	variaveis = replace(variaveis,"$","&")

'******************************** Registro de ocorrencias *******************************************
' Registra ocorrências 
if ocorrencia <> "" then
xsql = "up_ocorrencia_insert @cd_origem='"&cd_origem&"', @id_origem='"&id_origem&"', @dt_ocorrencia='"&dt_ocorrencia&"', @ocorrencia='"&ocorrencia&"'"
dbconn.execute(xsql)
	if cd_codigo <> "" Then
	'response.redirect("../manutencao.asp?tipo=andamento&cd_codigo="&cd_retorno&"&campo=cd_codigo"&variaveis)
	'response.redirect("../manutencao.asp")
	'response.write("../manutencao.asp?tipo=andamento&cd_codigo="&cd_retorno&"&campo=cd_codigo"&variaveis)
	
	end if
'continua a operação com a O.S.
end if

' Apaga ocorrências
if cd_cod_ocorrencia <> "" Then
xsql = "up_ocorrencia_delete @cd_cod_ocorrencia='"&cd_cod_ocorrencia&"'"
dbconn.execute(xsql)
	'Retorna para a origem
	'response.redirect("../manutencao.asp?tipo=andamento&cd_codigo="&cd_codigo&"&campo=cd_codigo")
	'response.redirect("../"&pag_retorno&"?tipo="&tipo&"&cd_codigo="&cd_codigo&"&campo=cd_codigo")
	response.redirect(variaveis)
end if


%>

variaveis : <%=variaveis%><br><br>

pag_retorno: <%=pag_retorno%><br>
cd_retorno: <%=cd_retorno%><br>
tipo: <%=tipo%>

