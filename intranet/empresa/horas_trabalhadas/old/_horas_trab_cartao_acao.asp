<!--#include file="../../includes/inc_open_connection.asp"-->
   <!--include file="../includes/inc_area_restrita.asp"-->
<%
cd_codigo = request("cd_codigo")
cd_cod_horas = request("cd_cod_horas")

diaentrada = request("diaentrada")
mesentrada = request("mesentrada")
anoentrada = request("anoentrada")
horaentrada_hr = request("horaentrada_hr")
horaentrada_min = request("horaentrada_min")
data_entrada = mesentrada&"/"&diaentrada&"/"&anoentrada&" "&horaentrada_hr&":"&horaentrada_min

diasaida = request("diasaida")
messaida = request("messaida")
anosaida = request("anosaida")
horasaida_hr = request("horasaida_hr")
horasaida_min = request("horasaida_min")
data_saida = messaida&"/"&diasaida&"/"&anosaida&" "&horasaida_hr&":"&horasaida_min

cd_status = request("cd_status")
cd_unidades = request("cd_unidades")
cd_funcao = request("cd_funcao")
acao = request("acao") 

If acao = "inserir" Then
xsql ="up_cartao_horas_insert @cd_codigo="&cd_codigo&",@data_entrada='"&data_entrada&"', @data_saida='"&data_saida&"',@cd_status="&cd_status&",@cd_unidades='"&cd_unidades&"',@cd_funcao='"&cd_funcao&"'"
dbconn.execute(xsql)
response.write("OK, inserido")
response.redirect "../horas_trabalhadas/horas_trab_cartao_acao.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""

Elseif acao = "editar" Then
xsql ="up_cartao_horas_update @cd_funcionario="&cd_codigo&",@cd_cod_horas="&cd_cod_horas&",@data_entrada='"&data_entrada&"', @data_saida='"&data_saida&"'"',@cd_status="&cd_status&",@cd_unidades='"&cd_unidades&"',@cd_funcao='"&cd_funcao&"'"
dbconn.execute(xsql)
'response.write("OK, alterado")
response.redirect "../produtividade.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""

ElseIf acao = "excluir" Then
xsql ="up_cartao_horas_excluir @cd_codigo='"&cd_codigo&"',@cd_cod_horas='"&cd_cod_horas&"'"
dbconn.execute xsql
'response.write("OK, excluído")
response.redirect "../horas_trabalhadas/horas_trab_cartao_acao.asp?mes_sel="&mesentrada&"&ano_sel="&anoentrada&"&cd_codigo="&cd_codigo&""

Else

response.write("Erro, não passou")
End if
%>

<br>
código: <%=cd_codigo%><br>
entrada: <%=data_entrada%><br>
saida: <%=data_saida%><br>
cd_status: <%=cd_status%><br>
cd_unidades: <%=cd_unidades%><br>
cd_funcao: <%=cd_funcao%><br>
acao: <%=acao%><br>
cd_hora: <%=cd_cod_horas%>