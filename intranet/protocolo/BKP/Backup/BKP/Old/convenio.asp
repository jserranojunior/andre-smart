
<!-- faz a conexão com o banco -->
<!--#include file="../../includes/inc_open_connection.asp"-->
<%
'resgata o que foi digitado no campo de busca
cd_convenio = request("cd_convenio")

'txtBusca = "a"

'aqui você pode fazer dois tipos de select's. Sendo um trazendo um nome especifico ou trazendo todos que contenham o que foi digitado
'no campo de busca.
'xsql = "SELECT * FROM TBL_equipamento WHERE nm_equipamento = '" & txtBusca & "' "
xsql = "SELECT * FROM TBL_convenios WHERE cd_codigo = '"&cd_convenio&"'"
Set objRS = dbconn.execute(xsql)
%>

<!-- começamos a montagem da nossa lista -->
<select size="6" style="width:315" id="lista" name="lista">
<%
'verifica e/ou varremos todas as linhas do banco, para preencher a lista com todos os registros.
'Depois trazemos todos os registros da coluna nome ou apenas o nome digitado no campo de busca. Como pode ser visto abaixo.
while not objRS.EOF
%>
<option value="0"><%=objRS("nm_convenio")%></option>
<%
 'partimos para o próximo registro. Até chegar o último e terminar o loop.
 objRS.MoveNext
  wend

  'fechamos e esvaziamos nosso recordset   
  objRS.close
  set objRS = nothing   
%>
</select>

