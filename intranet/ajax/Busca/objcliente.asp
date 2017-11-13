<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<!-- faz a conexão com o banco -->
<!--#include file="../../includes/inc_open_connection.asp"-->
<%
'resgata o que foi digitado no campo de busca
txtBusca = request("txtBusca")%>

<table>
<tr bgcolor=#c0c0c0>
	<td valign="top" align=center width="10%" class="textopadrao">		
		<b>CÓD.</b>
	</td>
	<td valign="top" width="30%" class="textopadrao">
		<a href="empresa.asp?tipo=lista&ordem="><b>EMPRESA</b></a>
	</td>
	<td valign="top" width="20%" class="textopadrao">
		<b>CNPJ</b>
	</td>
 </tr>
<tr>
	<td><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
	<td><img src="../../imagens/px.gif" alt="" width="450" height="1" border="0"></td>
	<td><img src="../../imagens/px.gif" alt="" width="300" height="1" border="0"></td>	
</tr>

<%
if txtBusca = "todos" Then
xsql = "SELECT * FROM TBL_clientes order by nm_empresa"
Else
xsql = "SELECT * FROM TBL_clientes WHERE nm_empresa LIKE '"&txtBusca&"%' ORDER by nm_empresa"
End if

if txtBusca <> "" Then


'txtBusca = "a"

'aqui você pode fazer dois tipos de select's. Sendo um trazendo um nome especifico ou trazendo todos que contenham o que foi digitado
'no campo de busca.
'xsql = "SELECT * FROM TBL_clientes WHERE nm_empresa LIKE '"&txtBusca&"%' "
Set objRS = dbconn.execute(xsql)

%>

<!-- começamos a montagem da nossa lista -->
<!--select size="6" style="width:315" id="lista" name="lista"-->
<%
if objRS.EOF then%>
<tr><td colspan="100%">Nenhum registro foi encontrado.</td></tr>
<%end if%>


<%while not objRS.EOF
cd_cliente = objRS("cd_codigo")
nm_empresa = objRS("nm_empresa")
cd_cnpj = objRS("cd_cnpj")
%>


<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('clientes.asp?tipo=cadastro&cd_cliente=<%=cd_cliente%>');" onmouseout="mOut(this,'');">
	<td valign="top" align=center>					
		<font color="<%'=strcor%>"><%=cd_cliente%>&nbsp;&nbsp;&nbsp;&nbsp;</font>	
	</td>
	<td valign="top" >					
		<a href="clientes.asp?tipo=cadastro&cd_cliente=<%=cd_cliente%>"><%=nm_empresa%></a>
		<!--font color="<%'=strcor%>"><%'=nm_empresa%></font-->
	</td>
	<td valign="top">
		<font color="<%=strcor%>"><%=cd_cnpj%></font>
	</td>	
</tr>
<!--/option-->
<%
 'partimos para o próximo registro. Até chegar o último e terminar o loop.
 objRS.MoveNext
  wend


  'fechamos e esvaziamos nosso recordset   
  objRS.close
  set objRS = nothing
%>

<%end if%>
</table>
<!--/select-->

