<!-- *** Início do formulário de ocorrências *** -->
<!--#include file="../includes/util.asp"-->
<table border="0" width="550"><tr bgcolor="#b4b4b4">


<tr><td><img src="imagens/px.gif" width="95" height="1"></td>
	<td><img src="imagens/px.gif" width="420" height="1"></td>
	<td><img src="imagens/px.gif" width="10" height="1"></td>
</tr>
<%strsql ="SELECT * From TBL_ocorrencias where id_origem='"&id_origem&"' Order by dt_ocorrencia"
  	Set rs_ocorr = dbconn.execute(strsql)
	
if not rs_ocorr.EOF then%>
	<tr bgcolor="#b4b4b4"><td colspan="100%" align="center" class="textopadrao">OCORRÊNCIAS</td></tr>
<%else%>
	<tr bgcolor="#b4b4b4"><td colspan="100%" align="center" class="textopadrao">NÃO EXISTEM OCORRÊNCIAS</td></tr>
<%end if

Do While not rs_ocorr.EOF
cd_cod_ocorrencia = rs_ocorr("cd_codigo")
dt_ocorrencia = rs_ocorr("dt_ocorrencia")
ocorrencia = rs_ocorr("ocorrencia")%>
<tr bgcolor="#f5f5f5">
	<td align="center">&nbsp;<%=zero(day(dt_ocorrencia))%>/<%=zero(month(dt_ocorrencia))%>/<%=year(dt_ocorrencia)%></td>
	<td><%=ocorrencia%>&nbsp;</td>
	<td align="right">&nbsp;<!--img src="imagens/ic_del.gif" onclick="javascript:JsDeleteocorrencia('<%=cd_cod_ocorrencia%>','<%=cd_codigo%>','<%=variaveis%>','lista<%'=tipo%>')" type="button" value="A" title="Apagar"--></td>
</tr>
<%rs_ocorr.movenext
loop%>
</table>
