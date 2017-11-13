<!-- *** Início do formulário de ocorrências *** -->
<!--#include file="../includes/util.asp"-->
<table border="0" width="550"><tr bgcolor="#b4b4b4">
<tr bgcolor="#b4b4b4"><td colspan="100%" align="center" class="textopadrao">OCORRÊNCIAS</td></tr>
<tr bgcolor="#f5f5f5">
	<td align="center">Histórico</td></td>
	<td colspan="2"><input type="text" name="dia_ocorrencia" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="mes_ocorrencia" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="ano_ocorrencia" size="4" maxlength="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"> - 
	<input type="text" name="ocorrencia" size="52" maxlength="500" class="inputs">
	<!--<br>cd_origem--><input type="hidden" name="cd_origem" value="<%=cd_origem%>" readonly="always">
	<!--<br>cd_retorno--><input type="hidden" name="cd_retorno" value="<%=cd_codigo%>" readonly="always">
	<!--<br>id_origem--><input type="hidden" name="id_origem" value="<%=id_origem%>" readonly="always">
	<!--<br>pag_retorno--><input type="hidden" name="pag_retorno" value="<%=pag_retorno%>" readonly="always">
	<!--<br>variaveis--><input type="hidden" name="variaveis" value="<%=variaveis%>" readonly="always">
	</td>
</tr>
<tr><td><img src="../../imagens/px.gif" width="95" height="1"></td>
	<td><img src="../../imagens/px.gif" width="420" height="1"></td>
	<td><img src="../../imagens/px.gif" width="10" height="1"></td>
</tr>
<%strsql ="SELECT * From TBL_ocorrencias where id_origem='"&id_origem&"' Order by dt_ocorrencia"
  	Set rs_ocorr = dbconn.execute(strsql)
Do While not rs_ocorr.EOF
cd_cod_ocorrencia = rs_ocorr("cd_codigo")
dt_ocorrencia = rs_ocorr("dt_ocorrencia")
ocorrencia = rs_ocorr("ocorrencia")%>
<tr bgcolor="#f5f5f5">
	<td align="center">&nbsp;<%=zero(day(dt_ocorrencia))%>/<%=zero(month(dt_ocorrencia))%>/<%=year(dt_ocorrencia)%></td>
	<td><%=ocorrencia%>&nbsp;</td>
	<td align="right"><img src="<%if jan=1 then%>../<%end if%>imagens/ic_del.gif" onclick="javascript:JsDeleteocorrencia('<%=cd_cod_ocorrencia%>','<%=cd_codigo%>','<%=variaveis%>','lista<%'=tipo%>')" type="button" value="A" title="Apagar"></td>
</tr>
<%rs_ocorr.movenext
loop%>
</table>
<!-- *** Fim do formulário de ocorrências *** -->	
<SCRIPT LANGUAGE="javascript">
	function JsDeleteocorrencia(cod1, cod2, cod3){
		if (confirm ("Tem certeza que deseja excluir a ocorrência?")){
			document.location.href('ocorrencia/acoes/ocorrencias_acao.asp?cd_cod_ocorrencia='+cod1+'&cd_codigo='+cod2+'&variaveis='+cod3);}
		}
</SCRIPT>