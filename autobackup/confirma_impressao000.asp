<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Confirma Impressão</title>
</head>
<%cd_codigo = request("cd_codigo")
cd_unidade = request("cd_unidade")
	strsql ="TBL_unidades where cd_codigo = '"&cd_unidade&"'"
	Set rs_uni = dbconn.execute(strsql)
		if not rs_uni.EOF Then						
			nm_unidade = rs_uni("nm_Unidade")
		end if
dia = day(now)
mes = month(now)
'janela = request("janela")
%>

<body onUnload="window.opener.location.reload()">
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../../js/mascarainput.js"></SCRIPT>
<script language="JavaScript">
function validar_confirma(){
var campo = document.form.cd_confirma;
var campo_preenchido = 0;

	for(var i = 0; i < campo.length; i++){
		if (campo[i].checked){
		campo_preenchido = 1;
		}
			}
				if (campo_preenchido == 0){
					alert("Selecione uma opção.");
				return false
			}
		}	
</script>

	<table cellspacing="1" cellpadding="1" border="0" align="center" width="300">
		<form  name="form" action="../acoes/etiquetas_acao.asp" method="post" id="funcao" onsubmit="return validar_confirma(document.form);">
		<input type="hidden" name="janela" value="1">
		<input type="hidden" name="acao" value="update">
		<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
		<%strsql ="Select * from TBL_protocolo_etiquetas where cd_codigo="&cd_codigo&""
			Set rs_etq = dbconn.execute(strsql)
			if not rs_etq.eof then
			cd_num_inicio = rs_etq("cd_num_inicio")
			'cd_num_fim = rs_etq("cd_num_fim")
			'else cd_num_fim = 0
			end if%>
		<tr><td colspan="100%" align="center" bgcolor="#808080" style="color:white;">CONFIRMAÇÃO</td></tr>
		<tr bgcolor="#e2e2e2"">
		    <td>&nbsp;Unidade:</td>
			<td>&nbsp; <%=nm_unidade%> <%'=cd_codigo&":"&cd_unidade%></td>
		</tr>
		<tr bgcolor="#dfdfdf">
		    <td colspan="2">&nbsp;</td>
		</tr>
		<tr bgcolor="#dfdfdf"">
		    <td align="right"><input type="radio" name="cd_confirma" value="1"></td>
			<td>&nbsp;Impressão ok.</td>
		</tr>
		<tr bgcolor="#e4e4e4"">
		    <td align="right"><input type="radio" name="cd_confirma" value="2"></td>
			<td>&nbsp;Impressão parcial.</td>
		</tr>
		<tr bgcolor="#e4e4e4"">
		    <td>&nbsp;</td>    
			<!--td>&nbsp;inicio: &nbsp;<%=zero(cd_unidade)%>.<input type="text" name="cd_num_inicio" size="6" maxlength="6" class="inputs" value="<%=cd_num_inicio%>">-X</td-->
			<td>&nbsp;inicio: &nbsp;<%=cdun(cd_unidade)%>.<%=proto(trim(cd_num_inicio))%>-X <input type="hidden" name="cd_num_inicio" class="inputs" value="<%=cd_num_inicio%>"></td>
			<!--td>&nbsp;inicio: &nbsp;<%'=digitov(zero(cd_unidade)&"."&proto(cd_num_inicio))%><input type="hidden" name="cd_num_inicio" value="<%=cd_num_inicio%>"></td-->
		</tr>
		<tr bgcolor="#e4e4e4"">
		    <td>&nbsp;</td>    
			<td>&nbsp;fim: &nbsp;&nbsp;&nbsp; <%=zero(cd_unidade)%>.<input type="text" name="cd_num_fim" size="6" maxlength="6" class="inputs" value="">-X</td>	
		</tr>
		<tr bgcolor="#e4e4e4"">
		    <td align="right"><input type="radio" name="cd_confirma" value="3" onclick="javascript:JsDelete('<%=cd_codigo%>')"></td>
			<td>&nbsp;Apagar registro.</td>
		</tr>
		<tr bgcolor="#e4e4e4"">
		    <td align="center" colspan="2"><br><input type="submit" value="OK"><br>&nbsp;</td>
		</tr>
		</form>
	</table>



	<SCRIPT LANGUAGE="javascript">
	function JsDelete(cod)
	{
		if (confirm ("Tem certeza que deseja excluir?"))
			{
				document.location=('../acoes/etiquetas_acao.asp?cd_codigo='+cod+'&janela=1&acao=delete');
				//document.location=('protocolo/acoes/protocolo_acao.asp?cd_codigo='+cod1+'&no_protocolo='+cod2+'&cd_unidade='+cod3+'&no_digito='+cod4+'&acao=excluir&obj_excl=6');
			}
	}
	</SCRIPT>
	
							
</body>
</html>