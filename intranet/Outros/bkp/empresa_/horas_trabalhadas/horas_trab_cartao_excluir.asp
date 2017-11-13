<html>
<%cd_cod_horas = request("cd_cod_horas")
cd_cod_abono = request("cd_cod_abono")
%>
<body onUnload="window.opener.location.reload()">
<table align="center">
	<tr>
		<td>Confirma a exclusão?</td>
	</tr>
	<tr><td>&nbsp;</td></tr>
	<form action="horas_trab_cartao_acao.asp" name="conf_excl" id="conf_excl">
	<input type="hidden" name="cd_codigo" value="97">
	<input type="hidden" name="cd_cod_horas" value="<%=cd_cod_horas%>">
	<input type="hidden" name="cd_cod_abono" value="<%=cd_cod_abono%>">
	<input type="hidden" name="mesentrada" value="8">
	<input type="hidden" name="anoentrada" value="2008">
	<input type="hidden" name="menos_min" value="600">
	<input type="hidden" name="cd_excedente" value="-300">
	<input type="hidden" name="cd_minutos_trab" value="600">
	<input type="hidden" name="tela" value="0">
	<input type="hidden" name="acao" value="excluir">
	<tr>
			<td><%if cd_cod_horas <> "" AND cd_cod_abono = "" Then 'exclui as horas trabalhadas%>
			<input type="hidden" name="cd_h_t_exclui" value="1">
			<%elseif cd_cod_horas = "" AND cd_cod_abono <> "" Then 'exclui o abono%>
			<input type="hidden" name="cd_abono_exclui" value="1">
			<%Elseif cd_cod_horas <> "" AND cd_cod_abono <> "" Then 'Pergunta quais dados excluir (horas trabalhadas ou abono%>
			<input type="checkbox" name="cd_h_t_exclui" value="1">Excluir as horas trabalhadas<br>
			<input type="checkbox" name="cd_abono_exclui" value="1">Excluir abono
			<%end if%>
		</td>
	</tr>
	<tr>
		<td><br>
		<input type="submit" value="Ok">
		<input type="reset" name="cancela" value="Cancela" onClick="JavaScript:window.close()"></td>
	</tr>
	</form>
</table>
<br>




<br>


</body>
</html>