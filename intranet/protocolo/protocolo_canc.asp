		<%if cd_cod_protocolo = "" then
			acao = "inserir"
		else
			acao = "editar"
		end if
		outros = "&modalidade=cancelar"
		cod_confirmar = request("cod_confirmar")
		%>
		
		<tr bgcolor="#f5f5f5">
			<td>&nbsp;Protocolo:</td>
			<td colspan="2">&nbsp;<b><%=digitov(zero(int(cd_unidade))&"."&proto(no_protocolo))%></b></td>
		</tr>
		<tr>
			<td>&nbsp;Usuário:</td>
			<td colspan="2">&nbsp;<b><%=session("nm_nome")%></b></td>
		</tr>
		<tr bgcolor="#f5f5f5">
			<td>&nbsp;Data:</td>
			<td colspan="2">&nbsp;<b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
		</tr>
		
		<!--
		<tr bgcolor="#808080">
			<td align="center" colspan="100%"><font color="white"><b>&nbsp;Motivo do cancelamento</b></font></td>
		</tr>
		<tr>
			<td align="right"><input type="radio" name="motivo_cancelamento" value="1"></td>
			<td colspan="2">&nbsp;Rasura</td>
		</tr>
		<tr>
			<td align="right"><input type="radio" name="motivo_cancelamento" value="2"></td>
			<td colspan="2">&nbsp;Cancelamento do procedimento</td>			
		</tr>
		<tr>
			<td align="right">Obs:</td>
			<td colspan="2"><textarea cols="30" rows="3" name="obs_canc">&nbsp;Cancelamento do procedimento</textarea></td>			
		</tr>
		-->
		<tr>
			<td colspan="100%">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="100%">&nbsp;<input type="button" name="cancela" value="cancela" onclick="javascript:JsProcCancela('<%=int(cd_unidade)%>','<%=int(no_protocolo)%>','<%=session_cd_usuario%>')" ></td>
		</tr>
		<%if cod_confirmar = 1 then%>
		<tr><td colspan="100%" align="center"><b style="color:red;"><%=mensagem%></b></td></tr>
		<tr>
			<td colspan="100%">&nbsp;<input type="checkbox" name="autoriza_cancelar" value="1"> Marque para autorizar o cancelamento do protocolo já digitado.</td>
		</tr>
		<%end if%>
		
		<!--******************************************************* Fim da inclusao das Oticas  ***********************************************************************-->

<SCRIPT LANGUAGE="javascript">
	function JsProcCancela(cod1,cod2,cod3)
	{
	  if (confirm ("Tem certeza que deseja invalidar o formulário?"))
		  {
			document.location=('protocolo/acoes/protocolo_acao.asp?cd_unidade='+cod1+'&no_protocolo='+cod2+'&cd_usuario='+cod3+'&acao=cancelar');
		  
		  }
	}

</SCRIPT>