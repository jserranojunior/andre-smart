<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->

<%cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)

cd_conta_banco = request("cd_conta_banco")
nm_cheque = request("nm_cheque")
cd_cheque = request("cd_cheque")

lista_pagto = request("lista_pagto")
	ver_virgula = instr(1,lista_pagto,",",1)
		if ver_virgula = 1 then lista_pagto = mid(lista_pagto,2,len(lista_pagto))
		
cor_item1 = "#afaf61"
cor_item2 = "#eaeaea"%>


<body onunload="window.opener.location.reload(true);">

<table border="1" style="border:2px solid black; border-collapse:collapse;">
	<tr style="background-color:<%=cor_item1%>; font-size:16px; color:white;">
		<td align="center" colspan="8"><b>EMISSÃO DE CHEQUES</b></td>
	</tr>
	<tr style="background-color:<%=cor_item2%>;">
		<td>&nbsp;</td>
		<td align="center" colspan="3"><b>Descrição</b></td>
		<td align="center"><b>Venc.</b></td>
		<td align="center"><b>Valor</b></td>
	</tr>
	<tr>
		<td><%'=lista_pagto&" ; "&ver_virgula%></td>
	</tr>
	<%num_linha = 1
	cor_linha = "#FFFFFF"
	cor_linha2 = "#e9e9e9"
	
	if cd_conta_banco = "" AND nm_cheque = "" AND lista_pagto <> "" Then
	
		pagto_array = split(lista_pagto,",")
		For Each pagto_item In pagto_array
			strsql="SELECT * FROM View_financeiro_valores3 where cd_codigo = '"&pagto_item&"'"
			Set rs_pagto = dbconn.execute(strsql)
			while not rs_pagto.EOF
				nm_fornecedor = rs_pagto("nm_fornecedor")
				nm_pagamento = rs_pagto("nm_pagamento")
				dt_vencimento = rs_pagto("dt_vencimento")
					dt_vencimento = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&right(year(dt_vencimento),2)
					
				nm_valor = rs_pagto("nm_valor")
				soma_valores = soma_valores + nm_valor%>
				<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
					<td align="right"><%=num_linha%></td>
					<td align="left" colspan="3"><%="<b>"&nm_fornecedor&"</b> - "&nm_pagamento%></td>
					<td align="center"><%=dt_vencimento%></td>
					<td align="right"><%if nm_valor > 0 then response.write(formatNumber(nm_valor))%></td>
				</tr>
			<%num_linha = num_linha + 1
			if cor > 0 then
				cor_linha = "#d7d7d7"
				cor_linha2 = "#d7d7d7"
				cor = 0
			else
				cor_linha = "#FFFFFF"
				cor_linha2 = "#e9e9e9"
				cor = 1
			end if
			rs_pagto.movenext
			wend
		next
	elseif cd_conta_banco <> "" AND nm_cheque <> "" AND lista_pagto = "" Then
		strsql="SELECT * FROM View_financeiro_valores3 where cd_conta_banco='"&cd_conta_banco&"' AND nm_cheque='"&nm_cheque&"'"
		Set rs_pagto = dbconn.execute(strsql)
			while not rs_pagto.EOF
				cd_cheque = rs_pagto("cd_cheque")
				nm_fornecedor = rs_pagto("nm_fornecedor")
				nm_pagamento = rs_pagto("nm_pagamento")
				dt_vencimento = rs_pagto("dt_vencimento")
					dt_vencimento = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&right(year(dt_vencimento),2)
					
				nm_valor = rs_pagto("nm_valor")
				soma_valores = soma_valores + nm_valor
				
				nm_empresa_abrv = rs_pagto("nm_empresa_abrv")
				nm_cheque = rs_pagto("nm_cheque")
				dt_pagto = rs_pagto("dt_pagto")
					dt_pagto = zero(day(dt_pagto))&"/"&zero(month(dt_pagto))&"/"&year(dt_pagto)%>
				<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
					<td align="right"><%=num_linha%></td>
					<td align="left" colspan="3"><%="<b>"&nm_fornecedor&"</b> - "&nm_pagamento%></td>
					<td align="center"><%=dt_vencimento%></td>
					<td align="right"><%if nm_valor > 0 then response.write(formatNumber(nm_valor))%></td>
					<td><img src="../imagens/ic_reprovado.gif" alt="Excluir (<%=cd_cheque%>)" width="10" height="12" border="0" onClick="javascript:JsValorDelete('<%=cd_cheque%>')"><%=cd_cheque%></td>
				</tr>
			<%num_linha = num_linha + 1
			if cor > 0 then
				cor_linha = "#d7d7d7"
				cor_linha2 = "#d7d7d7"
				cor = 0
			else
				cor_linha = "#FFFFFF"
				cor_linha2 = "#e9e9e9"
				cor = 1
			end if
			rs_pagto.movenext
			wend
	end if%>
		<tr>
			<td colspan="5"></td>
			<td align="right"><b><%if soma_valores > 0 then response.write(formatNumber(soma_valores))%></b></td>
		</tr>
		<%If cd_conta_banco = "" AND nm_cheque = "" AND lista_pagto <> "" Then%><br>
		<tr><td colspan="6">&nbsp;</td></tr>
		<tr style="background-color:<%=cor_item2%>;">
			<td></td>
			<td align="center"><b>Conta</b></td>
			<td align="center"><b>N° Cheque</b></td>
			<td align="center"><b>Data Pgto</b></td>
			<td></td>
			<td align="center">&nbsp;<!-- <b>Confirma</b>--></td>
		</tr>		
		<form action="acoes/acoes3.asp" method="post" name="cheques" id="cheques">
		<input type="hidden" name="acao" value="insere_cheque">
		<input type="hidden" name="lista_pagto" value="<%=lista_pagto%>">
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		
		<tr>
			<td></td>
			<td><select name="cd_conta_banco">
					<option value=""></option>
					<%strsql = "SELECT * FROM View_financeiro_banco where cd_status = 1"
					Set rs_banco = dbconn.execute(strsql)
						while not rs_banco.EOF
							nm_empresa = rs_banco("nm_empresa_abrv")
							cd_conta_bancaria = rs_banco("cd_codigo")
							cd_empresa = rs_banco("cd_empresa")
						rs_banco.movenext%>
						<option value="<%=cd_conta_bancaria%>" <%if cd_conta_bancaria = nm_conta_banco then response.write("SELECTED")%>><%=nm_empresa%></option>
						<%wend%>
				</select>
			</td>
			<td align="center"><input type="text" name="nm_cheque" maxlength="10" size="10"></td>
			<td align="center"><input type="text" name="dia_pagamento" value="<%=dia_pagamento%>" size="2" maxlength="2">
				<input type="text" name="mes_pagamento" value="<%=mes_pagamento%>" size="2" maxlength="2">
				<input type="text" name="ano_pagamento" value="<%=ano_pagamento%>" size="4" maxlength="4"></td>
			<td align="center"><input type="submit" name="Ok" value="OK"></td>
			<td align="center"><input type="checkbox" name="cd_confirma_pagto" value="1" ></td>
			
		</tr>
		</form>
		<%Elseif cd_conta_banco <> "" AND nm_cheque <> "" AND lista_pagto = "" Then%><br>
		<tr><td colspan="6">&nbsp;</td></tr>
		<tr>
			<td></td>
			<td align="center"><b>Conta</b></td>
			<td align="center"><b>N° Cheque</b></td>
			<td align="center"><b>Data Pgto</b></td>
			<td></td>
			<td align="center"><!--<b>Confirma--></b></td>
		</tr>		
		<tr>
			<td></td>
			<td align="center"><%=nm_empresa_abrv%></td>
			<td align="center"><%=nm_cheque%></td>
			<td align="center"><%=dt_pagto%></td>
			<td align="center"></td>
			<td align="center"></td>
			
		</tr>
		<%end if%>
		<tr>
			<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
			<td colspan="3"><img src="../imagens/blackdot.gif" alt="" width="450" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="70" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="90" height="1" border="0"></td>
			
		</tr>	
</table>
<SCRIPT LANGUAGE="javascript">
	function JsPagamentoDelete(cod)
			   {
				if (confirm ("Tem certeza que deseja excluir esse item do cheque?"))
			  {
			  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_valor='+cod2+'&acao=apaga_item_cheque');
				}
				  }
</script>
</body>