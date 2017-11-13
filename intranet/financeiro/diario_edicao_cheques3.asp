<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../css/geral.htm"-->
<%
cd_user = session("cd_codigo")

cd_cheque = request("cd_cheque")

%>



<%'=cd_cheque%>

<body>
<table border="1" style="border:2px solid black; border-collapse:collapse;" align="center">
	<tr style="background-color:<%=cor_item1%>; font-size:16px; color:white;">
		<td align="center" colspan="8"><b><%=ucase(titulo)%> - <%=mes_selecionado(dt_mes)&"/"&dt_ano%></b></td>
	</tr>
	<tr style="background-color:<%=cor_item2%>;">
		<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center" colspan="2"><b>Conta</b><br><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td></td>
		<%if cd_tipo = 1 then%><td align="center"><b>N° Cheque</b><br><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td><%end if%>
		<td align="center"><b>Data Pagamento</b><br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
		<td align="center"><b>Valor</b><br><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td><%'=lista_pagto&" ; "&ver_virgula%></td>
	</tr>
	<%num = 1
	cor = 1
	if cd_cheque <> "" then
		'strsql = "SELECT * FROM View_financeiro_banco_pagamentos3 WHERE nm_cheque IS NOT NULL AND dt_pagamento BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"' ORDER BY nm_cheque"
		strsql = "SELECT * FROM View_financeiro_banco_pagamentos3 WHERE cd_codigo = '"&cd_cheque&"'" 'nm_cheque IS NOT NULL AND dt_pagamento BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&dia_final&"/"&dt_ano&"' ORDER BY nm_cheque"
	end if
	
	Set rs_cheques = dbconn.execute(strsql)
	
	while not rs_cheques.EOF
		'nm_empresa_abrv = rs_cheques("nm_empresa_abrv")
		'cd_cheque = rs_cheques("cd_cheque")
		cd_cheque = rs_cheques("cd_codigo")
		nm_cheque = rs_cheques("nm_cheque")
		dt_pagto = rs_cheques("dt_pagamento")
			dt_pagamento = zero(day(dt_pagto))&"/"&zero(month(dt_pagto))&"/"&year(dt_pagto)
		cd_status = rs_cheques("cd_status")
		nm_obs_cancelados = rs_cheques("nm_obs")
		'valor_cheque = rs_cheques("total")
			str_sql="SELECT * FROM TBL_financeiro_valores3 where cd_cheque='"&cd_cheque&"' order by dt_vencimento"
			Set rs_pagto = dbconn.execute(str_sql)
				while not rs_pagto.EOF
					nm_valor = rs_pagto("nm_valor")
					valor_cheque = valor_cheque + nm_valor
					'soma_valor_cheque = soma_valor_cheque + valor_cheque
				 rs_pagto.movenext
				 wend
				 'soma_valor_cheque = soma_valor_cheque + valor_cheque
			%>
		<tr style="background-color:gray; color:white; font-size:12px;">
			<td>&nbsp;<%'=cd_cheque%></td>
			<td align="center" colspan="2"><%=nm_empresa_abrv%></td>
			<%if cd_tipo = 1 then%><td align="center"><a href="javascript:void(0);" return false;" onClick="window.open('diario_edicao_cheques3.asp?cd_cheque=<%=cd_cheque%>&ano_sel=<%=dt_ano%>','cheque_edit','location=0,status=0,width=400,height=300,scrollbars=1')"><%=nm_cheque%></a></td><%end if%>
			<td align="center"><%=dt_pagamento%></td>
			<td align="right" colspan="2"><%=formatnumber(valor_cheque,2)%></td>
		</tr>
		<%if cd_status = 0 then%>
				<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
						<td><img src="../imagens/px.gif" alt="" width="1" height="20" border="0"></td>
						<td align="right" valign="middle">0</td>
						<td align="left" valign="middle" colspan="2"><b style="color:red;">CHEQUE CANCELADO</b><%if nm_obs_cancelados <> "" Then response.write(" - "&nm_obs_cancelados)%></td>						
						<td align="center"> <%=dt_pagamento%> </td>
						<td align="center"> - </td>
					</tr>
		<%elseif cd_status = 1 then
			
			'**** Lista os cheques ***
			num_linha = 1
			strsql="SELECT * FROM View_financeiro_valores3 where cd_cheque='"&cd_cheque&"' order by dt_vencimento"
			Set rs_pagto = dbconn.execute(strsql)
				while not rs_pagto.EOF
					cd_codigo = rs_pagto("cd_codigo")
					cd_cheque = rs_pagto("cd_cheque")
					cd_diario = rs_pagto("cd_diario")
					nm_fornecedor = rs_pagto("nm_fornecedor")
					nm_pagamento = rs_pagto("nm_pagamento")
					if nm_pagamento = "" Then
						nm_pagamento = rs_pagto("nm_descricao")
					end if
						
					dt_vencimento = rs_pagto("dt_vencimento")
						dt_vencimento = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&right(year(dt_vencimento),2)
						
					nm_valor = rs_pagto("nm_valor")
					soma_valores = soma_valores + nm_valor
					
					nm_empresa_abrv = rs_pagto("nm_empresa_abrv")
					cd_modo_pagto = rs_pagto("cd_modo_pagto")
					nm_cheque = rs_pagto("nm_cheque")
					
					dt_pagto = rs_pagto("dt_pagamento")
						dt_pagto = zero(day(dt_pagto))&"/"&zero(month(dt_pagto))&"/"&year(dt_pagto)%>
					<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
						<td align="right" valign="=top"><%=num_linha%></td>
						<td align="left" colspan="2" valign="=top"><%="<b>"&nm_fornecedor&"</b> - "&nm_pagamento%> <%if cd_user = 46 then%><%=" : "&cd_codigo%><%end if%></td>
						<td align="center" valign="=top"><%=dt_vencimento%></td>
						<td align="right" valign="=top" style="padding-right:5px;"><%if nm_valor > 0 then response.write(formatNumber(nm_valor))%></td>
						<!--td><img src="../imagens/ic_reprovado.gif" alt="Excluir (<%=cd_cheque%>)" width="10" height="12" border="0" onClick="javascript:JsValorDelete('<%=cd_cheque%>')"><%=cd_cheque%>--></td>
						<td align="center" valign="=top"><!--img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="JavaScript:JsPagamentoDeleteItem()"-->&nbsp;</td>
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
			
			elseif cd_status = 2 then%>
				<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
						<td><img src="../imagens/px.gif" alt="" width="1" height="20" border="0"></td>
						<td align="right" valign="middle">1</td>
						<td align="left" valign="middle" colspan="2"><b style="color:green;">PAGAMENTOS DIVERSOS</b><%if nm_obs_cancelados <> "" Then response.write(" - "&nm_obs_cancelados)%></td>						
						<td align="center"> - </td>
						<td align="center"> - </td>
					</tr>
			<%end if%>
		
		
		<tr><td colspan="6"><img src="../imagens/px.gif" alt="" width="1" height="8" border="0"></td></tr>
	<%num = num + 1
		cor = 1
		cor_linha = "#FFFFFF"
		cor_linha2 = "#e9e9e9"
		valor_cheque = 0
	rs_cheques.movenext
	wend
	num_linha = 1
	
	%>
			
			
		
		<tr><td colspan="6" bgcolor="black"><img src="../imagens/px.gif" alt="" width="1" height="3" border="0"></td></tr>
		<tr style="font-size:12px;">
			<td colspan="3">
				<%if cd_user = 46 then%>
					Apagar este cheque/Transferencia: <img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="JavaScript:JsPagamentoDeletes(<%=cd_cheque%>)">
				<%end if%>
			</td>
			<td align="center"><b>TOTAL</b></td>
			<td align="right" colspan="2" bgcolor="#d7d7d7" style="padding-right:5px;"><b><%if soma_valores > 0 then response.write(formatNumber(soma_valores))%></b></td>
		</tr>
		
		
</table><br>

<SCRIPT LANGUAGE="javascript">
	function JsPagamentoDeletes()
			   {
				if (confirm ("Tem certeza que deseja excluir este cheque?"))
			  	{
					location.href="acoes/acoes3.asp?cd_cheque="+<%=cd_cheque%>+"&acao=apaga_cheque";
					 
					//alert("teste"+<%=cd_cheque%>);
				} 	
				
			   }
				  
	function JsPagamentoDeleteItem(cod)
			   {
				if (confirm ("Tem certeza que deseja excluir este item do cheque?"+cod))
			  {
			  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&acao=apaga_item_cheque');
				}
				  }
				
	/*function JsDelete(cod)
			   {
				if (confirm ("Tem certeza que deseja excluir este cheque?"))
			  {
			  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_valor='+cod2+'&acao=apaga_item_cheque');
				}
				  }*/
</script>
</body>