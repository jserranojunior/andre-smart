<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<%
cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)

dt_mes = request("mes_sel")
dt_ano = request("ano_sel")
%>
<html>
<head>
	<title>Financeiro - Diário</title>
</head>
<!--#include file="../js/ajax.js"-->
<!--#include file="../ferramentas/js/ferramentas.js"-->
<body onLoad="diario.cd_fornecedor.focus()">


<table border="1" align="center">
	<form action="../financeiro/acoes/acoes.asp" name="diario" id="diario" method="post">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="data_atual" value="<%=data_atual%>">
	<tr>
		<td colspan="2">Cadastramento de pagamentos</td>		
	</tr>

	<tr>
		<td>Favorecido</td>
		<td><span id="divForn"><%strsql = "Select * From TBL_fornecedor order by nm_fornecedor"
				Set rs_fornec = dbconn.execute(strsql)%>
				<select name="cd_fornecedor" class="inputs">
					<option value=""></option>
					<%While Not rs_fornec.eof
						cd_fornec = rs_fornec("cd_codigo")
						nm_fornec = rs_fornec("nm_fornecedor")%>	
						<option value="<%=cd_fornec%>" <%if int(strcd_fornecedor) = cd_fornec then%>Selected<%end if%>><%=nm_fornec%></option>
						<%rs_fornec.movenext
					Wend
						rs_fornec.close
						Set rs_fornec = nothing%>
				</select>
				</span>
				<img src="../imagens/reload6.gif" alt="Atualiza listagem" width="10" height="12" border="0" onClick="reload_fornecedor();">&nbsp;
				<img src="../imagens/add1.gif" alt="incluir fornecedor" width="10" height="12" border="0" onClick="window.open('../ferramentas/fornecedores_cad.asp','itens_nomes','width=500,height=210')" return false;">
							
		</td>
	</tr>
	<tr>
		<td>Item</td>
		<td><input type="text" name="nm_pagamento" value="<%=nm_despesa%>" size="40" maxlength="50"></td>
	</tr>
	<tr>
		<td>Data</td>
		<td><input type="text" name="dia_vencimento" value="<%'=day(dt_vencimento)%>" size="2" maxlength="2">
		<input type="hiddens" name="mes_vencimento" value="<%=dt_mes%>" size="2" maxlength="2">
		<input type="hiddens" name="ano_vencimento" value="<%=dt_ano%>" size="4" maxlength="4"></td>
	</tr>
	<tr>
		<td>Tipo</td>
		<td><%strsql = "Select * From TBL_financeiro_tipo_pagamento order by nm_tipo_pagamento"
				Set rs_tipopag = dbconn.execute(strsql)%>
				<select name="cd_tipo_pagamento">
					<option value="" SELECTED></option>
					<%While not rs_tipopag.EOF
						cd_tipo_pagamento = rs_tipopag("cd_codigo")
						nm_tipo_pagamento = rs_tipopag("nm_tipo_pagamento")%>
						<option value="<%=cd_tipo_pagamento%>"><%=nm_tipo_pagamento%></option>
					<%rs_tipopag.movenext
					wend
					
					rs_tipopag.close
					Set rs_tipopag = nothing%>
				</select></td>
	</tr>
	<tr>
		<td>Valor</td>
		<td><input type="text" name="nm_valor" value="<%=nm_valor%>" size="20" maxlength="11"> &nbsp; &nbsp;
		Qtd. Parcelas: <input type="text" name="cd_parcela" value="<%=cd_parcela%>" size="2" maxlength="2"></td>
	</tr>
	<tr>
		<td colspan="2"><img src="../imagens/blackdot.gif" alt="" width="410" height="1" border="0"></td>
	</tr>
	<tr>
		<td>Área</td>
		<td><%strsql = "Select * From TBL_financeiro_area order by nm_area"
				Set rs_area = dbconn.execute(strsql)%>
				<select name="cd_area" class="inputs">
					<option value=""></option>
					<%Do While Not rs_area.eof
						cd_area = rs_area("cd_codigo")
						nm_area = rs_area("nm_area")%>	
						<option value="<%=cd_area%>" <%if int(strcd_area) = cd_area then%>Selected<%end if%>><%=nm_area%></option>
						<%rs_area.movenext
					loop
						rs_area.close
						Set rs_area = nothing%>
				</select>		
		</td>
	</tr>
	<tr>
		<td>C.Custos</td>
		<td><%strsql = "Select * From TBL_financeiro_centrocusto order by nm_centrocusto"
				Set rs_cc = dbconn.execute(strsql)%>
				<select name="cd_centrocusto" class="inputs">
					<option value=""></option>
					<%Do While Not rs_cc.eof
						cd_centrocusto = rs_cc("cd_codigo")
						nm_centrocusto = rs_cc("nm_centrocusto")%>	
						<option value="<%=cd_centrocusto%>" <%if int(strcd_centrocusto) = cd_centrocusto then%>Selected<%end if%>><%=nm_centrocusto%></option>
						<%rs_cc.movenext
					loop
						rs_cc.close
						Set rs_cc = nothing%>
				</select>		
		</td>
	</tr>
	<tr>
		<td>Conta</td>
		<td><%strsql = "Select * From TBL_financeiro_conta order by nm_conta"
				Set rs_conta = dbconn.execute(strsql)%>
				<select name="cd_conta" class="inputs">
					<option value=""></option>
					<%Do While Not rs_conta.eof
						cd_conta = rs_conta("cd_codigo")
						nm_conta = rs_conta("nm_conta")%>	
						<option value="<%=cd_conta%>" <%if int(strcd_conta) = cd_conta then%>Selected<%end if%>><%=nm_conta%></option>
						<%rs_conta.movenext
					loop
						rs_conta.close
						Set rs_conta = nothing%>
				</select>		
		</td>
	</tr>
	<tr>
		<td>Unidade</td>
		<td><%strsql = "Select * From TBL_unidades where cd_status=1 and cd_hospital >= 1 order by nm_unidade"
				Set rs_unid = dbconn.execute(strsql)%>
				<select name="cd_unidade" class="inputs" multiple>
					<option value=""></option>
					<%Do While Not rs_unid.eof
						cd_unidade = rs_unid("cd_codigo")
						nm_unidade = rs_unid("nm_unidade")
						nm_sigla = rs_unid("nm_sigla")%>	
						<option value="<%=cd_unidade%>" <%if int(strcd_unidade) = cd_unidade then%>Selected<%end if%>><%=nm_sigla%></option>
						<%rs_unid.movenext
					loop
						rs_unid.close
						Set rs_unid = nothing%>
				</select>		
		</td>
	</tr>
	<tr><td align="center" colspan="2"><input type="submit" name="incluir" value="incluir"> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <input type="submit" name="Finalizar" value="Fechar"></td></tr>
	</form>
</table>



</body>
</html>
