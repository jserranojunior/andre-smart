<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->

<%cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)

cd_conta_banco = request("cd_conta_banco")
nm_cheque = request("nm_cheque")
'cd_cheque = request("cd_cheque")
cd_pagamento = request("cd_pagamento")

lista_pagto = request("lista_pagto")
	ver_virgula = instr(1,lista_pagto,",",1)
		if ver_virgula = 1 then lista_pagto = mid(lista_pagto,2,len(lista_pagto))
		
cor_item1 = "#afaf61"
cor_item2 = "#eaeaea"

escolha_documentos = request("escolha_documentos")
	if escolha_documentos = "" Then 
		'escolha_documentos = "none"
	end if
dia_sel = request("dia_pagto")
mes_sel = request("mes_pagto")
ano_sel = request("ano_pagto")

nm_valor_cheque = request("nm_valor_cheque")
cd_status = request("cd_status")
nm_obs = request("nm_obs")

mens = request("mens")
	mens = replace(mens,"_","&nbsp;")%>

<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "modo_pagamento"; // nome do primeiro campo do site
netscape = "";
ver = navigator.appVersion; len = ver.length;
for(iln = 0; iln < len; iln++) if (ver.charAt(iln) == "(") break;
	netscape = (ver.charAt(iln+1).toUpperCase() != "C");
	function keyDown(DnEvents) {
		// ve quando e o netscape ou IE
		k = (netscape) ? DnEvents.which : window.event.keyCode;
		if (k == 13) { // preciona tecla enter
		if (nextfield == 'done') {
			//alert("viu como funciona?");
			eval('document.cheques.' + nextfield + '.focus()');
			return false;
			//return true; // envia quando termina os campos
		} else {
			// se existem mais campos vai para o proximo
			eval('document.cheques.' + nextfield + '.focus()');
			return false;
		}
	}
}
document.onkeydown = keyDown; // work together to analyze keystrokes
if (netscape) document.captureEvents(Event.KEYDOWN|Event.KEYUP);


function validar_pagto(shipinfo){
	
	hoje = new Date
	dia_hoje = hoje.getDate();
		if (dia_hoje < 10 ) dia_hoje = '0'+ dia_hoje
	mes_hoje = hoje.getMonth()+1;
		if (mes_hoje < 10 ) mes_hoje = '0'+ mes_hoje
	ano_hoje = hoje.getFullYear();
		data_hoje = ano_hoje +''+ (mes_hoje) +''+ dia_hoje;
	 	data_informada = shipinfo.ano_pagamento.value +''+ shipinfo.mes_pagamento.value +''+ shipinfo.dia_pagamento.value;
	pgto_atrasado = shipinfo.pgto_atrasado.value;
	 
	if (document.getElementById('cheque').checked == false && document.getElementById('transf').checked == false){window.alert ("Selecione o tipo de pagamento");document.getElementById('cheque').focus(); return (false);}
	if (document.getElementById('cheque').checked == true && document.getElementById('conta').value == ""){window.alert ("Selecione a conta");document.getElementById('conta').focus(); return (false);}
	if (document.getElementById('cheque').checked == true && shipinfo.nm_cheque.value == ""){window.alert ("informe o no. do cheque");shipinfo.nm_cheque.focus(); return (false);}
	
	if (shipinfo.dia_pagamento.value.length==""){window.alert ("A data deve ser preenchida.");shipinfo.dia_pagamento.focus(); return (false);}
	if (shipinfo.mes_pagamento.value.length==""){window.alert ("A data deve ser preenchida.");shipinfo.mes_pagamento.focus(); return (false);}
	if (shipinfo.ano_pagamento.value.length==""){window.alert ("A data deve ser preenchida.");shipinfo.ano_pagamento.focus(); return (false);}
	
	//if (data_informada<data_hoje){window.alert ("A data preenchida é menor que a atual.");shipinfo.ano_pagamento.value="";shipinfo.mes_pagamento.value="";shipinfo.dia_pagamento.value="";shipinfo.dia_pagamento.focus(); return (false);}
	if (data_informada<data_hoje && pgto_atrasado==0){
		window.alert ("A data preenchida é menor que a atual.");
		if (confirm ("Deseja efetuar o pagamento com data atrasada?")) {
			shipinfo.pgto_atrasado.value=shipinfo.user.value;
			 
		}
		else{
			shipinfo.ano_pagamento.value="";
			shipinfo.mes_pagamento.value="";
			shipinfo.dia_pagamento.value="";
			shipinfo.pgto_atrasado.value="";
			shipinfo.dia_pagamento.focus();
		}
	return (false);}
	
	
return (true);
	}
// End -->
</script>
<body onunload="window.opener.location.reload(true);">

<table border="1" style="border:2px solid black; border-collapse:collapse;">
	<tr style="background-color:<%=cor_item1%>; font-size:16px; color:white;">
		<td align="center" colspan="8"><b><%if lista_pagto <> "" OR cd_pagamento <> "" Then%>EMISSÃO DE PAGAMENTOS<%else%>CHEQUES CANCELADOS / DIVERSOS<%end if%></b></td>
	</tr>
	<%num_linha = 1
	cor =  1
	cor_linha = "#FFFFFF"
	cor_linha2 = "#e9e9e9"
	
	'*** Monta a lista para gravar no Banco de Dados ***
	'if cd_conta_banco = "" AND nm_cheque = "" AND lista_pagto <> "" Then
	if cd_pagamento = "" AND lista_pagto <> "" Then
		pagto_array = split(lista_pagto,",")
		For Each pagto_item In pagto_array
				if cor > 1 then
					cor_linha = "#d7d7d7"
					'cor_linha2 = "#d7d7d7"
					cor = 1
				else
					cor_linha = "#FFFFFF"
					'cor_linha2 = "#e9e9e9"
					cor = 2
				end if
			strsql="SELECT * FROM View_financeiro_valores3 where cd_codigo = '"&pagto_item&"'"
			Set rs_pagto = dbconn.execute(strsql)
			while not rs_pagto.EOF
				nm_fornecedor = rs_pagto("nm_fornecedor")
				nm_pagamento = rs_pagto("nm_pagamento")
				dt_vencimento = rs_pagto("dt_vencimento")
					dt_vencimento = zero(day(dt_vencimento))&"/"&zero(month(dt_vencimento))&"/"&right(year(dt_vencimento),2)
					
				nm_valor = rs_pagto("nm_valor")
				soma_valores = soma_valores + nm_valor
				if cab = "" then%>
				<tr style="background-color:<%=cor_item2%>;">
					<td>&nbsp;</td>
					<td align="center" colspan="3"><b>Descrição</b></td>
					<td align="center"><b>Venc.</b></td>
					<td align="center"><b>Valor</b></td>
				</tr>
				<%cab = 1
				end if%>
				<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
					<td align="center"><%=num_linha%></td>
					<td align="left" colspan="3"><%="<b>"&nm_fornecedor&"</b> - "&nm_pagamento%> <%'=cor%></td>
					<td align="center"><%=dt_vencimento%></td>
					<td align="right" style="padding-right:5px;"><%if nm_valor > 0 then response.write(formatNumber(nm_valor))%></td>
				</tr>
			<%num_linha = num_linha + 1
			rs_pagto.movenext
			wend
		next
	
	'*** Lista o pagamento gravado ***
	'elseif cd_conta_banco <> "" AND nm_cheque <> "" AND lista_pagto = "" Then
	elseif cd_pagamento <> "" AND lista_pagto = "" Then
		'strsql="SELECT * FROM View_financeiro_valores3 where cd_conta_banco='"&cd_conta_banco&"' AND nm_cheque='"&nm_cheque&"'"
		strsql="SELECT * FROM View_financeiro_valores3 where cd_cheque='"&cd_pagamento&"' order by dt_vencimento"
		Set rs_pagto = dbconn.execute(strsql)
			while not rs_pagto.EOF
				cd_cheque = rs_pagto("cd_cheque")
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
				dt_pagamento = rs_pagto("dt_pagamento")
					dt_pagamento = zero(day(dt_pagamento))&"/"&zero(month(dt_pagamento))&"/"&year(dt_pagamento)
					
				if cab = "" then%>
				<tr style="background-color:<%=cor_item2%>;">
					<td>&nbsp;</td>
					<td align="center" colspan="3"><b>Descrição</b></td>
					<td align="center"><b>Venc.</b></td>
					<td align="center"><b>Valor</b></td>
				</tr>
				<%cab = 1
				end if%>
				<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
					<td align="center"><%=num_linha%></td>
					<td align="left" colspan="3"><%="<b>"&nm_fornecedor&"</b> - "&nm_pagamento%></td>
					<td align="center"><%=dt_vencimento%></td>
					<td align="right" style="padding-right:5px;"><%if nm_valor > 0 then response.write(formatNumber(nm_valor))%></td>
					<!--td><img src="../imagens/ic_reprovado.gif" alt="Excluir (<%=cd_cheque%>)" width="10" height="12" border="0" onClick="javascript:JsValorDelete('<%=cd_cheque%>')"><%=cd_cheque%>--></td>
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
			<td><img src="../imagens/blackdot.gif" alt="" width="20" height="2" border="0"></td>
			<td colspan="3"><img src="../imagens/blackdot.gif" alt="" width="450" height="2" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="70" height="2" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="100" height="2" border="0"></td>
			
		</tr>
		<tr>
			<td colspan="5"></td>
			<td align="right" bgcolor="#d7d7d7" style="padding-right:5px;"><b><%if soma_valores > 0 then response.write(formatNumber(soma_valores))%></b></td>
		</tr>
		<%If cd_pagamento = "" AND lista_pagto <> "" Then%><br>
		<form name="cheques" id="cheques" action="acoes/acoes3.asp" method="post" onsubmit="return validar_pagto(document.cheques);">
		<input type="hidden" name="acao" value="insere_cheque">
		<input type="hidden" name="pgto_atrasado">
		<input type="hidden" name="lista_pagto" value="<%=lista_pagto%>">
		<input type="hidden" name="cd_status" value="1">
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		<input type="hidden" name="nm_valor_cheque" value="<%=soma_valores%>">
		<tr><td colspan="6">&nbsp;</td></tr>
		<tr>
			<td style="background-color:<%=cor_item2%>;"></td>
			<td align="center" style="background-color:<%=cor_item2%>;"><b>Tipo</b><%=soma_valores%></td>
			<td align="center" colspan="2">
				<%str_seccao = "modo_pagto"%>
				<input type="hidden" name="escolha_documentos" value="<%=escolha_documentos%>">
				<input id="cheque" type="radio" name="modo_pagamento" value="1" onFocus="nextfield ='cd_conta_banco';" onclick="hideandshow_b('<%=str_seccao%>','inline');"> Cheque <img src="../imagens/px.gif" alt="" width="50" height="1" border="0">
				<input type="radio" name="modo_pagamento" id="transf" value="2" onFocus="nextfield ='dia_pagamento';" onclick="hideandshow_b('<%=str_seccao%>','none');"><!--checked--> Transferência
			</td>
			<td colspan="2" align="center">&nbsp;<!-- <b>Confirma</b>--></td>
		</tr>
		<tr><td colspan="6">&nbsp;</td></tr>
		<tr style="background-color:<%=cor_item2%>;">
			<td><img src="../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
			<td align="center"><span class="<%=str_seccao%>"><b>Conta</b></span></td>
			<td align="center"><span class="<%=str_seccao%>"><b>N° Cheque</b></span></td>
			<td align="center"><b>Data Pgto</b></td>
			<td></td>
			<td align="center">&nbsp;<!-- <b>Confirma</b>--></td>
		</tr>		
		
		
		<tr>
			<td></td>
			<td><span class="<%=str_seccao%>"><%=cd_conta_banco%>
				<select name="cd_conta_banco" id="conta" onFocus="nextfield ='nm_cheque';">
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
				</span>
			</td>
			<td align="center"><span class="<%=str_seccao%>"><input type="text" name="nm_cheque" maxlength="10" size="10" onFocus="nextfield ='dia_pagamento';"></span></td>
			<td align="center"><input type="text" name="dia_pagamento" value="<%=dia_sel%>" size="2" maxlength="2" onFocus="nextfield ='mes_pagamento';" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">
				<input type="text" name="mes_pagamento" value="<%=mes_sel%>" size="2" maxlength="2" onFocus="nextfield ='ano_pagamento';" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">
				<input type="text" name="ano_pagamento" value="<%=ano_sel%>" size="4" maxlength="4" onFocus="nextfield ='';"></td>
			<td align="center"><input type="submit" name="Ok" value="OK"></td>
			<td align="center"><!--<input type="checkbox" name="cd_confirma_pagto" value="1">-->&nbsp;</td>
			
		</tr>
		</form>
		<%Elseif cd_pagamento <> "" AND lista_pagto = "" Then%><br>
		<tr><td colspan="6">&nbsp;</td></tr>
		<tr style="background-color:<%=cor_item2%>;">
			<td></td>
			<td align="center"><%if cd_modo_pagto = 1 then response.write("<b>Conta</b>")%></td>
			<td align="center"><%if cd_modo_pagto = 1 then response.write("<b>N° Cheque</b>")%></td>
			<td align="center"><b>Data Pgto</b></td>
			<td>&nbsp;</td>
			<td align="center">&nbsp;<!--<b>Confirma</b>--></td>
		</tr>		
		<tr>
			<td></td>
			<td align="center"><%=nm_empresa_abrv%></td>
			<td align="center"><%=nm_cheque%></td>
			<td align="center"><%=dt_pagamento%></td>
			<td align="center">&nbsp;</td>
			<td align="center">&nbsp;</td>
			
		</tr>
		<%else%>
		<form name="cheques" id="cheques" action="acoes/acoes3.asp" method="post" onsubmit="return validar_pagto(document.cheques);">
		<input type="hidden" name="acao" value="insere_cheque">
		<input type="hidden" name="lista_pagto" value="<%=lista_pagto%>">
		<!--input type="hiddenS" name="cd_status" value="0"-->
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		<tr style="background-color:<%=cor_item2%>;">
			<td></td>
			<td align="center"><span class="<%=str_seccao%>"><b>Conta</b></span></td>
			<td align="center"><span class="<%=str_seccao%>"><b>N° Cheque</b></span></td>
			<td align="center"><b>Data</b></td>
			<td align="center"><b>Valor</b></td>
			<td align="center"><b>Tipo</b><!--&nbsp; <b>Confirma</b>--></td>
		</tr>
		<tr>
			<td></td>
			<td align="center"><span class="<%=str_seccao%>">
				<select name="cd_conta_banco" id="conta" onFocus="nextfield ='nm_cheque';">
					<option value=""></option>
					<%strsql = "SELECT * FROM View_financeiro_banco where cd_status = 1"
					Set rs_banco = dbconn.execute(strsql)
						while not rs_banco.EOF
							nm_empresa = rs_banco("nm_empresa_abrv")
							cd_conta_bancaria = rs_banco("cd_codigo")
							cd_empresa = rs_banco("cd_empresa")
						rs_banco.movenext%>
						<option value="<%=cd_conta_bancaria%>" <%if abs(cd_conta_bancaria) = abs(cd_conta_banco) then response.write("SELECTED")%>><%=nm_empresa%></option>
						<%wend%>
				</select>
				</span></td>
			<td align="center"><span class="<%=str_seccao%>"><input type="text" name="nm_cheque" maxlength="10" size="10" value="<%=nm_cheque%>" onFocus="nextfield ='dia_pagamento';"></span></td>
			<td align="center"><input type="text" name="dia_pagamento" value="<%=dia_pagamento%>" size="2" maxlength="2" onFocus="nextfield ='mes_pagamento';" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">
				<input type="text" name="mes_pagamento" value="<%=mes_pagamento%>" size="2" maxlength="2" onFocus="nextfield ='ano_pagamento';" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">
				<input type="text" name="ano_pagamento" value="<%=ano_pagamento%>" size="4" maxlength="4" onFocus="nextfield ='';"></td>
			<td align="center"><input type="text" name="nm_valor_cheque" size="7" value="<%=nm_valor_cheque%>"></td>
			<td align="left">
				&nbsp;<input type="radio" name="cd_status" value="0" <%if cd_status = 0 then response.write("checked")%>>Cancelados<br>
				&nbsp;<input type="radio" name="cd_status" value="2" <%if cd_status = 2 then response.write("checked")%>>Diversos</td>						
		</tr>
		<tr style="background-color:<%=cor_item2%>;">
			<td align="center">&nbsp;</td>
			<td align="center"><b>Obs.:</b></td>
			<td align="right" colspan="2"><textarea cols="35" rows="2" name="nm_obs"><%=nm_obs%></textarea> &nbsp;</td>
			<td align="center"><input type="submit" name="Ok" value="OK"></td>
			<td align="center">&nbsp;</td>	
		</tr>
		<%if mens <> "" Then%><tr><td colspan="5" align="center"><span style="color:red;"><%=mens%></span></td></tr><%end if%>
		</form>
		<%end if%>
		<tr>
			<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="110" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="180" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="150" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="70" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
			
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