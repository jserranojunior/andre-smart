<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<%dt_ano = request("ano_sel")
dt_mes = request("mes_sel")
	if dt_mes = "" OR dt_ano = "" then
		dt_mes = month(now)
		dt_ano = year(now)
	end if
		dia_final = ultimodiames(dt_mes,dt_ano)
		
		mes_anterior = dt_mes - 1
		ano_anterior = dt_ano
			if mes_anterior < 1 then
				mes_anterior = 12
				ano_anterior = dt_ano - 1
			end if
			dia_final_anterior = ultimodiames(mes_anterior,ano_anterior)
			'response.write("/"&mes_anterior&"/"&ano_anterior)
		
		mes_posterior = dt_mes + 1
		ano_posterior = dt_ano
			if mes_posterior > 12 then
				mes_posterior = 1
				ano_posterior = dt_ano + 1
			end if
			dia_final_posterior = ultimodiames(mes_posterior,ano_posterior)



cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)
	
	'if  year(now)&zero(month(now))&zero(day(now)) > dt_ano&zero(dt_mes)&dia_final then
	if  year(now)&zero(month(now)) > dt_ano&zero(dt_mes) then
		ordem = " dt_vencimento "'
	elseif year(now)&zero(month(now)) = dt_ano&zero(dt_mes) then
		ordem = " pre_ordem "
	elseif year(now)&zero(month(now)) < dt_ano&zero(dt_mes) then
		ordem = " dia_vencimento_padrao "
		'ordem = " pre_ordem "
	end if%>
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
	end if%>

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
	 
	if (document.getElementById('cheque').checked == false && document.getElementById('transf').checked == false){window.alert ("Selecione o tipo de pagamento");document.getElementById('cheque').focus(); return (false);}
	if (document.getElementById('cheque').checked == true && document.getElementById('conta').value == ""){window.alert ("Selecione a conta");document.getElementById('conta').focus(); return (false);}
	if (document.getElementById('cheque').checked == true && shipinfo.nm_cheque.value == ""){window.alert ("informe o no. do cheque");shipinfo.nm_cheque.focus(); return (false);}
	
	if (shipinfo.dia_pagamento.value.length==""){window.alert ("A data deve ser preenchida.");shipinfo.dia_pagamento.focus(); return (false);}
	if (shipinfo.mes_pagamento.value.length==""){window.alert ("A data deve ser preenchida.");shipinfo.mes_pagamento.focus(); return (false);}
	if (shipinfo.ano_pagamento.value.length==""){window.alert ("A data deve ser preenchida.");shipinfo.ano_pagamento.focus(); return (false);}
	
	if (data_informada<data_hoje){window.alert ("A data preenchida é menor que a atual.");shipinfo.ano_pagamento.value="";shipinfo.mes_pagamento.value="";shipinfo.dia_pagamento.value="";shipinfo.dia_pagamento.focus(); return (false);}
	return (true);
	}
// End -->
</script>
<body onunload="window.opener.location.reload(true);">
<table cellspacing="1" cellpadding="1" border="0" class="no_print" style="background-color:#b8d8da; border:black solid 2px;" align="center">
	<form action="../financeiro.asp" name="seleciona_periodo" id="seleciona_período">
	<input type="hidden" name="tipo" value="emissao_cheques">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="data_atual" value="<%=data_atual%>">
	<tr>
		<td align="center" colspan="2">CONTAS A PAGAR - Emissão de cheques <%'=ordem%></td>
	</tr>
	<tr>
		<td><b>M&ecirc;s</b></td>
		<td><b>Ano</b></td>
	</tr>
	<tr>
		
		<td><select name="mes_sel">
					<option value=""></option>
					<%for imes = 1 to 12%>
						<option value="<%=imes%>" <%if abs(imes)=abs(dt_mes) then response.write("SELECTED")%>><%=mes_selecionado(imes)%></option>
					<%next%>
			</select></td>	
		<td><input type="text" name="ano_sel" maxlength="4" size="4" value="<%=dt_ano%>">&nbsp;<input type="submit" name="OK" value="OK" width="4" height="5"></td>
	</tr>
	</form>
	<tr>
		<td align="center" colspan="2"><a href="financeiro.asp?tipo=emissao_cheques&mes_sel=<%=mes_anterior%>&ano_sel=<%=ano_anterior%>"><< Anterior </a>&nbsp; &nbsp; &nbsp; <a href="financeiro.asp?tipo=emissao_cheques&mes_sel=<%=month(now)%>&ano_sel=<%=year(now)%>">Atual</a> &nbsp; &nbsp; &nbsp; <a href="financeiro.asp?tipo=emissao_cheques&mes_sel=<%=mes_posterior%>&ano_sel=<%=ano_posterior%>">Próximo>></a></td>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
</table>
<br class="no_print">
<table style="border:1px solid white; border-collapse:collapse;" align="center">
	<tr><td>|<a href="financeiro.asp?tipo=diario3"> Contas a pagar </a>|<a href="financeiro.asp?tipo=relat_areas"> Relatório áreas </a>|</td></tr>	
</table>
<br class="no_print">
<table border="1" style="border:2px solid black; border-collapse:collapse;" align="center">
	<tr style="background-color:<%=cor_item1%>; font-size:16px; color:white;">
		<td align="center" colspan="8"><b>CHEQUES EMITIDOS</b></td>
	</tr>
	<tr style="background-color:<%=cor_item2%>;">
		<td>&nbsp;</td>
		<td >&nbsp;</td>
		<td align="center"><b>Conta</b></td></td>
		<td align="center"><b>N° Cheque</b></td>
		<td align="center"><b>Data Pagamento</b></td>
		<td align="center"><b>Valor</b></td>
	</tr>
	<tr>
		<td><%'=lista_pagto&" ; "&ver_virgula%></td>
	</tr>
	<%num = 1
	cor = 1
	strsql = "SELECT nm_empresa_abrv, cd_cheque, nm_cheque, dt_pagamento, SUM(nm_valor) AS total FROM View_financeiro_valores3 WHERE (nm_cheque IS NOT NULL) AND (MONTH(dt_vencimento) = "&dt_mes&") AND (YEAR(dt_vencimento) = "&dt_ano&") GROUP BY nm_empresa_abrv, cd_cheque, nm_cheque, dt_pagamento order by nm_empresa_abrv, dt_pagamento"
	Set rs_cheques = dbconn.execute(strsql)
	
	while not rs_cheques.EOF
		nm_empresa_abrv = rs_cheques("nm_empresa_abrv")
		cd_cheque = rs_cheques("cd_cheque")
		nm_cheque = rs_cheques("nm_cheque")
		dt_pagto = rs_cheques("dt_pagamento")
			dt_pagamento = zero(day(dt_pagto))&"/"&zero(month(dt_pagto))&"/"&year(dt_pagto)
			
		valor_cheque = rs_cheques("total")%>
		<tr style="background-color:gray; color:white; font-size:12px;">
			<td align="center"><%=zero(num)%></td>
			<td>&nbsp;</td>
			<td align="center"><%=nm_empresa_abrv%></td>
			<td align="center"><%=nm_cheque%></td>
			<td align="center"><%=dt_pagamento%></td>
			<td align="right"><%=formatnumber(valor_cheque,2)%></td>
		</tr>
			<%'**** Lista os cheques ***
			num_linha = 1
			strsql="SELECT * FROM View_financeiro_valores3 where cd_cheque='"&cd_cheque&"' order by dt_vencimento"
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
					cd_modo_pagto = rs_pagto("cd_modo_pagto")
					nm_cheque = rs_pagto("nm_cheque")
					dt_pagto = rs_pagto("dt_pagamento")
						dt_pagto = zero(day(dt_pagto))&"/"&zero(month(dt_pagto))&"/"&year(dt_pagto)%>
					<tr bgcolor="<%=cor_linha%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');">
						<td>&nbsp;</td>
						<td align="right"><%=num_linha%></td>
						<td align="left" colspan="2"><%="<b>"&nm_fornecedor&"</b> - "&nm_pagamento%></td>
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
				wend%>
		<tr><td colspan="6"><img src="../imagens/px.gif" alt="" width="1" height="5" border="0"></td></tr>
		
		
	<%num = num + 1
		cor = 1
		cor_linha = "#FFFFFF"
		cor_linha2 = "#e9e9e9"
	rs_cheques.movenext
	wend
	num_linha = 1
	
	%>		
		
		<tr><td colspan="6" bgcolor="black"><img src="../imagens/px.gif" alt="" width="1" height="3" border="0"></td></tr>
		<tr style="font-size:12px;">
			<td colspan="4"></td>
			<td align="center"><b>TOTAL</b></td>
			<td align="right" bgcolor="#d7d7d7" style="padding-right:5px;"><b><%if soma_valores > 0 then response.write(formatNumber(soma_valores))%></b></td>
		</tr>
		
		<tr>
			<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
			
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