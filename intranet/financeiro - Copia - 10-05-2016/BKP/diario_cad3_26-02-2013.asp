<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<%
cd_user = session("cd_codigo")
data_atual = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)

dt_dia = request("dia_sel")
dt_mes = request("mes_sel")
dt_ano = request("ano_sel")
	dia_final = ultimodiames(dt_mes,dt_ano)
	
	mes_anterior = dt_mes - 1
	ano_anterior = dt_ano
		if mes_anterior < 1 then
			mes_anterior = 12
			ano_anterior = dt_ano - 1
		end if
		dia_final_anterior = ultimodiames(mes_anterior,ano_anterior)

cd_diario = request("cd_diario")
str_origem = request("cd_origem")
	response.write(str_origem)
	sep_origem1 = instr(1,str_origem,".",1)
	if sep_origem1 > 0 then
		cd_origem = mid(str_origem,1,sep_origem1)
		codigo_origem = mid(str_origem,sep_origem1, len(str_origem))
			cd_origem = replace(cd_origem,".","")
			codigo_origem = replace(codigo_origem,".","")
	end if
	
		
		if cd_origem = "" then
		elseif cd_origem = 6 then
			strsql ="SELECT * FROM View_manutencao_orcamento WHERE cd_codigo='"&codigo_origem&"' ORDER BY dt_orcamento"
			Set rs = dbconn.execute(strsql)
			if not rs.EOF then nr_orcamento = rs("nr_orcamento")
		end if
			
	
acao = request("acao")



%>
<SCRIPT LANGUAGE="javascript">

function mostra_cancelamento(){
	 var oHTTPRequest_canc = createXMLHTTP(); 
     oHTTPRequest_canc.open("Post", "../financeiro/ajax/ajax_cancela.asp", true);
     oHTTPRequest_canc.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_canc.onreadystatechange=function(){
      if (oHTTPRequest_canc.readyState==4){
         document.all.divCanc_mostra.innerHTML = oHTTPRequest_canc.responseText;}}
       oHTTPRequest_canc.send("cd_cancela=1");
 }

function parcelado(){
	if (document.getElementById('tipo_valor').value == "3")  hideandshow_b('mostra_parcelas','inline')
	else
		hideandshow_b('mostra_parcelas','none');
		//document.diario.cd_parcela = "";
	}	


function parcelamento(){
	var qtd = document.diario.cd_qtd_parcelas.value;
	var valor = document.diario.nm_valor_ori.value;
	var	valor = valor.replace('.','');
	var	valor = valor.replace(',','.');	
		if (qtd > 0){
			var valor_total = parseFloat(valor/qtd);	
			var	valor = valor.replace('.','');
			}
		else{
			var valor_total = parseFloat(valor/1);
			}
			
			document.diario.nm_valor.value = valor_total.toFixed(2);
		}
 
 </script>
<html>
<head>
	<title>Financeiro - Diário</title>
</head>
<!--#include file="../js/ajax.js"-->
<!--#include file="../ferramentas/js/ferramentas.js"-->
<body onLoad="diario.cd_fornecedor.focus();" onunload="window.opener.location.reload(true);">


<table border="1" align="center" style="font-size:12px; font-family:arial;">
	<form action="../financeiro/acoes/acoes3.asp" name="diario" id="diario" method="post">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="data_atual" value="<%=data_atual%>">
	
	<input type="hidden" name="cd_diario" value="<%=cd_diario%>">
	
	
	<tr style="background-color:gray; font-size:16px; color:white;">
		<td colspan="3">Cadastramento de pagamentos - v3 <%=str_origem%><%'=acao%><%'=cd_os%><%=cd_diario& "-" &dt_mes&"/1/"&dt_ano&" - "&dt_mes&"/"&dia_final&"/"&dt_ano%></td>		
	</tr>
	<%if cd_diario <> "" Then
		xsql = "up_financeiro_diario3_lista_individual @cd_diario='"&cd_diario&"', @dt_i='"&dt_mes&"/1/"&dt_ano&"', @dt_f='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i_a='"&mes_anterior&"/1/"&ano_anterior&"', @dt_f_a='"&mes_anterior&"/"&dia_final_anterior&"/"&ano_anterior&"'"
		Set rs_diario = dbconn.execute(xsql)
			if NOT rs_diario.EOF Then
				
				cd_fornecedor = rs_diario("cd_fornecedor")
				nm_fornecedor = rs_diario("nm_fornecedor")
				dt_inicial = rs_diario("dt_inicial")
					ano_inicial = year(dt_inicial)
					mes_inicial = zero(month(dt_inicial))
					
				dt_final = rs_diario("dt_final")
					ano_final = year(dt_final)
					mes_final = zero(month(dt_final))
					
				nm_pagamento = rs_diario("nm_pagamento")
				nm_descricao = rs_diario("nm_descricao")
				cd_tipo_valor = rs_diario("cd_tipo_valor")
				dt_vencimento = rs_diario("dt_vencimento")
				ordem_padrao = rs_diario("dia_vencimento_padrao")
				cd_valor = rs_diario("cd_valor")
					'if cd_valor <> "" Then
					'	acao = "editar_valor"
					'else
					'	acao = "inserir_valor"
					'end if
				cd_suspenso = rs_diario("cd_suspenso")
				
				nm_valor = rs_diario("nm_valor")
				cd_qtd_parcelas = rs_diario("cd_qtd_parcelas")
				cd_parcela = rs_diario("cd_parcela")
				nm_obs = rs_diario("nm_obs")
				cd_area = rs_diario("cd_area")
				cd_centrocusto = rs_diario("cd_centrocusto")
				cd_conta = rs_diario("cd_conta")
				cd_unidade = rs_diario("cd_unidade")
				str_origem = rs_diario("cd_origem")
					if str_origem = "" then
						str_origem = "5.1"
					end if
			end if
	else
		'if cd_os <> "" Then
		if codigo_origem <> "" Then
			acao = "inserir"'_valor"
			'cd_os = request("cd_os")
			cd_fornecedor = request("cd_forn")
				strsql = "Select * From TBL_fornecedor where cd_codigo="&cd_fornecedor&""
				Set rs_fornec = dbconn.execute(strsql)
				if not rs_fornec.EOF then
					nm_fornecedor = rs_fornec("nm_fornecedor")
				end if
				
			'cd_equipamento = request("cd_equip")
			num_os_assist = nr_orcamento' str_origem'request("num_os_assist")
				'strsql = "Select * From TBL_equipamento where cd_codigo="&cd_equipamento&""
				'Set rs_equip = dbconn.execute(strsql)
				'if not rs_equip.EOF then
				'	nm_item_orcamento = rs_equip("nm_equipamento_novo")
				'end if
			nm_valor = request("cd_valor_orc")
			cd_tipo_orc = request("cd_tipo_orc")
				'if cd_tipo_orc = 2 then
					nm_item_orcamento = "Orçamento: "&num_os_assist
				'elseif cd_tipo_orc = 3 then
				'	nm_item_orcamento = "Orç.: "&num_os_assist
				'end if
			'cd_unidade = request("cd_unidade")
				
			cd_area = 1
			cd_centrocusto = 5
			cd_conta = 10
		end if
	end if%>
	<input type="hidden" name="cd_origem" value="<%=str_origem%>">
	<%if acao = "inserir_valor" OR acao = "editar_valor" then '*** Impede que apareçam campos desnecessários ***%>
		<td>Favorecido</td>
		<td colspan="2"><%=nm_fornecedor%></td>
	<%else%>
	<tr>
		<td>Favorecido<%=cd_origem&"."&codigo_origem%></td>
		<td colspan="2"><span id="divForn"><%strsql = "Select * From TBL_fornecedor order by nm_fornecedor"
				Set rs_fornec = dbconn.execute(strsql)%>
				<select name="cd_fornecedor" class="inputs">
					<option value=""></option>
					<%While Not rs_fornec.eof
						cd_fornec = rs_fornec("cd_codigo")
						nm_fornec = rs_fornec("nm_fornecedor")%>
						<option value="<%=cd_fornec%>" <%if int(cd_fornecedor) = cd_fornec then%>Selected<%end if%>><%=nm_fornec%></option>
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
	<%end if%>
	<tr>
		<td>Item</td>
		<td colspan="2"><input type="text" name="nm_descricao" value="<%if acao = "editar" OR acao = "inserir_valor" then%><%=nm_descricao%><%elseif acao = "editar_valor" then%><%=nm_pagamento%><%end if%><%=nm_item_orcamento%>" size="40" maxlength="500"></td>
	</tr>
	<%'if acao = "inserir_valor" OR acao = "editar_valor" then '*** Impede que apareçam campos desnecessários ***
	'else%>
	<tr>
		<td>Tipo</td>
		<td ><%strsql = "Select * From TBL_financeiro_tipo_pagamento order by nm_tipo_pagamento"
				Set rs_tipopag = dbconn.execute(strsql)%>
				<select name="cd_tipo_valor" onclick="parcelado();" id="tipo_valor">
					<option value=""></option>
					<%While not rs_tipopag.EOF
						cd_tipo_pagamento = rs_tipopag("cd_codigo")
						nm_tipo_pagamento = rs_tipopag("nm_tipo_pagamento")%>
						<option value="<%=cd_tipo_pagamento%>" <%if int(cd_tipo_valor) = cd_tipo_pagamento then%>SELECTED<%end if%>><%=nm_tipo_pagamento%></option>
					<%rs_tipopag.movenext
					wend
					
					rs_tipopag.close
					Set rs_tipopag = nothing%>
				</select>
			<input type="hidden" name="cd_tipo_valor_atual" value="<%=cd_tipo_valor%>" size="3""></td>
		<td><span style="display:<%if cd_tipo_valor = 3 then%>block<%else%>none<%end if%>;" class="mostra_parcelas"><b>Qtd. Parcelas:</b> <%=cd_parcela%>/<input type="text" name="cd_qtd_parcelas" value="<%=cd_qtd_parcelas%>" size="2" maxlength="2" onkeyup="parcelamento()"></span>
			<input type="hidden" name="cd_qtd_parcelas_atual" value="<%=cd_qtd_parcelas%>" size="2">
			<input type="hidden" name="cd_parcela" value="<%=cd_parcela%>" size="2" maxlength="2"></td>
		
	</tr>
	<!--tr><td  colspan="2">&nbsp;</td></tr-->
	<%'end if%>
	
	<tr>
		<td>Vencimento<%'=ordem_padrao%></td>
		<td colspan="2"><input type="text" name="dia_vencimento" value="<%if acao = "inserir" Then%> <%elseif acao = "editar" Then%><%=ordem_padrao%><%elseif acao = "editar_valor" Then%><%=day(dt_vencimento)%><%'=zero(day(dt_vencimento))%><%else%><%=zero(dt_dia)%><%end if%>" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">
		<input type="<%if acao = "editar_valor" OR str_origem <> "" Then%>text<%else%>hidden<%end if%>" name="mes_vencimento" value="<%'if cd_diario <> "" Then%><%=zero(dt_mes)%><%'else%><%'=left(mes_selecionado(dt_mes),3)&"/"&dt_ano%><%'end if%>" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">
		<input type="<%if acao = "editar_valor" OR str_origem <> "" Then%>text<%else%>hidden<%end if%>" name="ano_vencimento" value="<%=dt_ano%>" size="4" maxlength="4">
		<%if acao = "inserir" AND str_origem = "" then%><%=ucase(left(mes_selecionado(dt_mes),3))&" / "&dt_ano%><%end if%>
		<img src="../imagens/px.gif" alt="" width="30" height="1" border="0">
		<input type="checkbox" name="cd_suspenso" value="1" <%if cd_suspenso = 1 then response.write("checked")%>>Suspender Pagamento</td>
	</tr>
	
	<%if acao = "editar" then '*** Impede que apareçam campos desnecessários ***%>
	<tr bgcolor="#ebebeb">
		<td>Início<%'=ordem_padrao%></td>
		<td colspan="1"><input type="text" name="mes_inicial" value="<%=mes_inicial%>" size="2" maxlength="2">/<input type="text" name="ano_inicial" value="<%=ano_inicial%>" size="4" maxlength="4"></td>
		<td align="center"><%if isnumeric(mes_final) AND isnumeric(ano_final) Then%>&nbsp;<%else%><img src="../imagens/ic_cancela.gif" alt="" width="15" height="15" border="0" onclick="mostra_cancelamento();" title="Cancela o pagamento"><%end if%></td>
	</tr>
	<tr bgcolor="#ebebeb">
		<td>Final</td>
		<td colspan="2">
			<div id="divCanc_mostra">
				<%if isnumeric(mes_final) AND isnumeric(ano_final) Then%>
					<%if cd_tipo_valor = 1 then%>
						<%=zero(mes_final)&"/"&ano_final%>
					<%else%>
						<input type="hiddens" name="mes_final" value="<%=mes_final%>" size="2" maxlength="2">/<input type="hiddens" name="ano_final" value="<%=ano_final%>" size="4" maxlength="4">				
					<%end if%>
				<%else%>
					Data Indefinida
				<%end if%>
			</div></td>
	</tr>
	<%elseif acao = "inserir" OR acao = "inserir_valor" OR acao = "editar_valor" then%>
	<tr>
		<td>Valor</td><input type="hidden" name="cd_valor" value="<%=cd_valor%>">
		<td colspan="2"><input type="text" name="nm_valor" value="<%if nm_valor <> "" Then response.write(formatNumber(nm_valor,2))%>" size="20" maxlength="11"  style="text-align:right"> &nbsp; &nbsp;
					<input type="hidden" name="nm_valor_atual" value="<%if nm_valor <> "" Then%><%'nm_valor = replace(nm_valor,",",".")%><%=nm_valor%><%end if%>" size="2" maxlength="2"> 
		<%'=cd_parcela%></td>
	</tr>
	<%end if%>
	<!--tr>
		<td>Nota Fiscal</td>
		<td><input type="text" name="nm_nf" value="<%if nm_nf <> "" Then response.write(nm_nf)%>" size="15" maxlength="11"> &nbsp; &nbsp;
		<%strsql = "Select * From TBL_NF_tipo order by nm_nf_tipo"
				Set rs_nf = dbconn.execute(strsql)%>
				<select name="cd_nf_tipo" class="inputs">
					<option value=""></option>
					<%Do While Not rs_nf.eof
						strcd_nf_tipo = rs_nf("cd_codigo")
						nm_nf_tipo = rs_nf("nm_nf_tipo")%>	
						<option value="<%=strcd_nf_tipo%>" <%if int(cd_nf_tipo) = strcd_nf_tipo then%>Selected<%end if%>><%=nm_nf_tipo%></option>
						<%rs_nf.movenext
					loop
						rs_nf.close
						Set rs_nf = nothing%>
				</select>	</td>
	</tr-->
	<%if acao = "inserir" OR acao = "inserir_valor" OR acao = "editar_valor" then%><tr>
		<td colspan="3"><img src="../imagens/blackdot.gif" alt="" width="410" height="1" border="0"></td>
	</tr>
	<tr>
		<td>Obs.</td>
		<td colspan="2"><textarea cols="40" rows="2" name="nm_obs"><%if nm_obs <> "" Then response.write(nm_obs)%></textarea></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="3"><img src="../imagens/blackdot.gif" alt="" width="410" height="1" border="0"></td>
	</tr>
	
	<%if acao = "inserir_valor" OR acao = "editar_valor" then '*** Impede que apareçam campos desnecessários ***
	else%>
	<tr>
		<td>Área</td>
		<td colspan="2"><%strsql = "Select * From TBL_financeiro_area order by nm_area"
				Set rs_area = dbconn.execute(strsql)%>
				<select name="cd_area" class="inputs">
					<option value=""></option>
					<%Do While Not rs_area.eof
						strcd_area = rs_area("cd_codigo")
						nm_area = rs_area("nm_area")%>	
						<option value="<%=strcd_area%>" <%if int(cd_area) = strcd_area then%>Selected<%end if%>><%=nm_area%></option>
						<%rs_area.movenext
					loop
						rs_area.close
						Set rs_area = nothing%>
				</select>		
		</td>
	</tr>
	<tr>
		<td>C.Custos</td>
		<td colspan="2"><%strsql = "Select * From TBL_financeiro_centrocusto order by nm_centrocusto"
				Set rs_cc = dbconn.execute(strsql)%>
				<select name="cd_centrocusto" class="inputs">
					<option value=""></option>
					<%Do While Not rs_cc.eof
						strcd_centrocusto = rs_cc("cd_codigo")
						nm_centrocusto = rs_cc("nm_centrocusto")%>	
						<option value="<%=strcd_centrocusto%>" <%if int(cd_centrocusto) = strcd_centrocusto then%>Selected<%end if%>><%=nm_centrocusto%></option>
						<%rs_cc.movenext
					loop
						rs_cc.close
						Set rs_cc = nothing%>
				</select>		
		</td>
	</tr>
	<tr>
		<td>Conta</td>
		<td colspan="2"><%strsql = "Select * From TBL_financeiro_conta order by nm_conta"
				Set rs_conta = dbconn.execute(strsql)%>
				<select name="cd_conta" class="inputs">
					<option value=""></option>
					<%Do While Not rs_conta.eof
						strcd_conta = rs_conta("cd_codigo")
						nm_conta = rs_conta("nm_conta")%>	
						<option value="<%=strcd_conta%>" <%if int(cd_conta) = strcd_conta then%>Selected<%end if%>><%=nm_conta%></option>
						<%rs_conta.movenext
					loop
						rs_conta.close
						Set rs_conta = nothing%>
				</select>		
		</td>
	</tr>
	<tr>
		<td>Unidade</td>
		<td colspan="2"><%strsql = "Select * From TBL_unidades where cd_status=1 and cd_hospital >= 1 order by nm_unidade"
				Set rs_unid = dbconn.execute(strsql)%>
				<select name="cd_unidade" class="inputs">
					<option value=""></option>
					<%Do While Not rs_unid.eof
						strcd_unidade = rs_unid("cd_codigo")
						nm_unidade = rs_unid("nm_unidade")
						nm_sigla = rs_unid("nm_sigla")%>	
						<option value="<%=strcd_unidade%>" <%if int(cd_unidade) = strcd_unidade then%>Selected<%end if%>><%=nm_sigla%></option>
						<%rs_unid.movenext
					loop
						rs_unid.close
						Set rs_unid = nothing%>
				</select>		
		</td>
	</tr>
	<%end if%>
	<tr><td align="center" colspan="3">
		<%'=replace(acao,"_"," ")
			
		
			if acao = "inserir" OR acao = "inserir_valor" then
				nm_botao = "inserir"
			elseif acao = "editar" OR acao = "editar_valor" then
				nm_botao = "Atualizar"
			end if%>
		<input type="hidden" name="acao" value="<%=acao%>"><img src="../imagens/px.gif" alt="" width="100" height="1" border="0">
		<input type="submit" name="incluir" value="<%=nm_botao%>"><img src="../imagens/px.gif" alt="" width="80" height="1" border="0">
		<input type="button" name="Finalizar" value="Sair" onclick="window.close();"><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"> 
		<%if cd_diario <> "" Then%><img src="../imagens/ic_del.gif" alt="Excluir" width="13" height="14" border="0" onclick="JsDelete(<%=cd_diario%><%'=","&str_origem%>)"><%end if%></td></tr>
	</form>	
</table>
<SCRIPT LANGUAGE="javascript">
	formataMoeda(document.forms.diario.nm_valor, 2);
	
	//function JsDelete(cod,cod2)
	function JsDelete(cod)
	   {
		if (confirm ("Tem certeza que deseja excluir?"))
	  {
	  //document.location.href('acoes/acoes3.asp?cd_diario='+cod+'&cd_origem='+cod2+'&acao=excluir_pagamento');
	  document.location.href('acoes/acoes3.asp?cd_diario='+cod+'&acao=excluir_pagamento');
		}
		  }
</SCRIPT>


</body>
</html>
