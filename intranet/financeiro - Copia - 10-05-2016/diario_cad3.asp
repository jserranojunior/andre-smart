<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->

<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../../js/ajax2.js"></SCRIPT>
<%
cd_user = session("cd_codigo")
pasta_loc = "financeiro\"
arquivo_loc = "diario_cad3.asp"

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
	
	num_hoje = year(now)&zero(month(now))'&zero(day(now))
	num_sel = dt_ano&zero(dt_mes)'&zero(dt_dia)
	dif_datas = 0'num_hoje - num_sel

cd_diario = request("cd_diario")
str_origem = request("cd_origem")
	'response.write(str_origem)
	sep_origem1 = instr(1,str_origem,".",1)
	if sep_origem1 > 0 then
		cd_origem = mid(str_origem,1,sep_origem1)
		codigo_origem = mid(str_origem,sep_origem1, len(str_origem))
			cd_origem = replace(cd_origem,".","")
			codigo_origem = replace(codigo_origem,".","")
			response.write(cd_origem&"."&codigo_origem)
	end if
	
		
		if cd_origem = "" then
		elseif cd_origem = 6 then
			strsql ="SELECT * FROM View_manutencao_orcamento WHERE cd_codigo='"&codigo_origem&"' ORDER BY dt_orcamento"
			Set rs = dbconn.execute(strsql)
			if not rs.EOF then nr_orcamento = rs("nr_orcamento")
		end if
			
acao = request("acao")
orient = request("orient")

%>
<SCRIPT LANGUAGE="javascript">

/*function mostra_cancelamento(cod_cancela){
	 var oHTTPRequest_canc = createXMLHTTP(); 
     oHTTPRequest_canc.open("Post", "../financeiro/ajax/ajax_cancela.asp", true);
     oHTTPRequest_canc.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_canc.onreadystatechange=function(){
      if (oHTTPRequest_canc.readyState==4){
         document.all.divCanc_mostra.innerHTML = oHTTPRequest_canc.responseText;}}
       oHTTPRequest_canc.send("cod_cancela="+ cod_cancela);
 }
 */
 
 /*function mostra_cancelamento1(cod_cancela){
	 var oHTTPRequest_canc = createXMLHTTP(); 
     oHTTPRequest_canc.open("Post", "../financeiro/ajax/ajax_cancela.asp", true);
     oHTTPRequest_canc.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_canc.onreadystatechange=function(){
      if (oHTTPRequest_canc.readyState==4){
         document.all.divCanc_mostra.innerHTML = oHTTPRequest_canc.responseText;}}
       oHTTPRequest_canc.send("cod_cancela="+ cod_cancela);
 }*/
 function configura_pag()
 	{
		document.getElementById("cd_cancela").value=0;
	}
	
 
function mostra_cancelamento(cod_cancela)
		{
			if (document.getElementById("cd_cancela").value==0){
			loadXMLDoc("../financeiro/ajax/ajax_cancela.asp?cod_cancela="+ cod_cancela,function()
		  		{
		  			if (xmlhttp.readyState==4 && xmlhttp.status==200)
		    			{
		    				document.getElementById("divCanc_mostra").innerHTML = xmlhttp.responseText;
							document.getElementById("btn_cancela").src='../../imagens/px.gif';
		    			}
		  			}
				)
			}
		}

		

function cancela_pagamento(mes,ano)
	{
		cancela_pag = document.getElementById("cd_cancela").value
		
		if (cancela_pag==0){
			document.forms[0].mes_final.value=mes;
			document.forms[0].ano_final.value=ano;
				document.getElementById("check_cancela").src='../../imagens/check_ok.gif';
				cancela_pag=1;
		}
		else
		{
			document.forms[0].mes_final.value="";
			document.forms[0].ano_final.value="";
				document.getElementById("check_cancela").src='../../imagens/check_inativo.gif';
				cancela_pag=0;
		}
	document.getElementById("cd_cancela").value=cancela_pag;
	}

function altera_data_cancelamento()
	{
		
		document.forms[0].mes_final.value=document.forms[0].mes_final.value;
		document.forms[0].ano_final.value=document.forms[0].ano_final.value;
	}

function parcelado(){
	if (document.getElementById('tipo_valor').value == "1"){  
			hideandshow_b('mostra_parcelas','none');
			hideandshow_b('mostra_categoria','none');
		}
	else if (document.getElementById('tipo_valor').value == "2"){
			hideandshow_b('mostra_parcelas','none');
			hideandshow_b('mostra_categoria','none');
		}
	else if (document.getElementById('tipo_valor').value == "3"){  
			hideandshow_b('mostra_parcelas','inline');
			hideandshow_b('mostra_categoria','inline');
		}
	else if (document.getElementById('tipo_valor').value == "4"){
			hideandshow_b('mostra_parcelas','none');
			hideandshow_b('mostra_categoria','none');
		}
	else{
			hideandshow_b('mostra_parcelas','none');
		}
}	


/*function parcelamento(){
	//if (document.diario.nm_valor_ori.value != ""){
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
		
		//alert("teste");
	//}
}*/
 
function validar_diario(shipinfo){
	
	hoje = new Date
	dia_hoje = hoje.getDate();
		if (dia_hoje < 10 ) dia_hoje = '0'+ dia_hoje
	mes_hoje = hoje.getMonth()+1;
		if (mes_hoje < 10 ) mes_hoje = '0'+ mes_hoje
	ano_hoje = hoje.getFullYear();
		data_hoje = ano_hoje +''+ (mes_hoje) +''+ dia_hoje;
	//********************************************************
	//alert("teste");
	//data_informada = shipinfo.dt_ano_cirurgia.value +''+ shipinfo.dt_mes_cirurgia.value +''+ shipinfo.dt_dia_cirurgia.value;
	
	//dif_datas = data_hoje-data_informada;
	//alert(data_informada+' ~ '+data_hoje+'='+(data_hoje-data_informada));
	if (shipinfo.cd_fornecedor.value=="") {window.alert("Selecione o favorecido");shipinfo.cd_fornecedor.focus(shipinfo.cd_fornecedor.value="");shipinfo.cd_fornecedor.focus(shipinfo.cd_fornecedor.value="");return (false);}
	if (shipinfo.dia_vencimento.value=="") {window.alert("Insira o dia do vencimento");shipinfo.dia_vencimento.focus(shipinfo.dia_vencimento.value="");shipinfo.dia_vencimento.focus(shipinfo.dia_vencimento.value="");return (false);}
	
	
	return (true);
	}
	
function altera_periodo(mes,ano,origem){
		//alert(mes+ano+origem);
		if (origem==6){
			document.getElementById("mes_inicial").value = mes;
			document.getElementById("ano_inicial").value = ano;
			document.getElementById("mes_final").value = mes;
			document.getElementById("ano_final").value = ano;
		}
	}
	
function altera_inicio_a(ano,origem){
	//document.diario.mes_inicial.value = mes;
		if (origem==6){
			document.getElementById("ano_inicial").value = ano;
		}
	}
</script>

<html>
<head>
	<title>Financeiro - Diário</title>
</head>
<%if acao = "inserir" OR acao = "editar" then
	diario_bg_color = "97c8ec"
elseif acao = "inserir_valor" OR acao = "editar_valor" Then
	diario_bg_color = "42ff42"
end if
%>

<!--#include file="../js/ajax2.js"-->
<!--#include file="../ferramentas/js/ferramentas.js"-->
<!--body onLoad="diario.cd_fornecedor.focus();" onunload="window.opener.location.reload(true);"-->
<!--body onLoad="diario.cd_fornecedor.focus();" onunload="window.opener.diario.submit(true);"-->
<!--body onLoad="configura_pag();" onunload="window.opener.diario.submit(true);" bgcolor="#<%'=diario_bg_color%>"-->
<body onLoad="configura_pag();" onunload="window.opener.location.reload(true);" bgcolor="#<%=diario_bg_color%>">
<!--document.forms["myform"].submit();-->
<!--#include file="../includes/arquivo_loc.asp"-->

<table border="1" align="center" style="width: 450px; font-size:12px; font-family:arial;border-collapse:collapse;" bgcolor="#ffffff">
	<tr><td style="width:85px;"></td><td style="width:315px;"></td><td style="width:50px;"></td></tr>
	<form action="../financeiro/acoes/acoes3.asp" name="diario" id="diario" method="post"  onSubmit="return validar_diario(document.diario);">
		<input type="hidden" name="cd_user" value="<%=cd_user%>">
		<input type="hidden" name="data_atual" value="<%=data_atual%>">
		<input type="hidden" name="cd_diario" value="<%=cd_diario%>">
		<!--input type="hiddens" name="cd_origem" value="<%'=cd_diario%>"-->
		
	
	<tr style="background-color:gray; font-size:16px; color:white;">
		<td colspan="3" align="center">Cadastramento de pagamentos <%'=str_origem%><%="("&acao&")"%><%'=cd_os%><%'=cd_diario& "-" &dt_mes&"/1/"&dt_ano&" - "&dt_mes&"/"&dia_final&"/"&dt_ano%><%'=dif_datas%></td>		
	</tr>
	<!--tr>
		<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
		<td colspan="2"><img src="../imagens/blackdot.gif" alt="" width="340" height="1" border="0"></td>
	</tr-->
	<%if cd_diario <> "" Then
		xsql = "up_financeiro_diario3_lista_individual @cd_diario='"&cd_diario&"', @dt_i='"&dt_mes&"/1/"&dt_ano&"', @dt_f='"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @dt_i_a='"&mes_anterior&"/1/"&ano_anterior&"', @dt_f_a='"&mes_anterior&"/"&dia_final_anterior&"/"&ano_anterior&"'"
		Set rs_diario = dbconn.execute(xsql)
			if NOT rs_diario.EOF Then
				cd_valor = rs_diario("cd_codigo")
				cd_fornecedor = rs_diario("cd_fornecedor")
				nm_fornecedor = rs_diario("nm_fornecedor")
				dt_inicial = rs_diario("dt_inicial")
					ano_inicial = year(dt_inicial)
					mes_inicial = zero(month(dt_inicial))
					
				dt_final = rs_diario("dt_final")
					ano_final = year(dt_final)
					mes_final = zero(month(dt_final))
					'response.write(mes_final&"****<br>")
					
				nm_pagamento = rs_diario("nm_pagamento")
				nm_descricao = rs_diario("nm_descricao")
				cd_tipo_valor = rs_diario("cd_tipo_valor")
					strsql = "Select * From TBL_financeiro_tipo_pagamento Where cd_codigo="&cd_tipo_valor&""
					Set rs_tipopag = dbconn.execute(strsql)
					if not rs_tipopag.EOF then
						nm_tipo_pagamento = rs_tipopag("nm_tipo_pagamento")
					end if
					
				cd_categoria_valor = rs_diario("cd_categoria_valor")
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
					if not isnumeric(cd_parcela) then
						'*** verifica qual foi a ultima parcela antes desta ***
						'xsql_parc = "SELECT * FROM TBL_financeiro_valores3 WHERE cd_diario = '"&cd_diario&"' AND dt_vencimento < '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_vencimento DESC"
						xsql_parc = "SELECT * FROM TBL_financeiro_valores3 WHERE cd_diario = '"&cd_diario&"' AND dt_vencimento < '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_vencimento DESC"
						Set rs_parc = dbconn.execute(xsql_parc)
							if not rs_parc.EOF Then
								cd_parcela = rs_parc("cd_parcela")+1
							end if
					end if
				nm_obs = rs_diario("nm_obs")
				cd_area = rs_diario("cd_area")
				cd_centrocusto = rs_diario("cd_centrocusto")
				cd_conta = rs_diario("cd_conta")
				cd_unidade = rs_diario("cd_unidade")
				str_origem = rs_diario("cd_origem")
					if str_origem = "" OR ISNULL(str_origem) then
						str_origem = "5.1"
					else
						response.write(str_origem)
						cd_origem = left(str_origem,1)
						codigo_origem = mid(str_origem,3,len(str_origem))
					end if
			end if
			
	
	else
			if cd_origem = 6 AND codigo_origem <> "" Then
				strsql = "Select * From TBL_manutencao_orcamento Where cd_codigo="&codigo_origem&""
				Set rs = dbconn.execute(strsql)
				if NOT rs.EOF Then
					cd_fornecedor = rs("cd_fornecedor")
					nr_orcamento = rs("nr_orcamento")
					dt_orcamento = rs("dt_orcamento")
					dt_entrega = rs("dt_entrega")
					cd_qtd_parcelas = rs("nr_parcela")
					nm_valor = rs("nm_valor")
					
					
					cd_qtd_parcelas = cd_qtd_parcelas
					if cd_qtd_parcelas = 1 then
						cd_tipo_valor = 1
					elseif cd_qtd_parcelas > 1 Then
						cd_tipo_valor = 3
						nm_valor = nm_valor/cd_qtd_parcelas
					end if
					nm_item_orcamento = "Orç.: "&nr_orcamento&" - entrega:"&zero(day(dt_entrega))&"/"&zero(month(dt_entrega))&"/"&right(year(dt_entrega),2)
					nm_obs = "Orç.: "&nr_orcamento&" - data entrega:"&zero(day(dt_entrega))&"/"&zero(month(dt_entrega))&"/"&right(year(dt_entrega),2)
					
					cd_area = 2
					cd_centrocusto = 5
					cd_conta = 10
					cd_unidade = 27
					
					response.write("***"&cd_origem)
				end if
			
			
			end if
	
	'	end if
	
	end if%>
	<input type="hidden" name="cd_origem" value="<%=str_origem%>">
	<%if acao = "inserir_valor" OR acao = "editar_valor" then '*** Impede que apareçam campos desnecessários ***%>
		<td><b>Favorecido</b></td>
		<td colspan="2"><%=nm_fornecedor%></td>
	<%else%>
	<tr>
		<td><b>Favorecido</b><%'=cd_origem&"."&codigo_origem%></td>
		<td colspan="2">
			<%if cd_origem <> 6 then%>
			<span id="divForn">
				<%strsql = "Select * From TBL_fornecedor order by nm_fornecedor"
				Set rs_fornec = dbconn.execute(strsql)%>
				<select name="cd_fornecedor" class="inputs">
					<option value=""></option>
					<%While Not rs_fornec.eof
						cd_fornec = rs_fornec("cd_codigo")
						nm_fornec = rs_fornec("nm_fornecedor")%>
						<option value="<%=cd_fornec%>" <%if abs(cd_fornecedor)=cd_fornec then response.write("SELECTED")%>><%=nm_fornec%></option>
						<%rs_fornec.movenext
					Wend
						rs_fornec.close
						Set rs_fornec = nothing%>						
				</select>
			</span>
				<img src="../imagens/reload6.gif" alt="Atualiza listagem" width="10" height="12" border="0" onClick="reload_fornecedor();">&nbsp;
				<img src="../imagens/add1.gif" alt="incluir fornecedor" width="10" height="12" border="0" onClick="window.open('../ferramentas/fornecedores_cad.asp','itens_nomes','width=500,height=210')" return false;">
			<%else%>
				<%strsql = "Select * From TBL_fornecedor Where cd_codigo="&cd_fornecedor&""
				Set rs_fornec = dbconn.execute(strsql)
				if not rs_fornec.EOF Then response.write("<b>"&rs_fornec("nm_fornecedor")&"</b>")%>
				<input type="hidden" name="cd_fornecedor" value="<%=cd_fornecedor%>">
			<%end if%>
		 </td>
	</tr>
	<%end if%>
	<tr>
		<td><b>Item</b></td>
		<%if acao = "inserir_valor" OR acao = "editar_valor" then '*** Impede que apareçam campos desnecessários ***%>
			<td colspan="2"><%=nm_descricao%></td>
		<%else%>
			<td colspan="2"><input type="text" name="nm_descricao" value="<%if acao = "editar" OR acao = "inserir_valor" then%><%=nm_descricao%><%elseif acao = "editar_valor" then%><%=nm_pagamento%><%end if%><%=nm_item_orcamento%>" size="40" maxlength="500"></td>
		<%end if%>
	</tr>
	<%'if acao = "inserir_valor" OR acao = "editar_valor" then '*** Impede que apareçam campos desnecessários ***
	'else%>
	<tr>
		<td><b>Tipo</b> <%'=cd_tipo_valor%></td>
		<%'if acao = "inserir_valor" OR acao = "editar_valor" then '*** Impede que apareçam campos desnecessários ***
		if acao = "inserir_valor" OR acao = "editar_valor" then '*** Impede que apareçam campos desnecessários ***%>
		
			<td><%=nm_tipo_pagamento%>
			<input type="hidden" name="cd_tipo_valor" value="<%=cd_tipo_valor%>">
			<input type="hidden" name="cd_qtd_parcelas_atual" value="<%=cd_qtd_parcelas%>" size="2"> <%if cd_tipo_valor = 3 then%> (<%=cd_qtd_parcelas%>x) - <%=cd_parcela%>ª parcela<%end if%>
			<input type="hidden" name="cd_parcela" value="<%=cd_parcela%>" size="2" maxlength="2"></td>
		<%else%>
			<td>
			<%if cd_origem <> 6 Then%>
				<%strsql = "Select * From TBL_financeiro_tipo_pagamento order by nm_tipo_pagamento"
				Set rs_tipopag = dbconn.execute(strsql)%>
				<select name="cd_tipo_valor" onchange="parcelado();" id="tipo_valor">
					<option value="0"></option>
					<%While not rs_tipopag.EOF
						cd_tipo_pagamento = rs_tipopag("cd_codigo")
						nm_tipo_pagamento = rs_tipopag("nm_tipo_pagamento")
						%>
						<option value="<%=cd_tipo_pagamento%>" <%if int(cd_tipo_valor)=cd_tipo_pagamento then response.write("SELECTED")%>><%=nm_tipo_pagamento%> <%'=cd_tipo_valor&"-"&cd_tipo_pagamento%> </option>
					<%rs_tipopag.movenext
					wend
					
					rs_tipopag.close
					Set rs_tipopag = nothing%>
				</select>
				<input type="hidden" name="cd_tipo_valor_atual" value="<%=cd_tipo_valor%>" size="3""><%'=cd_tipo_valor%>
			<%elseif cd_origem = 6 then	%>	
				<input type="hidden" name="cd_tipo_valor" value="<%=cd_tipo_valor%>" size="3">
				<input type="hidden" name="cd_tipo_valor_atual" value="<%=cd_tipo_valor%>" size="3">
				<%strsql = "Select * From TBL_financeiro_tipo_pagamento where cd_codigo="&cd_tipo_valor&"  order by nm_tipo_pagamento"
				Set rs_tipopag = dbconn.execute(strsql)
				if not rs_tipopag.EOF Then 
					response.write("<b>"&rs_tipopag("nm_tipo_pagamento")&"</b>")
					'cd_tipo_pagamento = rs_tipopag("cd_codigo")
					'nm_tipo_pagamento = rs_tipopag("nm_tipo_pagamento")
				end if%>
				<!--input type="hiddens" name="cd_tipo_valor" value="y<%=cd_tipo_pagamento%>" size="3""-->
				<%'=nm_tipo_pagamento%>
				
				
				
			<%end if%>
			</td>
			<td>
				<!--span style="display:<%if cd_tipo_valor = 3 then%>block<%else%>none<%end if%>;" class="mostra_categoria"-->
				<span style="display:block;" class="mostra_categoria">
					<%if cd_origem <> 6 then%>
					 <select name="cd_categoria_valor">
						<option value="1" <%if cd_categoria_valor = "" OR cd_categoria_valor = 1 then response.write("Selected")%>>$ constante</option>
						<option value="2" <%if cd_categoria_valor = 2 then%>Selected<%end if%>>$ variável</option>
					</select>
					<%elseif cd_origem = 6 then%>
						<%if cd_categoria_valor = "" OR ISNULL(cd_categoria_valor) OR cd_categoria_valor = 1 then 
							response.write("<b>valor Constante</b>")
							cd_categoria_valor = 1
						elseif cd_categoria_valor = 2 then
							response.write("<b>valor variável</b>")
							
						end if%>
						<input type="hidden" name="cd_categoria_valor" value="<%=cd_categoria_valor%>" size="3"">
						
					<%end if%>
				</span>
			</td>
	</tr>
	<tr>
		<td><span style="display:<%if cd_tipo_valor = 3 then%>block<%else%>none<%end if%>;" class="mostra_parcelas"><b>Qtd. Parcelas:</b></td>
		<td><span style="display:<%if cd_tipo_valor = 3 then%>block<%else%>none<%end if%>;" class="mostra_parcelas">
			<%if cd_origem <> 6 then%>
				<%if cd_qtd_parcelas > 1 then%>
					<%=cd_parcela%>/<%=cd_qtd_parcelas%>
				<%end if%>
				<input type="text" name="cd_qtd_parcelas" value="<%=cd_qtd_parcelas%>" size="2" maxlength="2" onkeyup="parcelamento()"> vezes
			<%elseif cd_origem =6 then%>
				<b><%=cd_qtd_parcelas%> vezes</b>
				<input type="hidden" name="cd_qtd_parcelas" value="<%=cd_qtd_parcelas%>">
			<%end if%>
			</span>
			<input type="hidden" name="cd_qtd_parcelas_atual" value="" size="2">
			<input type="hidden" name="cd_parcela" value="<%=cd_parcela%>" size="2" maxlength="2"></td>
		<%end if%>
		
	</tr>
	
	<!--tr><td  colspan="2">&nbsp;</td></tr-->
	<%'end if%>
	
	<tr>
		<td><b><%if acao = "editar_valor" OR str_origem <> "" Then%>Vencimento<%else%>Dia venc.<%end if%><%'=ordem_padrao%></b></td>
		<td colspan="2"><input type="text" name="dia_vencimento" value="<%if acao = "inserir" Then%> <%elseif acao = "editar" Then%><%=ordem_padrao%><%elseif acao = "editar_valor" Then%><%=day(dt_vencimento)%><%'=zero(day(dt_vencimento))%><%else%><%=zero(dt_dia)%><%end if%>" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">
		<input type="<%if acao = "editar_valor" OR str_origem <> "" Then%>text<%else%>hidden<%end if%>" name="mes_vencimento" id="mes_vencimento" value="<%=zero(dt_mes)%>" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" <%if cd_user = 46 then%> onchange="altera_periodo(this.value,document.getElementById('ano_vencimento').value,<%=cd_origem%>)" <%end if%> onFocus="PararTAB(this);">
		<input type="<%if acao = "editar_valor" OR str_origem <> "" Then%>text<%else%>hidden<%end if%>" name="ano_vencimento" id="ano_vencimento" value="<%=dt_ano%>" size="4" maxlength="4"  <%if cd_user = 46 then%> onkeyup="altera_periodo(document.getElementById('mes_vencimento').value,this.value,<%=cd_origem%>)" <%end if%>>
		<%if acao = "inserir" AND str_origem = "" then%><%=ucase(left(mes_selecionado(dt_mes),3))&" / "&dt_ano%><%end if%>
		<img src="../imagens/px.gif" alt="" width="30" height="1" border="0">
		<input type="checkbox" name="cd_suspenso" value="1" <%if cd_suspenso = 1 then response.write("checked")%>>Suspender Pagamento</td>
	</tr>
	<!--tr><td><%'=mes_final&"/"&ano_final%></td></tr-->
</table> 
<!--********************************************************************************
		Fim da area comum ao registro de pagamento e a area de valor do pagamento 
	********************************************************************************-->
<img src="../imagens/px.gif" alt="" width="1" height="5" border="0">
<table border="1" align="center" style="width: 450px; font-size:12px; font-family:arial;border-collapse:collapse;" bgcolor="#ffffff">
	<%if acao = "editar" then '*** Impede que apareçam campos desnecessários ***%>
	<tr bgcolor="#ebebeb">
		<td style="width:85px;">Início<%'=ordem_padrao%></td>
		<td style="width:315px;">
			<input type="text" name="mes_inicial" id="mes_inicial" value="<%=mes_inicial%>" size="2" maxlength="2">/
			<input type="text" name="ano_inicial" id="ano_inicial" value="<%=ano_inicial%>" size="4" maxlength="4">
		</td>
		<td style="width:50px;" align="center">
			<%if int(mes_inicial) = int(dt_mes) AND int(ano_inicial) = int(dt_ano) then' OR isnumeric(mes_final) AND isnumeric(ano_final) Then%>
				&nbsp;
			<%else%>
				<img src="../imagens/ic_cancela.gif" alt="" width="15" height="15" border="0" id="btn_cancela" onclick="mostra_cancelamento('<%=ano_inicial&mes_inicial&"_"&dt_ano&zero(dt_mes)%>');" title="Cancela o pagamento">
			<%end if%>
		</td>
	</tr>
	<tr bgcolor="#ebebeb">
		<td>Final</td>
		<td colspan="2"><input type="hidden" name="mes_final" id="mes_final" value="<%=mes_final%>" size="2">
						<input type="hidden" name="ano_final" id="ano_final" value="<%=ano_final%>" size="2">
						<input type="hidden" name="cd_cancela" id="cd_cancela" value="0" size="2">
			<div id="divCanc_mostra">
				<%if isnumeric(mes_final) AND isnumeric(ano_final) Then%>
					<%if cd_tipo_valor = 1 then%>
						<%=mes_final&"/"&ano_final%>
					<%else%>
						<input type="text" name="mes_finalx" value="<%=mes_final%>" size="2" maxlength="2" onKeyup="altera_data_cancelamento();">/<input type="text" name="ano_finalx" value="<%=ano_final%>" size="4" maxlength="4" onKeyup="altera_data_cancelamento();">
					<%end if%>
				<%else%>
					<%if int(mes_inicial) < int(dt_mes) AND int(ano_inicial) <= int(dt_ano) Then%>
						Conta ativa (Para cancelar clique no botão acima) <img src="../imagens/seta_sup_dir.gif" alt="" width="16" height="12" border="0">
					<%else%>
						Conta ativa. 
					<%end if%>
				<%end if%>
			</div></td>
	</tr>
	<%elseif acao = "inserir" OR acao = "inserir_valor" OR acao = "editar_valor" then%>
		<%'*** Verifica se há pendencia ***
			strsql_verpend = "Select * From TBL_financeiro_pendencias where cd_valor='"&cd_valor&"' and dt_inicio_pend <= '"&dt_mes&"/1/"&dt_ano&"'"
			Set rs = dbconn.execute(strsql_verpend)
				if not rs.EOF then
					pendencia = "1"
					cor_linha_valor = "yellow"
				else
					cor_linha_valor = "white"
				end if%>
						
	<tr style="background-color:<%=cor_linha_valor%>">
		<td>Valor</td><input type="hidden" name="cd_valor" value="<%=cd_valor%>">
		<td colspan="2"><input type="text" name="nm_valor" value="<%if nm_valor <> "" Then response.write(formatNumber(nm_valor,2))%>" size="20" maxlength="11"  style="text-align:right"> &nbsp; &nbsp;
					<input type="hidden" name="nm_valor_atual" value="<%if nm_valor <> "" Then%><%'nm_valor = replace(nm_valor,",",".")%><%=nm_valor%><%end if%>" size="2" maxlength="2"> &nbsp; &nbsp;
					
						<select name="cd_pendencia">
							<option value="0">OK</option>
							<option value="1" <%if pendencia = 1 then response.write("selected")%>>Pendente</option>
							<option value="2" <%if pendencia = 2 then response.write("selected")%>>Postergar</option>
						</select>
		</td>
		<input type="hidden" name="mes_final" value="<%=mes_final%>" size="2" maxlength="2">
		<input type="hidden" name="ano_final" value="<%=ano_final%>" size="4" maxlength="4">
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
						<option value="<%=strcd_area%>" <%if int(cd_area) = strcd_area then response.write("Selected")%>><%=nm_area%></option>
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
						<option value="<%=strcd_centrocusto%>" <%if int(cd_centrocusto) = strcd_centrocusto then response.write("Selected")%>><%=nm_centrocusto%></option>
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
						<option value="<%=strcd_conta%>" <%if int(cd_conta) = strcd_conta then response.write("Selected")%>><%=nm_conta%></option>
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
						<option value="<%=strcd_unidade%>" <%if int(cd_unidade) = strcd_unidade then response.write("Selected")%>><%=nm_sigla%></option>
						<%rs_unid.movenext
					loop
						rs_unid.close
						Set rs_unid = nothing%>
				</select>		
		</td>
	</tr>
	<%end if%>
	
	<%if dif_datas <= 0 then%>
	<tr><td align="center" colspan="3">
		<%'=replace(acao,"_"," ")
			
		
			if acao = "inserir" OR acao = "inserir_valor" then
				nm_botao = "inserir"
			elseif acao = "editar" OR acao = "editar_valor" then
				nm_botao = "Atualizar"
			end if
			
			cd_dt_atual = year(data_atual)%>
		<input type="hiddens" name="acao" value="<%=acao%>"><img src="../imagens/px.gif" alt="" width="100" height="1" border="0">
		<input type="submit" name="incluir" value="<%=nm_botao%>"><%'=acao%><%'=cd_tipo_valor%><img src="../imagens/px.gif" alt="" width="80" height="1" border="0">
		<input type="button" name="Finalizar" value="Sair" onclick="window.close();"><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"> 
		<%if cd_diario <> "" and orient = "reg" Then%>
			<img src="../imagens/ic_del.gif" alt="Exclui a dívida definitivamente." width="13" height="14" border="0" onclick="JsDelete_registro(<%=cd_diario%><%'=","&str_origem%>)"><%'=cd_diario%>
		<%elseif cd_diario <> "" and orient = "valor" Then%>
			<img src="../imagens/ic_del.gif" alt="Apaga o valor" width="13" height="14" border="0" onclick="JsDelete_valor(<%=cd_valor%><%'=","&str_origem%>)"><%'=cd_valor%>
		<%end if%></td></tr>
	<%end if%>
	</form>	
</table>
<SCRIPT LANGUAGE="javascript">
	//formataMoeda(document.forms.diario.nm_valor, 2);
	
	//function JsDelete(cod,cod2)
	function JsDelete_registro(cod)
	   {
		if (confirm ("Tem certeza que deseja excluir este registro?")){
	  		if (confirm ("Esta conta será apagada dos meses anteriores também. Confirma?")){
				document.location=('acoes/acoes3.asp?cd_diario='+cod+'&acao=excluir_pagamento_reg');
			}
		}
	}
		  
	function JsDelete_valor(cod){
		if (confirm ("Tem certeza que deseja excluir registro deste valor?")){
	  		document.location=('acoes/acoes3.asp?cd_valor='+cod+'&acao=excluir_pagamento_valor');
		}
	}
</SCRIPT>


</body>
</html>
