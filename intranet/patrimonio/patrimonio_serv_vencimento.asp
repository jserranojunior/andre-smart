
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
usuario = session("cd_codigo")
cd_pat_codigo = request("cd_pat_codigo")
cd_unidade = request("cd_unidade")

cd_user = usuario
pasta_loc = "patrimonio\"
arquivo_loc = "patrimonio_serv_vencimento.asp"

str_tipo = request("tipo")
	if str_tipo = "serv_confirmadas" then
		str_titulo = " CONFIRMADAS"
		str_tb_cor = "green"
		cod_confirma = 2
	elseif str_tipo = "serv_vencer" then
		str_titulo = " VENCENDO"
		str_tb_cor = "orange"
		cod_confirma = 1
	end if

qtd_tempo = request("qtd_tempo")
qtd_tempo_especifico = request("qtd_tempo_especifico")
mes_sel = request("mes_sel")
	if mes_sel = "" Then
		mes_sel = month(now)
	end if
	
ano_sel = request("ano_sel")
	if ano_sel = "" Then
		ano_sel = year(now)
	end if
	
nome_arquivo = request("nome_arquivo")
	
strcd_unidade  =  request("cd_unidade")
ordem = request("ordem")
	if ordem = "" Then
	ordem = "nm_tipo, no_patrimonio"
	end if

mensagem = request("mens")
ok = request("ok")
%>
<head>
	<title>Vd Lap</title>
</head>
<script language="JavaScript">

function patrimonio_manut_confirm(cd_patrimonio,cd_manutencao,cd_categoria,cod_confirma,dt_data){
	form2.cod_pat_manut.value = cd_patrimonio+'#'+cd_manutencao+'$'+cd_categoria+'*'+cod_confirma+'|'+dt_data;	
	{
		var oHTTPRequest_patrimonio_manut = createXMLHTTP(); 
		oHTTPRequest_patrimonio_manut.open("post", "patrimonio/ajax/ajax_patrimonio_manut_confirm.asp", true);
		oHTTPRequest_patrimonio_manut.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_patrimonio_manut.onreadystatechange=function(){
		if (oHTTPRequest_patrimonio_manut.readyState==4){
	    document.all.divPatrimonio_manut.innerHTML = oHTTPRequest_patrimonio_manut.responseText;}}
	    oHTTPRequest_patrimonio_manut.send("cod_pat_manut=" + form2.cod_pat_manut.value);
	} 
}

</script>



<body>
<!--#include file="../includes/arquivo_loc.asp"-->
<br class="no_print">
<table class="no_print" style="border:3 solid <%=str_tb_cor%>; left:150px;" align="center">
	<form action="patrimonio.asp" method="post" name="form" onsubmit="return validar_busca(document.form);">
	<input type="hidden" name="tipo" value="<%=str_tipo%>">
	<tr style="border:1 solid orange; background-color:silver;">
		<td colspan="10" align="center"><B>LISTAGEM DE MANUTENÇÕES <%=str_titulo%></B></td>
	</tr>
	<tr>
		<td>&nbsp; Ano:</td>
		<td><input type="text" name="ano_sel" value="<%=ano_sel%>"></td>
	</tr>
	<tr>
	    <td>&nbsp; Unidade:</td>
			<%strsql = "SELECT * FROM TBL_unidades where cd_status <> 0 and cd_hospital > 0"
			Set rs_uni = dbconn.execute(strsql)%>
	    <td><select name="cd_unidade" class="inputs" onFocus="nextfield ='num_hospital';">
						<option value=""></option>
						<%Do While Not rs_uni.eof
						if int(strcd_unidade) = rs_uni("cd_codigo") then%><%unidade_check = "selected"%><%end if%>
						<option value="<%=rs_uni("cd_codigo")%>" <%=unidade_check%>><%=rs_uni("nm_unidade")%></option>
						<%rs_uni.movenext
						unidade_check = ""
						loop%>
						</select>						
		</td>
	</tr>
	<tr>
		<td>&nbsp; Período:</td>
		<td><select name="qtd_tempo">
				<option value="-1" <%if int(qtd_tempo) = (-1) then%>Selected<%end if%>>Vencido</option>
				<%mes_atual = month(now)
				mes_busca = 12 - mes_atual
				for i_mes = 0 to mes_busca
					if i_mes = 0 then
						mostra_mes = "Mês atual"
					elseif i_mes = 1 then
						mostra_mes = "Próximo mês"
					else
						mostra_mes = "Próximos "&i_mes&" meses"
					end if
				%>
				<option value="<%=i_mes%>" <%if int(qtd_tempo) = i_mes then%>Selected<%end if%>><%=mostra_mes%></option>
				<%next%>
			</select>
			<input type="radio" name="qtd_tempo_especifico" value="0" <%if qtd_tempo_especifico = 0 then%>checked<%end if%>>Menor/igual &nbsp;
			<input type="radio" name="qtd_tempo_especifico" value="1" <%if qtd_tempo_especifico = 1 then%>checked<%end if%>>Específico</td>
	</tr>
	<tr>		
		<td><input type="submit" name="ok" value="ok"></td>
		<td valign="top"><img src="../imagens/ic_print.gif" alt="Imprimir" width="24" height="24" border="0" onClick="window.print();"> &nbsp;
		<img src="../imagens/ic_print_view.gif" alt="Visualizar" width="24" height="26" border="0" onclick="visualizarImpressao();"> &nbsp;
		<%if str_tipo = "serv_confirmadas" then
			img_confirma = "undo"
			tit_confirma = "desfaz confirmação"%>
			<a href="patrimonio.asp?tipo=serv_vencer&cd_unidade=<%=cd_unidade%>&ano_sel=<%=ano_sel%>"><img src="../imagens/ic_clock.gif" alt="Vencendo" border="0"></a>
		<%elseif str_tipo = "serv_vencer" then
			img_confirma = "check_inativo"
			tit_confirma = "confirma realização"%>
			<a href="patrimonio.asp?tipo=serv_confirmadas&cd_unidade=<%=cd_unidade%>&ano_sel=<%=ano_sel%>"><img src="../imagens/check_ok.gif" alt="Confirmadas" border="0"></a>
		<%end if%>&nbsp;
		<%if strcd_unidade <> "" Then%>
		<!--<img src="../imagens/ic_save.gif" alt="Salvar" width="24" height="26" border="0" onclick="grava_bkp('save');">--><%end if%><!-- &nbsp;
		<img src="../imagens/ic_open.gif" alt="Abrir" width="24" height="26" border="0" onclick="abre_bkp();">--></td>
	</tr>	
	<tr>
		<td valign="top" align="right"><b style="color:red;">Atenção:</b></td>
		<td valign="top"><b style="color:red;">Imprimir com orientação PAISAGEM<br> em folha A4.</b></td>
	</tr>
	<tr ><td colspan="2" align="center"></form><b><%=mensagem%></b>
					<div id="save" style="border:3px solid green; background-color:#dbfdd7; width:200;position:relative; display:none; visibility:hiddden; left:240px; top:-50px;">
					<form action="../patrimonio.asp" name="gravar" id="gravar">
						<input type="hidden" name="ano_sel" value="<%=ano_sel%>">
						<input type="hidden" name="cd_unidade" value="<%=strcd_unidade%>">
						<input type="hidden" name="tipo" value="plan_serv">&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; SALVAR BACKUP<br>&nbsp;Nome: 
						<input type="text" name="nm_bkp" size="18" maxlength="18" style="background-color:white;"><br>&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						<input type="submit" name="Ok" value="Salvar" style="color:white;background-color:#2cbc16;"><input type="button" value="cancela" style="color:silver;background-color:#effeed;" onclick="cancela('<%=ano_sel%>','<%=strcd_unidade%>');"></form></div></td>	
	</tr>
</table><br class="no_print">
<table border="1" style="border:1px solid black; border-collapse=collapse;" align="center">
	<%if strcd_unidade <> "" Then
		strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&strcd_unidade&""
		Set rs_uni = dbconn.execute(strsql)
			If Not rs_uni.eof then
				nm_unidade = rs_uni("nm_unidade")
			end if
	end if
	
	nm_unidade = replace(nm_unidade,"&atilde;","ã")
	nm_unidade = replace(nm_unidade,"&ecirc;","ê")
	nm_unidade = replace(nm_unidade,"&aacute;","á")
	nm_unidade = replace(nm_unidade,"&oacute;","ó")
	%><br>
	<tr style="background-color:#494949; color:#FFFFFF;">
		<td colspan="10" align="center"><b style="font-size:15px;"><%=ucase(nm_unidade)%></b></td>
	</tr>
	<tr style="background-color:gray; color:#FFFFFF;">
		<td colspan="10" align="center"><b style="font-size:12px;">
			
			<%if str_tipo = "serv_confirmadas" then%>Listagem da confirmação dos<br>serviços de prevenção- <%=mesdoano(mes_sel)%>/<%=ano_sel%>
			<%elseif str_tipo = "serv_vencer" then%>Listagem de vencimento para realização<br> dos serviços de prevenção - <%=mesdoano(mes_sel)%>/<%=ano_sel%>
			<%end if%></b>
			
			</td>
	</tr>
	<tr><td colspan="10"><div id="divPatrimonio_manut"></div></td></tr>
	<form  name="form2" action="patrimonio_serv_vencimento.asp" method="post">
	<input type="hidden" name="cod_pat_manut" value="teste">
<%if strcd_unidade <> "" Then
	for i=1 to 3
	lista_num_pat = lista_num_pat&"#m"&i&"-"%><br>
	
	
	
	<%cab_num = 0
	num_linha = 1
	
	if qtd_tempo_especifico = 0 OR qtd_tempo_especifico = "" then
		sinal_busca = "<="
	elseif qtd_tempo_especifico = 1 then
		sinal_busca = "="
	end if
	
	strsql = "SELECT DISTINCT cd_codigo, no_patrimonio, cd_posicao, nm_equipamento_novo FROM VIEW_Patrimonio_lista_vencimento WHERE (cd_unidade_comodato = '"&strcd_unidade&"') AND (tipo_manut IS NOT NULL) AND (tipo_manut = "&i&") ORDER BY cd_posicao, nm_equipamento_novo, no_patrimonio"
	Set rs_lista = dbconn.execute(strsql)%>
	<%do while not rs_lista.EOF
		cd_codigo = rs_lista("cd_codigo")
		lista_patrimonio = lista_patrimonio &","& cd_codigo
	rs_lista.movenext
	loop
		
		lista_patrimonio = mid(lista_patrimonio,2,len(lista_patrimonio))%>
			
				<%lista_array = split(lista_patrimonio,",")%>
				 
					<%For Each lista_item In lista_array
						'*** Retorna os dados das manutenções contidas na ARRAY LISTA
						'xsql ="SELECT * FROM VIEW_Patrimonio_lista_vencimento WHERE (cd_unidade = '"&strcd_unidade&"') AND (cd_codigo = '"&lista_item&"') AND (tipo_manut = "&i&") AND (YEAR(dt_plan_prev) <= '"&year(now)&"') AND dt_vencimento "&sinal_busca&" '"&qtd_tempo&"' ORDER BY dt_plan_prev DESC"
						'xsql ="SELECT * FROM VIEW_Patrimonio_lista_vencimento WHERE (cd_unidade = '"&strcd_unidade&"') AND (cd_codigo = '"&lista_item&"') AND (tipo_manut = "&i&") AND (YEAR(dt_plan_prev) <= '"&year(now)&"') AND dt_vencimento "&sinal_busca&" '"&qtd_tempo&"' ORDER BY dt_plan_prev DESC"
						xsql ="SELECT * FROM VIEW_Patrimonio_lista_vencimento WHERE (cd_unidade_comodato = '"&strcd_unidade&"') AND (cd_codigo = '"&lista_item&"') AND (tipo_manut = "&i&") AND (YEAR(dt_plan_prev) <= '"&year(now)&"') AND dt_vencimento "&sinal_busca&" '"&qtd_tempo&"' ORDER BY dt_plan_prev DESC"
						Set rs_pat = dbconn.execute(xsql)
						
							'*** Verifica se o registro acima Consta como confirmado ***
							'xsql ="SELECT * FROM TBL_Patrimonio_manutencoes_confirma WHERE (cd_unidade = '"&strcd_unidade&"') AND (cd_codigo = '"&lista_item&"') AND (tipo_manut = "&i&") AND (YEAR(dt_data) <= '"&ano_sel&"') AND dt_vencimento "&sinal_busca&" '"&qtd_tempo&"' ORDER BY dt_plan_prev DESC"
							'xsql ="SELECT * FROM TBL_Patrimonio_manutencoes_confirma WHERE cd_patrimonio = '"&lista_item&"' AND cd_manutencao = "&i&" AND cd_categoria = '"&cd_categoria&"' AND YEAR(dt_data) = '"&ano_sel&"' ORDER BY dt_data DESC"
							
							if str_tipo = "serv_confirmadas" then
								xsql ="SELECT * FROM TBL_Patrimonio_manutencoes_confirma WHERE cd_patrimonio='"&lista_item&"' AND cd_manutencao="&i&" AND YEAR(dt_data)='"&ano_sel&"' ORDER BY dt_data DESC"
								Set rs_confirm = dbconn.execute(xsql)
								if not rs_confirm.EOF Then '*** Se houver regisstro, Mostra os confirmados ***
									confirmacao = 1
								else
									confirmacao = 2
								end if
							elseif str_tipo = "serv_vencer" then
								xsql ="SELECT * FROM TBL_Patrimonio_manutencoes_confirma WHERE cd_patrimonio='"&lista_item&"' AND cd_manutencao="&i&" AND YEAR(dt_data)='"&ano_sel&"' ORDER BY dt_data DESC"
								Set rs_confirm = dbconn.execute(xsql)
								if rs_confirm.EOF Then '*** Se NÃO houver registro, Mostra as pendencias ***
									confirmacao = 1
								else
									confirmacao = 2
								end if
							end if
							
							
							if confirmacao = 1 Then
							'if not rs_confirm.EOF Then
							'	dt_confirma = rs_confirm("dt_data")
							'end if
						
								if cab_num = 0 Then%>
								<tr><td colspan="10" align="center"><b><%if i = 1 then%>PREVENTIVA<%elseif i = 2 then%>CALIBRAÇÃO<%elseif i = 3 then%>SEGURANÇA ELÉTRICA<%end if%></b></td></tr>
								<tr style="background-color:silver; color:#4a4a4a">
									<td>&nbsp;<img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
									<td align="center"><b>Nome</b><br><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
									<td align="center"><b>Pat<%'=cd_categoria%></b><br><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td align="center"><b>Marca</b><br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
									<td align="center"><b>Modelo</b><br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
									<td align="center"><b>Nº Série</b><br><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
									<td align="center"><b>Periodicidade</b><br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
									<!--td align="center"><b> </b></td-->
									<td align="center"><b>Venc.</b><br><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center"><b>Agenda</b><br><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
									<td align="center" class="no_print"><b>&nbsp;</b><br><img src="../imagens/px.gif" alt="" width="5" height="1" border="0"></td>
								</tr>
								<%cab_num = cab_num + 1
								end if%>
							
							<%'if confirmacao = 1 Then
								if not rs_pat.EOF then
									cd_patrimonio = rs_pat("cd_codigo")
									nm_tipo = rs_pat("nm_tipo")
									no_patrimonio = rs_pat("no_patrimonio")
									nm_equipamento_novo = rs_pat("nm_equipamento_novo")
									nm_marca = rs_pat("nm_marca")
									cd_pn = rs_pat("cd_pn")
									cd_ns = rs_pat("cd_ns")
									dt_periodicidade_prev = rs_pat("dt_periodicidade_prev")
									dt_periodicidade_cal = rs_pat("dt_periodicidade_cal")
									dt_periodicidade_elet = rs_pat("dt_periodicidade_elet")
									last_cat = cd_categoria
									cd_categoria = rs_pat("cd_categoria")
									tipo_manut = rs_pat("tipo_manut")
									dt_plan_prev = rs_pat("dt_plan_prev")
									dt_vencimento = rs_pat("dt_vencimento")
										if str_tipo	= "serv_confirmadas" then
											cor_vencimento = "green"
										elseif str_tipo = "serv_vencer" then
											if int(ano_sel) < year(now) then 'Se o ano selecionado for menor que o atual mostra todas as manutençoes como vencidas
												cor_vencimento = "red"
											else
												if dt_vencimento < 0 then
													cor_vencimento = "red"
												elseif dt_vencimento = 0 then
													 cor_vencimento = "orange"
												elseif dt_vencimento = 1 then
													 cor_vencimento = "amarelo"
												elseif dt_vencimento = 2 then
													 cor_vencimento = "green"
												elseif dt_vencimento = 3 then
													 cor_vencimento = "#ccffff"
												elseif dt_vencimento >= 4 then
													 cor_vencimento = "white"
												end if
											End if
										end if
									cd_preventiva = rs_pat("cd_preventiva")
									
									'*** Monta a lista de patrimonios e categorias ***
										if int(cd_categoria) <> last_cat then
											sep_cat = ";"
										end if
									categoria = sep_cat&"cat"&cd_categoria&":"
									
									lista_num_pat = lista_num_pat&categoria&no_patrimonio&","
									lista_num_pat = replace(lista_num_pat,",#",",;#")
										
										if usuario = 4 then
											abre_pag = "patrimonio_cad"
											tamanho_janela = "600"
										else
											abre_pag = "patrimonio_consulta"
											tamanho_janela = "500"
										end if
										
										'abre_pag = "patrimonio_cad"%>
									<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
										<td align="center" style="background-color:<%=cor_vencimento%>;" class="nok_print"><%=zero(num_linha)%></td>
										<!--td align="center" style="background-color:white;" id="ok_print"><%=zero(num_linha)%></td-->
										<td onclick="window.open('patrimonio/<%=abre_pag%>.asp?tipo=cadastro&cd_pat_codigo=<%=cd_patrimonio%>&acao=alterar','Alterações','width=700,height=<%=tamanho_janela%>'); return false;"><%=nm_equipamento_novo%>-<%=cd_patrimonio%></td>
										<td align="center"><%'=cd_patrimonio&"-"%><%'=nm_tipo%><%=prefx_pat(no_patrimonio)%></td>
										<td align="center"><%=nm_marca%></td>
										<td align="center"><%=cd_pn%></td>
										<td align="center"><%=cd_ns%></td>
										<td align="center"><%if i = 1 then%><%=dt_periodicidade_prev%><%elseif i = 2 then%><%=dt_periodicidade_cal%><%elseif i = 3 then%><%=dt_periodicidade_elet%><%end if%> meses</td>
										<!--td><%'=cd_categoria%></td-->
										<td align="center"><span><%'=dt_vencimento%></span><%=mesdoano(month(dt_plan_prev))%></td>
										<%'xsql = "SELECT * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio='"&cd_patrimonio&"' AND dt_plan_prev <= '12/1/2013' AND tipo_manut='"&cd_categoria&"' order by dt_plan_prev desc"
										xsql = "SELECT * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio='"&cd_patrimonio&"' AND dt_plan_prev <= '12/1/2013' AND tipo_manut='"&i&"' order by dt_plan_prev desc"
										Set rs_agenda = dbconn.execute(xsql)
										
										if not rs_agenda.EOF then
											cod = rs_agenda("cd_codigo")
											mes_agendado = mesdoano(month(rs_agenda("dt_plan_prev")))
											tipo_m = rs_agenda("tipo_manut")
										else
											cod = ""
											mes_agendado = ""
											tipo_m = ""
										end if%>
										
										<td align="center"><%=mes_agendado%></td>
										<td class="no_print"><a href="#" title="confirma realização" javascript:void(0); onclick="patrimonio_manut_confirm(<%=cd_patrimonio%>,<%=tipo_manut%>,<%=cd_categoria%>,<%=cod_confirma%>,'1/<%=month(dt_plan_prev)&"/"&ano_sel%>')"><img src="../imagens/<%=img_confirma%>.gif" alt="" border="0" title="<%=tit_confirma%>" ></a><!-- cd_pat:<%=cd_patrimonio%> / man: <%=tipo_manut%> / cat<%=cd_categoria%> / <%=month(dt_plan_prev)&"/"&ano_sel%>--></td>
									</tr>
								<%if cd_categoria = 1 AND i = 1 then%><%cat_1_1 = cat_1_1 + 1%>
								<%elseif cd_categoria = 1 AND i = 2 then%><%cat_1_2 = cat_1_2 + 1%>
								<%elseif cd_categoria = 1 AND i = 3 then%><%cat_1_3 = cat_1_3 + 1%>
								
								<%elseif cd_categoria = 2 AND i = 1 then%><%cat_2_1 = cat_2_1 + 1%>
								<%elseif cd_categoria = 2 AND i = 2 then%><%cat_2_2 = cat_2_2 + 1%>
								<%elseif cd_categoria = 2 AND i = 3 then%><%cat_2_3 = cat_2_3 + 1%>
								
								<%elseif cd_categoria = 3 AND i = 1 then%><%cat_3_1 = cat_3_1 + 1%>
								<%elseif cd_categoria = 3 AND i = 2 then%><%cat_3_2 = cat_3_2 + 1%>
								<%elseif cd_categoria = 3 AND i = 3 then%><%cat_3_3 = cat_3_3 + 1%>
								
								<%elseif cd_categoria = 4 AND i = 1 then%><%cat_4_1 = cat_4_1 + 1%>
								<%elseif cd_categoria = 4 AND i = 2 then%><%cat_4_2 = cat_4_2 + 1%>
								<%elseif cd_categoria = 4 AND i = 3 then%><%cat_4_3 = cat_4_3 + 1%>
								
								<%elseif cd_categoria = 5 AND i = 1 then%><%cat_5_1 = cat_5_1 + 1%>
								<%elseif cd_categoria = 5 AND i = 2 then%><%cat_5_2 = cat_5_2 + 1%>
								<%elseif cd_categoria = 5 AND i = 3 then%><%cat_5_3 = cat_5_3 + 1%>
								
								<%elseif cd_categoria = 6 AND i = 1 then%><%cat_6_1 = cat_6_1 + 1%>
								<%elseif cd_categoria = 6 AND i = 2 then%><%cat_6_2 = cat_6_2 + 1%>
								<%elseif cd_categoria = 6 AND i = 3 then%><%cat_6_3 = cat_6_3 + 1%>
								
								<%elseif cd_categoria = 7 AND i = 1 then%><%cat_7_1 = cat_7_1 + 1%>
								<%elseif cd_categoria = 7 AND i = 2 then%><%cat_7_2 = cat_7_2 + 1%>
								<%elseif cd_categoria = 7 AND i = 3 then%><%cat_7_3 = cat_7_3 + 1%>
								
								<%elseif cd_categoria = 8 AND i = 1 then%><%cat_8_1 = cat_8_1 + 1%>
								<%elseif cd_categoria = 8 AND i = 2 then%><%cat_8_2 = cat_8_2 + 1%>
								<%elseif cd_categoria = 8 AND i = 3 then%><%cat_8_3 = cat_8_3 + 1%>
								
								<%elseif cd_categoria = 9 AND i = 1 then%><%cat_9_1 = cat_9_1 + 1%>
								<%elseif cd_categoria = 9 AND i = 2 then%><%cat_9_2 = cat_9_2 + 1%>
								<%elseif cd_categoria = 9 AND i = 3 then%><%cat_9_3 = cat_9_3 + 1%>
								
								<%elseif cd_categoria = 10 AND i = 1 then%><%cat_10_1 = cat_10_1 + 1%>
								<%elseif cd_categoria = 10 AND i = 2 then%><%cat_10_2 = cat_10_2 + 1%>
								<%elseif cd_categoria = 10 AND i = 3 then%><%cat_10_3 = cat_10_3 + 1%>
								
								<%elseif cd_categoria = 11 AND i = 1 then%><%cat_11_1 = cat_11_1 + 1%>
								<%elseif cd_categoria = 11 AND i = 2 then%><%cat_11_2 = cat_11_2 + 1%>
								<%elseif cd_categoria = 11 AND i = 3 then%><%cat_11_3 = cat_11_3 + 1%>
								
								<%elseif cd_categoria = 12 AND i = 1 then%><%cat_12_1 = cat_12_1 + 1%>
								<%elseif cd_categoria = 12 AND i = 2 then%><%cat_12_2 = cat_12_2 + 1%>
								<%elseif cd_categoria = 12 AND i = 3 then%><%cat_12_3 = cat_12_3 + 1%>
								
								<%elseif cd_categoria = 13 AND i = 1 then%><%cat_13_1 = cat_13_1 + 1%>
								<%elseif cd_categoria = 13 AND i = 2 then%><%cat_13_2 = cat_13_2 + 1%>
								<%elseif cd_categoria = 13 AND i = 3 then%><%cat_13_3 = cat_13_3 + 1%>
								
								<%elseif cd_categoria = 14 AND i = 1 then%><%cat_14_1 = cat_14_1 + 1%>
								<%elseif cd_categoria = 14 AND i = 2 then%><%cat_14_2 = cat_14_2 + 1%>
								<%elseif cd_categoria = 14 AND i = 3 then%><%cat_14_3 = cat_14_3 + 1%>
								
								<%elseif cd_categoria = 15 AND i = 1 then%><%cat_15_1 = cat_15_1 + 1%>
								<%elseif cd_categoria = 15 AND i = 2 then%><%cat_15_2 = cat_15_2 + 1%>
								<%elseif cd_categoria = 15 AND i = 3 then%><%cat_15_3 = cat_15_3 + 1%>
								<%end if%>
								
								<%num_linha = num_linha + 1
								end if
							'end if
							'cab_num = cab_num + 1
							ciclo_tempo = ciclo_tempo - 1
							sep_cat = ""
							dt_confirma = ""
							end if
					Next
			'wend%>
	<tr>
		<td colspan="10">&nbsp;</td>
	</tr>
	<%lista_patrimonio = ""
	cab_num = 0
	num_linha = 1
	next
end if%>
	<tr><td colspan="10"><b>TOTAL: <%=cat_1_1+cat_2_1+cat_3_1+cat_4_1+cat_5_1+cat_6_1+cat_7_1+cat_8_1+cat_9_1+cat_10_1+cat_11_1+cat_12_1+cat_13_1+cat_14_1+cat_15_1+cat_1_2+cat_2_2+cat_3_2+cat_4_2+cat_5_2+cat_6_2+cat_7_2+cat_8_2+cat_9_2+cat_10_2+cat_11_2+cat_12_2+cat_13_2+cat_14_2+cat_15_2+cat_1_3+cat_2_3+cat_3_3+cat_4_3+cat_5_3+cat_6_3+cat_7_3+cat_8_3+cat_9_3+cat_10_3+cat_11_3+cat_12_3+cat_13_3+cat_14_3+cat_15_3%></b></td></tr>
</form>
</table>
<!--p class="no_print"><%'=lista_num_pat%> </p--><br>
<table border="1" style="border:1px solid gray; border-collapse:collapse;" align="center">
	<tr><td align="center" colspan="4" style="background-color:gray; color:white;" >Listagem geral das manutenções</td></tr>
	<%if  cd_preventiva <> "" then%>		
		<tr><td colspan="4" style="background-color:silver;" align="center"><b>PREVENTIVA</b></td></tr>
		<tr style="background-color:#e5e5e5;"><td>&nbsp;</td><td><b>Item</b></td><td><b>Patrimonio</b></td><td><b>Qtd</b></td></tr>
		<%lista_num_pat = lista_num_pat&";"
		manut_1 = instr(1,lista_num_pat,"#m1",0)
		manut_2 = instr(1,lista_num_pat,"#m2",0)
		manut_3 = instr(1,lista_num_pat,"#m3",0)
		
		if manut_1 <> 0 then
			string_trabalho1 = lista_num_pat
			str_m1 = mid(string_trabalho1,manut_1, manut_2)
			len_m1 = len(str_m1)
			str_m1 = replace(str_m1,"#","")
		end if
		if manut_2 <> 0 then
			string_trabalho2 = mid(lista_num_pat,manut_2,manut_3)
			str_m2 = mid(string_trabalho2,1,manut_3-manut_2)
			len_m2 = len(str_m2)
			'str_m2 = replace(str_m2,"#","")
		end if
		if manut_3 <> 0 then
			string_trabalho3 = mid(lista_num_pat,manut_3,len(lista_num_pat))
			str_m3 = mid(string_trabalho3,1,len(lista_num_pat))
			len_m3 = len(str_m3)
			str_m3 = replace(str_m3,"#","")
		end if%>
		
		
		<!--tr><td colspan="4"><%=str_m1%></td></tr-->
		<%if cat_1_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 1)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat1"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_1_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		
		<%if cat_2_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 2)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat2"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_2_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_3_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 3)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat3"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_3_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_4_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 4)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat4"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_4_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_5_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 5)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat5"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_5_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_6_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 6)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat6"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_6_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_7_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 7)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat7"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_7_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_8_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 8)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat8"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_8_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_9_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 9)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat9"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_9_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_10_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 10)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat10"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_10_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_11_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 11)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat11"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td><br>
			<td align="center"><%=zero(cat_11_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_12_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 12)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat12"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_12_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_13_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 13)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat13"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_13_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_14_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 14)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat14"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td><br>
			<td align="center"><%=zero(cat_14_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_15_1 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 15)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat15"
			str_m = str_m1
			len_m = len_m1
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_15_1)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<tr><td colspan="2">&nbsp;</td>
		<td align="right"><b>Total:</b></td>
		<td align="center"><b><%=cat_1_1+cat_2_1+cat_3_1+cat_4_1+cat_5_1+cat_6_1+cat_7_1+cat_8_1+cat_9_1+cat_10_1+cat_11_1+cat_12_1+cat_13_1+cat_14_1+cat_15_1%></b></td></tr>
		<tr><td colspan="4">&nbsp;</td></tr>
		<%%>	
	<%end if%>
	
	<!--tr><td colspan="4"><%=str_m2%></td></tr-->	
	<%if  cd_preventiva <> "" then%>		
		<tr><td colspan="4" style="background-color:silver;" align="center"><b>CALIBRAÇÃO</b></td></tr>
		<tr style="background-color:#e5e5e5;"><td>&nbsp;</td><td><b>Item</b></td><td><b>Patrimonio</b></td><td><b>Qtd</b></td></tr>
		<%if cat_1_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 1)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat1"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_1_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		
		<%if cat_2_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 2)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat2"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_2_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_3_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 3)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat3"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_3_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_4_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 4)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat4"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_4_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_5_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 5)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat5"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_5_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_6_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 6)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat6"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_6_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_7_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 7)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat7"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_7_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_8_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 8)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat8"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_8_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_9_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 9)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat9"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_9_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_10_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 10)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat10"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_10_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_11_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 11)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat11"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_11_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_12_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 12)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat12"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_12_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_13_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 13)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat13"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td><br>
			<td align="center"><%=zero(cat_13_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_14_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 14)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat14"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_14_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_15_2 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 15)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat15"
			str_m = str_m2
			len_m = len_m2
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_15_2)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<tr><td colspan="2">&nbsp;</td>
		<td align="right"><b>Total:</b></td>
		<td align="center"><b><%=cat_1_2+cat_2_2+cat_3_2+cat_4_2+cat_5_2+cat_6_2+cat_7_2+cat_8_2+cat_9_2+cat_10_2+cat_11_2+cat_12_2+cat_13_2+cat_14_2+cat_15_2%></b></td></tr>
		<tr><td colspan="4">&nbsp;</td></tr>
		<%%>	
	<%end if%>
	
	<!--tr><td colspan="4"><%=str_m3%></td></tr-->
	<%if  cd_preventiva <> "" then%>	
		<tr><td colspan="4" style="background-color:silver;" align="center"><b>SEGURANÇA ELÉTRICA</b></td></tr>
		<tr style="background-color:#e5e5e5;"><td>&nbsp;</td><td><b>Item</b></td><td><b>Patrimonio</b></td><td><b>Qtd</b></td></tr>
		<%if cat_1_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 1)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat3"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_1_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_2_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 2)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat2"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_2_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_3_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 3)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat3"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_3_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_4_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 4)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat4"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_4_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_5_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 5)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat5"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_5_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_6_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 6)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat6"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_6_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_7_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 7)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat7"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_7_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_8_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 8)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat8"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_8_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_9_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 9)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat9"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_9_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_10_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 10)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat10"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_10_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_11_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 11)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat11"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_11_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_12_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 12)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat12"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_12_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_13_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 13)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat13"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_13_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_14_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 14)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat14"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_14_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<%if cat_15_3 <> "" Then%><%strsql = "SELECT * FROM TBL_Patrimonio_categoria WHERE (cd_codigo = 15)"%>
			<%Set rs_categoria = dbconn.execute(strsql)%><%if not rs_categoria.EOF Then%><%nm_categoria = rs_categoria("nm_categoria")%><%end if%>
			<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');"><td align="center"><b><%=zero(num_linha)%></b></td><td><%=nm_categoria%></td>
			<td><%categoria = "cat15"
			str_m = str_m3
			len_m = len_m3
			
			cati = instr(1,str_m,categoria,1)	'Inicio
			cats = mid(str_m,cati,len_m)	'nova String
			catf = instr(1,cats,";",1)	'Fim
				if int(catf) = 0 then
					cat_str = mid(cats,1,len(str_m))
				else
					cat_str = mid(cats,1,catf)
				end if
			cat_str = replace(cat_str,categoria&":"," ")
			cat_str = replace(cat_str,",;","")
			response.write(cat_str)%></td>
			<td align="center"><%=zero(cat_15_3)%></td></tr>
		<%num_linha = num_linha + 1
		end if%>
		<tr><td colspan="2">&nbsp;</td>
			<td align="right"><b>Total:</b></td>
			<td align="center"><b><%=cat_1_3+cat_2_3+cat_3_3+cat_4_3+cat_5_3+cat_6_3+cat_7_3+cat_8_3+cat_9_3+cat_10_3+cat_11_3+cat_12_3+cat_13_3+cat_14_3+cat_15_3%></b></td></tr>
		<tr><td colspan="4">&nbsp;</td></tr>
		<%%>	
	<%end if%>
	<tr><td colspan="4"><b>TOTAL GERAL: <%=cat_1_1+cat_2_1+cat_3_1+cat_4_1+cat_5_1+cat_6_1+cat_7_1+cat_8_1+cat_9_1+cat_10_1+cat_11_1+cat_12_1+cat_13_1+cat_14_1+cat_15_1+cat_1_2+cat_2_2+cat_3_2+cat_4_2+cat_5_2+cat_6_2+cat_7_2+cat_8_2+cat_9_2+cat_10_2+cat_11_2+cat_12_2+cat_13_2+cat_14_2+cat_15_2+cat_1_3+cat_2_3+cat_3_3+cat_4_3+cat_5_3+cat_6_3+cat_7_3+cat_8_3+cat_9_3+cat_10_3+cat_11_3+cat_12_3+cat_13_3+cat_14_3+cat_15_3%></b></td></tr>
	<tr>
		<td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
	</tr>
</table>	

</body>
</html>