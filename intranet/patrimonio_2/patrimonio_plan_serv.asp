<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_pat_codigo = request("cd_pat_codigo")

mes_sel = request("mes_sel")
	if mes_sel = "" Then
		mes_sel = year(now)
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
	<title>Untitled</title>
</head>
<script language="JavaScript">
function grava_bkp(save){
		save=save;	 
		var el=document.getElementById(save);
		el.style.display=(el.style.display!='block'?'block':'');
	}
function abre_bkp(){
		 window.open('patrimonio/patrimonio_plan_serv_abre.asp','abre','width=400,height=300');
	 }
	 	
function cancela(ano,unidade){
		document.location.href('http://server/patrimonio.asp?ano_sel='+ano+'&cd_unidade='+unidade+'&tipo=plan_serv');	
	}
	
//function validar_salvar){
	//if (shipinfo.cd_regime_trabalho.value.length==""){window.alert ("O item Contrato deve ser preenchido.");shipinfo.cd_regime_trabalho.focus();return (false);}
	//return (true);
	//}
function validar_busca(shipinfo){
	if (shipinfo.ano_sel.value.length==""){window.alert ("Informe o ano.");shipinfo.ano_sel.focus();return (false);}
	if (shipinfo.cd_unidade.value.length==""){window.alert ("Informe a unidade.");shipinfo.cd_unidade.focus();return (false);}	
	return (true);}
</script>



<body>
<br id="no_print">
<table id="no_print" style="border:1 solid green; left:150px;" align="center">
	<form action="patrimonio.asp" method="post" name="form" onsubmit="return validar_busca(document.form);">
	<input type="hidden" name="tipo" value="plan_serv">
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">patrimonio_plan_serv.asp</span></td></tr>
	<%end if%>	
	<tr style="border:1 solid orange; background-color:silver;">
		<td colspan="100%" align="center"><B>PLANEJAMENTO DE SERVIÇOS</B></td>
	</tr>
	<tr>
		<td>&nbsp; Ano</td>
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
	<tr>		
		<td><input type="submit" name="ok" value="ok"></td>
		<td><img src="../imagens/ic_print.gif" alt="Imprimir" width="24" height="24" border="0" onClick="window.print();"> &nbsp; 
		<img src="../imagens/ic_print_view.gif" alt="Visualizar" width="24" height="26" border="0" onclick="visualizarImpressao();"> &nbsp; 
		<%if strcd_unidade <> "" Then%>
		<img src="../imagens/ic_save.gif" alt="Salvar" width="24" height="26" border="0" onclick="grava_bkp('save');"><%end if%> &nbsp;
		<img src="../imagens/ic_open.gif" alt="Abrir" width="24" height="26" border="0" onclick="abre_bkp();"></td>
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
</table><br id="no_print">
<%if strcd_unidade <> "" Then

	'*** 28/06/11 - a unidade EDV não faz Seg. Eletrica ***
	if strcd_unidade = 20 then
		qtd_prevent = 2
	else
		qtd_prevent = 3
	end if


for i = 1 to qtd_prevent%>
<%if i <= 2 then%>
	<table class="txt" border="0" width="800" align="center" style="border-collapse:collapse; page-break-after:always;">	
<%else%>
	<table class="txt" border="0" width="800" align="center" style="border-collapse:collapse;">
<%end if%>
				<%condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				'strsql ="up_patrimonio_lista @ordem= '"&ordem&"', @condicao = '"&condicao&"'"
			if i = 1 then
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade_comodato = "&strcd_unidade&" AND nm_tipo = 'E' AND cd_preventiva = 1 order by cd_posicao, nm_equipamento_novo"
			elseif i = 2 then
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade_comodato = "&strcd_unidade&" AND nm_tipo = 'E' AND cd_calibracao = 1 order by cd_posicao, nm_equipamento_novo"
			elseif i = 3 then
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade_comodato = "&strcd_unidade&" AND nm_tipo = 'E' AND cd_seg_elet = 1 order by cd_posicao, nm_equipamento_novo"
			end if
			
			  	Set rs_patrimonio = dbconn.execute(strsql)
				
				Do while not rs_patrimonio.EOF
				cd_pat_codigo = rs_patrimonio("cd_codigo")
				no_patrimonio = rs_patrimonio("no_patrimonio")
				cd_equipamento = rs_patrimonio("cd_equipamento")
				cd_posicao = rs_patrimonio("cd_posicao")
				nm_equipamento = rs_patrimonio("nm_equipamento_novo")
				cd_marca = rs_patrimonio("cd_marca")
				nm_marca = rs_patrimonio("nm_marca")
				cd_ns = rs_patrimonio("cd_ns")
				cd_pn = rs_patrimonio("cd_pn")
				nm_tipo = rs_patrimonio("nm_tipo")
				cd_especialidade = rs_patrimonio("cd_especialidade")
				cd_unidade = rs_patrimonio("cd_unidade_comodato")
					strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" AND cd_status <> 0"
					Set rs_uni = dbconn.execute(strsql)
					while not rs_uni.EOF
						nm_unidade = rs_uni("nm_unidade")
					rs_uni.movenext
					wend
					
					
				nm_sigla = rs_patrimonio("nm_sigla")
				cd_rack = rs_patrimonio("cd_rack")
				num_hospital =  rs_patrimonio("num_hospital")
				cd_rack = rs_patrimonio("cd_rack")
				num_hospital = rs_patrimonio("num_hospital")
				dt_data = rs_patrimonio("dt_data")
				
				cd_preventiva = rs_patrimonio("cd_preventiva")
				dt_periodicidade_prev = rs_patrimonio("dt_periodicidade_prev")
				
				cd_calibracao = rs_patrimonio("cd_calibracao")
				dt_periodicidade_cal = rs_patrimonio("dt_periodicidade_cal")
				
				cd_seg_elet = rs_patrimonio("cd_seg_elet")
				dt_periodicidade_elet = rs_patrimonio("dt_periodicidade_elet")
				
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
					strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"' "
				  	Set rs_espec = dbconn.execute(strsql_espec)
						if not rs_espec.EOF then
							nm_especialidade = rs_espec("nm_especialidade")
						end if
				
				if i = 1 then
					tipo_servico = "preventiva"
					dt_periodicidade = dt_periodicidade_prev
				elseif i = 2 then
					tipo_servico = "calibraçao"
					dt_periodicidade = dt_periodicidade_cal
				elseif i = 3 then
					tipo_servico = "segurança elétrica"
					dt_periodicidade = dt_periodicidade_elet
				end if
				
				
		estilo_titulo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:12px; color:black;"
		estilo_corpo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:12px;"%>
	
		
		
	<%if cabeca = 0 then%>
	<tr>
		<!--td><img src="../imagens/blackdot.gif" alt="" width="10" height="15" border="0"></td-->
		<td id="ok_print"><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>	
		<td id="no_print"><img src="../imagens/px.gif" alt="" width="10" height="5" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="180" height="5" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
	</tr>
	<tr>		
		<td align="center">&nbsp;</td>
		<td  valign="bottom">
			<%if cd_unidade = 2 or cd_unidade = 3 or cd_unidade = 15 then%><img src="../imagens/logotipo_sao_luiz.gif" alt="" width="120" height="37" border="0"><%end if%>
			</td>
		<td colspan="12" align="center" valign="bottom" style="font-family:arial; font-size:12px; color:black;"><b>PLANEJAMENTO DE SERVIÇOS REFERENTE A <%=Ucase(tipo_servico)%> ANO <%=ano_sel%><br>VDLAP</b></td>
		<td colspan="4" align="right" valign="bottom"><img src="../imagens/logotipo_vdlap.gif" alt="" width="100" height="40" border="0"></td>
	</tr>
	<tr><td align="center">&nbsp;</td><td align="center">&nbsp;</td><td colspan="12" align="center" style="font-size:12px;"><%if cd_unidade = 20 then%><%=nm_unidade%><%end if%></td></tr>
	<tr><td colspan="1000"><img src="../imagens/px.gif" alt="" width="10" height="20" border="0"></td></tr>
	<tr>
		<td align="center" valign="bottom">&nbsp;</td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>EQUIPAMENTO</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>CONTROLE DE<br>MANUTENÇÃO</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>MARCA</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>MODELO</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>PERIODICIDADE</b></td>		
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>JAN</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>FEV</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>MAR</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>ABR</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>MAI</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>JUN</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>JUL</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>AGO</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>SET</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>OUT</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>NOV</b></td>
		<td align="center" valign="bottom" style="<%=estilo_corpo%>"><b>DEZ</b></td>
	</tr>
	<%end if%>
	
	<%data_agora = zero(month(now))&"/"&zero(day(now))&"/"&year(now)%>
			
	<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');">
		<td align="center" valign="bottom">&nbsp;</td>
		<td valign="top" align="left" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<a href="javascript:viod(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar','Alterações','width=700,height=600,scrollbars=1'); return false;"><%=nm_equipamento%></a> <%'=dt_periodicidade_prev%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=prefx_pat(no_patrimonio)%></td>				
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=nm_marca%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=cd_pn%></td>
	<%qtd_meses = 12
	repeticoes = 120
	
	if i = 1 then%>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=dt_periodicidade_prev%> <%if dt_periodicidade_prev <> "" Then response.write("meses") end if%></td>
		<%'*** Resgata o ultimo registro ***
		strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = "&i&" AND (dt_plan_prev <= '12/31/"&ano_sel&"') order by dt_plan_prev desc"
			Set rs_prev = dbconn.execute(strsql)
				'while not rs_prev.EOF
				if not rs_prev.EOF then
						cd_plan_prev = rs_prev("cd_codigo")
						dt_plan_prev = rs_prev("dt_plan_prev")
						dt_plan_prev_mes = month(dt_plan_prev)
						dt_plan_prev_ano = year(dt_plan_prev)
						dt_garantia = rs_prev("dt_garantia")
				
					diferenca_anos = int(ano_sel) - dt_plan_prev_ano
				
					if diferenca_anos = 0 then
						multi_soma = qtd_meses
						monta_grade_i = 1
						monta_grade_f = qtd_meses
					
					elseif diferenca_anos >= 1 then
						multi_soma = repeticoes / dt_periodicidade_prev'
						monta_grade_i = (qtd_meses)*diferenca_anos + 1
						monta_grade_f = qtd_meses + monta_grade_i - 1
					end if
				
						if int(ano_sel) = dt_plan_prev_ano then '***
							'acao = "editar"
							var_plan_prev = "cd_plan_prev="&cd_plan_prev&"&acao=editar"
							mes_ultima = dt_plan_prev_mes
								for imarca = 1 to multi_soma
									if imarca = 1 then
										meses = dt_periodicidade_prev + mes_ultima
									Else
										meses = dt_periodicidade_prev + meses
									end if
									var = var&";"&meses
								next
								meses = ";"&mes_ultima&var&";"
						
						elseif int(ano_sel) > dt_plan_prev_ano then 
							'acao = "inserir"
							var_plan_prev = "cd_plan_prev="&cd_pat_codigo&"&acao=inserir"
							mes_ultima = dt_plan_prev_mes
								for imarca = 1 to multi_soma
									if imarca = 1 then
										meses = dt_periodicidade_prev + mes_ultima
									Else
										meses = dt_periodicidade_prev + meses
									end if
									var = var&";"&meses
								next
								meses = ";"&mes_ultima&var&";"
						end if
				else
					var_plan_prev = "cd_plan_prev="&cd_pat_codigo&"&acao=inserir"
					meses = ""
				'end if
				
				end if
				'rs_prev.movenext
				'wend
						
						
						
						
						'*** Marca a célula da preventiva ***
						mes_atribuido = 0
							if monta_grade_f < 12 then
								monta_grade_f = 11
							end if
						for imes = monta_grade_i to monta_grade_f
							mes_atribuido = mes_atribuido + 1
							
							dt_selecionada = ano_sel&zero(mes_atribuido)&"01"
							garantia = year(dt_garantia)&zero(month(dt_garantia))&"01"
							if instr(1,meses,";"&imes&";",1) <> 0 then								
								if garantia = dt_selecionada then
									marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&mes_atribuido="&mes_atribuido&"&ano_atribuido="&ano_sel&"&tipo_manut="&i&"&"&var_plan_prev&"','nome','width=400,height=250'); return false;><img src='../imagens/x0.gif' height=15 border='0'></a>"
								else
									marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&mes_atribuido="&mes_atribuido&"&ano_atribuido="&ano_sel&"&tipo_manut="&i&"&"&var_plan_prev&"','nome','width=400,height=250'); return false;><img src='../imagens/x1.gif' height=15 border='0'></a>"
									mes_servico = mes_atribuido
								end if
							else
								marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&mes_atribuido="&mes_atribuido&"&ano_atribuido="&ano_sel&"&tipo_manut="&i&"&"&var_plan_prev&"','nome','width=400,height=250'); return false;><img src='../imagens/x0.gif' height=15 border='0'></a>"
							end if%>					
								
							<td style="<%=estilo_corpo%>"><%=marca%></td>
						<%next
						'gravar_txt = gravar_txt&"#"&i&";"&nm_equipamento&";"&no_patrimonio&";"&nm_marca&";"&cd_pn&";"&dt_periodicidade&";"&mes_servico%>
						<!--td><%=cd_pat_codigo%><%'=monta_grade_f&"-"&diferenca_anos&"/"&meses%></td-->
		
	<%elseif i = 2 then%>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=dt_periodicidade_cal%> <%if dt_periodicidade_cal <> "" Then response.write("meses") end if%></td>
		<%'*** Resgata o ultimo registro ***
		strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = "&i&" AND (dt_plan_prev <= '12/31/"&ano_sel&"') order by dt_plan_prev desc"
			Set rs_prev = dbconn.execute(strsql)
				'while not rs_prev.EOF
				if not rs_prev.EOF then
						cd_plan_prev = rs_prev("cd_codigo")
						dt_plan_prev = rs_prev("dt_plan_prev")
						dt_plan_prev_mes = month(dt_plan_prev)
						dt_plan_prev_ano = year(dt_plan_prev)
						dt_garantia = rs_prev("dt_garantia")
				
					diferenca_anos = int(ano_sel) - dt_plan_prev_ano
				
					if diferenca_anos = 0 then
						multi_soma = qtd_meses
						monta_grade_i = 1
						monta_grade_f = qtd_meses
					
					elseif diferenca_anos >= 1 then
						multi_soma = repeticoes / dt_periodicidade_prev'
						monta_grade_i = (qtd_meses)*diferenca_anos + 1
						monta_grade_f = qtd_meses + monta_grade_i - 1
					end if
				
						if int(ano_sel) = dt_plan_prev_ano then '***
							'acao = "editar"
							var_plan_prev = "cd_plan_prev="&cd_plan_prev&"&acao=editar"
							mes_ultima = dt_plan_prev_mes
								for imarca = 1 to multi_soma
									if imarca = 1 then
										meses = dt_periodicidade_prev + mes_ultima
									Else
										meses = dt_periodicidade_prev + meses
									end if
									var = var&";"&meses
								next
								meses = ";"&mes_ultima&var&";"
						
						elseif int(ano_sel) > dt_plan_prev_ano then 
							'acao = "inserir"
							var_plan_prev = "cd_plan_prev="&cd_pat_codigo&"&acao=inserir"
							mes_ultima = dt_plan_prev_mes
								for imarca = 1 to multi_soma
									if imarca = 1 then
										meses = dt_periodicidade_prev + mes_ultima
									Else
										meses = dt_periodicidade_prev + meses
									end if
									var = var&";"&meses
								next
								meses = ";"&mes_ultima&var&";"
						end if
				else
					var_plan_prev = "cd_plan_prev="&cd_pat_codigo&"&acao=inserir"
					meses = ""
				'end if
				
				end if
				'rs_prev.movenext
				'wend
						'*** Marca a célula da preventiva ***
						mes_atribuido = 0
							if monta_grade_f < 12 then
								monta_grade_f = 11
							end if
						for imes = monta_grade_i to monta_grade_f
							mes_atribuido = mes_atribuido + 1
							
							dt_selecionada = ano_sel&zero(mes_atribuido)&"01"
							garantia = year(dt_garantia)&zero(month(dt_garantia))&"01"
							
								if instr(1,meses,";"&imes&";",1) <> 0 then
									if garantia = dt_selecionada then
										marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&mes_atribuido="&mes_atribuido&"&ano_atribuido="&ano_sel&"&tipo_manut="&i&"&"&var_plan_prev&"','nome','width=400,height=250'); return false;><img src='../imagens/x0.gif' height=15 border='0'></a>"
									else
										marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&mes_atribuido="&mes_atribuido&"&ano_atribuido="&ano_sel&"&tipo_manut="&i&"&"&var_plan_prev&"','nome','width=400,height=250'); return false;><img src='../imagens/x1.gif' height=15 border='0'></a>"
										mes_servico = mes_atribuido
									end if
								else
									marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&mes_atribuido="&mes_atribuido&"&ano_atribuido="&ano_sel&"&tipo_manut="&i&"&"&var_plan_prev&"','nome','width=400,height=250'); return false;><img src='../imagens/x0.gif' height=15 border='0'></a>"
								end if%>					
								
							<td style="<%=estilo_corpo%>"><%=marca%></td>
						<%next%>
						<!--td><%=cd_pat_codigo%><%'=monta_grade_f&"-"&diferenca_anos&"/"&meses%></td-->
	<%elseif i = 3 then%>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%=dt_periodicidade_elet%> <%if dt_periodicidade_elet <> "" Then response.write("meses") end if%></td>
		<%'*** Resgata o ultimo registro ***
		strsql ="SELECT top 1 * FROM TBL_patrimonio_manutencoes WHERE cd_patrimonio="&cd_pat_codigo&" and cd_suspenso <> 1 AND tipo_manut = "&i&" AND (dt_plan_prev <= '12/31/"&ano_sel&"') order by dt_plan_prev desc"
			Set rs_prev = dbconn.execute(strsql)
				'while not rs_prev.EOF
				if not rs_prev.EOF then
						cd_plan_prev = rs_prev("cd_codigo")
						dt_plan_prev = rs_prev("dt_plan_prev")
						dt_plan_prev_mes = month(dt_plan_prev)
						dt_plan_prev_ano = year(dt_plan_prev)
						dt_garantia = rs_prev("dt_garantia")
				
					diferenca_anos = int(ano_sel) - dt_plan_prev_ano
				
					if diferenca_anos = 0 then
						multi_soma = qtd_meses
						monta_grade_i = 1
						monta_grade_f = qtd_meses
					
					elseif diferenca_anos >= 1 then
						multi_soma = repeticoes / dt_periodicidade_prev'
						monta_grade_i = (qtd_meses)*diferenca_anos + 1
						monta_grade_f = qtd_meses + monta_grade_i - 1
					end if
				
						if int(ano_sel) = dt_plan_prev_ano then '***
							'acao = "editar"
							var_plan_prev = "cd_plan_prev="&cd_plan_prev&"&acao=editar"
							mes_ultima = dt_plan_prev_mes
								for imarca = 1 to multi_soma
									if imarca = 1 then
										meses = dt_periodicidade_prev + mes_ultima
									Else
										meses = dt_periodicidade_prev + meses
									end if
									var = var&";"&meses
								next
								meses = ";"&mes_ultima&var&";"
						
						elseif int(ano_sel) > dt_plan_prev_ano then 
							'acao = "inserir"
							var_plan_prev = "cd_plan_prev="&cd_pat_codigo&"&acao=inserir"
							mes_ultima = dt_plan_prev_mes
								for imarca = 1 to multi_soma
									if imarca = 1 then
										meses = dt_periodicidade_prev + mes_ultima
									Else
										meses = dt_periodicidade_prev + meses
									end if
									var = var&";"&meses
								next
								meses = ";"&mes_ultima&var&";"
						end if
				else
					var_plan_prev = "cd_plan_prev="&cd_pat_codigo&"&acao=inserir"
					meses = ""
				'end if
				
				end if
				'rs_prev.movenext
				'wend
						'*** Marca a célula da preventiva ***
						mes_atribuido = 0
							if monta_grade_f < 12 then
								monta_grade_f = 11
							end if
						for imes = monta_grade_i to monta_grade_f
							mes_atribuido = mes_atribuido + 1
							
							dt_selecionada = ano_sel&zero(mes_atribuido)&"01"
							garantia = year(dt_garantia)&zero(month(dt_garantia))&"01"
							
								if instr(1,meses,";"&imes&";",1) <> 0 then
									if garantia = dt_selecionada then
										marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&mes_atribuido="&mes_atribuido&"&ano_atribuido="&ano_sel&"&tipo_manut="&i&"&"&var_plan_prev&", jan=1','nome','width=400,height=250'); return false;><img src='../imagens/x0.gif' height=15 border='0'></a>"
									else
										marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&mes_atribuido="&mes_atribuido&"&ano_atribuido="&ano_sel&"&tipo_manut="&i&"&"&var_plan_prev&", jan=1','nome','width=400,height=250'); return false;><img src='../imagens/x1.gif' height=15 border='0'></a>"
										mes_servico = mes_atribuido
									end if
								else
									marca = "<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manutencoes.asp?tipo=cadastro&mes_atribuido="&mes_atribuido&"&ano_atribuido="&ano_sel&"&tipo_manut="&i&"&"&var_plan_prev&", jan=1','nome','width=400,height=250'); return false;><img src='../imagens/x0.gif' height=15 border='0'></a>"
								end if%>					
								
							<td style="<%=estilo_corpo%>"><%=marca%></td>
						<%next%>
						<!--td><%=cd_pat_codigo%><%'=monta_grade_f&"-"&diferenca_anos&"/"&meses%></td-->
						
	<%end if%>
	
	</tr>

		
	<%conteudo = conteudo&"$"&i&";"&strcd_unidade&";"&nm_equipamento&";"&no_patrimonio&";"&rtrim(nm_marca)&";"&cd_pn&";"&dt_periodicidade&";"&mes_servico&";"&ano_sel
	dt_plan_prev_mes = ""
	dt_plan_prev_ano = ""
	meses = ""
	var = ""
	cd_emprestimo = ""
	
	dt_plan_prev = ""
	cabeca = 1
	nm_especialidade = ""
	num_lista = num_lista + 1
	rs_patrimonio.movenext
	loop%>
	<%if cd_unidade = 2 or cd_unidade = 3 or cd_unidade = 15 then%>
	<tr>
		<td align="center" valign="bottom">&nbsp;</td>
		<td colspan="5">ITA20211.PQ.006 - ANEXO 3 - PLANEJAMENTO DE SERVIÇOS - VDLAP<p>OBS.: As datas marcadas são referentes a <%=tipo_servico%></p></td>
	</tr>
	<%conteudo = conteudo&"$"&i&";"&strcd_unidade&";0;0;ITA20211.PQ.006 - ANEXO 3 - PLANEJAMENTO DE SERVIÇOS - VDLAP<p>OBS.: As datas marcadas são referentes a "&tipo_servico&";0;0;0;0" 
	end if%>
	<tr>
		<td>&nbsp;</td>
		<td colspan="14">
			<%if cd_unidade = "" Then
			cd_unidade = 0
			end if
			
			strsql ="SELECT * FROM TBL_patrimonio_manut_obs WHERE cd_unidade="&cd_unidade&" AND tipo_manut = "&i&" AND dt_obs = '1/1/"&ano_sel&"' "
			Set rs_prev = dbconn.execute(strsql)
				if not rs_prev.EOF Then
					cd_obs = rs_prev("cd_codigo")
					nm_obs = rs_prev("nm_obs")%>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manut_obs.asp?cd_obs=<%=cd_obs%>&jan=1','nome','width=400,height=230'); return false;>&nbsp;<%=nm_obs%></a><img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="javascript:JsDelObs('<%=cd_obs%>')" id="no_print">
					<%conteudo = conteudo&"$"&i&";"&strcd_unidade&";0;0;"&nm_obs&";0;0;0;0"
				else%>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manut_obs.asp?ano_atribuido=<%=ano_sel%>&tipo_manut=<%=i%>&cd_unidade=<%=cd_unidade%>&jan=1','nome','width=400,height=230'); return false;>&nbsp;<img src='../imagens/ic_novo.gif' border='0' id="no_print"></a>
				<%end if%></td>
		<td align="center" colspan="3"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>	
</table><br id="no_prints">
<%cabeca = 0
next%>
	<!--form  id="no_print">
	<table id="no_print">
		<tr>
			<td colspan="17">
				<%'gravar_txt = replace(gravar_txt,"#","<br>")%><%'gravar_txt = mid(gravar_txt,2,len(gravar_txt))%>
				<textarea cols="200" rows="4" name="gravar_txt"><%'=gravar_txt%></textarea>
			</td>
		</tr>
	</table>
	</form-->
<%end if%>
<form  name="gravarbkp" id="no_print">
	<input type="hidden" name="ano_bkp" value="<%=ano_sel%>"  id="no_print">
	<input type="hidden" name="unidade_bkp" value="<%=cd_unidade%>"  id="no_print">
</form>
		
<SCRIPT LANGUAGE="javascript">
function JsDelObs(cod1)
	{
	  if (confirm ("Tem certeza que deseja excluir a observação?"))
		  {
			document.location.href('patrimonio/acoes/patrimonio_acao.asp?cd_codigo='+cod1+'&cd_tipo=manut_obs&acao=apagar');
		  }
	}
</SCRIPT>

<%if ok = "Salvar" Then
	origem = zero(1)
	dt_gravacao = right(year(now),2)&zero(month(now))&zero(day(now))&zero(hour(now))&zero(minute(now))
	usuario =  session("cd_codigo")
	
	nome_arquivo = request("nm_bkp")
	nome_gravacao = origem&"_"&zero(strcd_unidade)&"_"&dt_gravacao&"_"&zero(usuario)&"-"&nome_arquivo
	grava_txt = conteudo'request("gravar_txt")
	
	dt_hoje = month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)
	
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	
	caminho = Server.MapPath("gravados/"&nome_gravacao&".txt") 'especifique aqui o caminho onde ficará/está o TXT
	Set GRAVAR = FSO.CreateTextFile(caminho,true)
	'Foi criado o objeto e logo após busca o txt em caminho para gravar, se não achar, vai cria-lo (por causa da marcação TRUE)
		
		'*** grava no banco de dados ***
		'if dt_plan_mes_prev <> "" AND dt_plan_ano_prev <> "" AND cd_preventiva  <> "0" Then
			xsql = "up_grava_bkp_insert @cd_origem="&origem&", @cd_unidade="&strcd_unidade&", @dt_gravacao='"&month(now)&"/"&day(now)&"/"&year(now)&" "&hour(now)&":"&minute(now)&":"&second(now)&"', @cd_usuario='"&usuario&"', @nm_arquivo='"&nome_gravacao&"'"
			dbconn.execute(xsql)
		'end if
		
	gravar.write (conteudo)
	gravar.close
	'response.write "GRAVADO!"&nome_gravacao&"***"&conteudo&""
	'apos abrir o TXT, gravará a linha com o texto "TESTE DE GRAVAÇÃO" a confirmação no cliente aparecerá como "GRAVADO"
	response.redirect("patrimonio.asp?ano_sel="&ano_sel&"&cd_unidade="&strcd_unidade&"&tipo=plan_serv&mens=planilha+salva+com+sucesso!")
end if%>

<%'=nome_arquivo%>
</body>
</html>
