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

strcd_unidade  =  request("cd_unidade")
ordem = request("ordem")
	if ordem = "" Then
	ordem = "nm_tipo, no_patrimonio"
	end if

mensagem = request("mensagem")
%>
<head>
	<title>Untitled</title>
</head>
<script language="JavaScript">
	 function grava_bkp(){
		 conteudo = document.gravarbkp.conteudo.value;
		 window.open('patrimonio/patrimonio_plan_serv.asp?conteudo='+conteudo+'','salva','width=400,height=150');
	 }
	 
	 function abre_bkp(){
		 window.open('patrimonio/patrimonio_abre.asp','abre','width=400,height=300');
	 }
</script>



<body>
<br id="no_print">
<%if strcd_unidade <> "" Then

for i = 1 to 3%>
<table class="txt" border="0" width="600" align="center" style="border-collapse:collapse;>	
				<%condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				'strsql ="up_patrimonio_lista @ordem= '"&ordem&"', @condicao = '"&condicao&"'"
			if i = 1 then
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' AND cd_preventiva = 1 order by cd_posicao, nm_equipamento"
			elseif i = 2 then
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' AND cd_calibracao = 1 order by cd_posicao, nm_equipamento"
			elseif i = 3 then
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' AND cd_seg_elet = 1 order by cd_posicao, nm_equipamento"
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
				cd_unidade = rs_patrimonio("cd_unidade")
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
		<td colspan="12" align="center" valign="bottom" style="font-family:arial; font-size:12px; color:black;"><b><%=Ucase(tipo_servico)%>, <%=ano_sel%></b></td>
	</tr>
	<tr><td colspan="1000"><img src="../imagens/px.gif" alt="" width="10" height="25" border="0"></td></tr>
	
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');">
		
		<td align="center" valign="bottom" style="<%=estilo_corpo%>">&nbsp;<%=i%></td>
		<td valign="top" align="left" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<a href="javascript:viod(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar','Alterações','width=700','height=600'); return false;"><%=nm_equipamento%></a> <%'=dt_periodicidade_prev%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');" style="<%=estilo_corpo%>">&nbsp;<%="0"&no_patrimonio%></td>				
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
							'	if garantia = dt_selecionada then%>
									<td style="<%=estilo_corpo%>">++</td>
								<%'else
									mes_servico = mes_atribuido%>
									<td style="<%=estilo_corpo%>"><%=mes_servico%></td>
									<td style="<%=estilo_corpo%>"><%=ano_sel%></td>
								<%'end if%>								
							<%end if%>					
								
							<!--td style="<%=estilo_corpo%>"><%=marca&"-"&mes_servico%></td-->
						<%'mes_servico = ""
						next%>
						
		
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
						
	<%end if%>
	</tr>
		
	<%gravar_txt = gravar_txt&"$"&i&";"&nm_equipamento&";"&no_patrimonio&";"&nm_marca&";"&cd_pn&";"&dt_periodicidade&";"&mes_servico&";"&ano_sel
	
	dt_plan_prev_mes = ""
	dt_plan_prev_ano = ""
	meses = ""
	var = ""
	
	
	dt_plan_prev = ""
	cabeca = 1
	nm_especialidade = ""
	num_lista = num_lista + 1
	rs_patrimonio.movenext
	loop%>	
	<tr>
		<td>&nbsp;</td>
		<td colspan="17">
			<%strsql ="SELECT * FROM TBL_patrimonio_manut_obs WHERE cd_unidade="&cd_unidade&" AND tipo_manut = "&i&" AND dt_obs = '1/1/"&ano_sel&"' "
			Set rs_prev = dbconn.execute(strsql)
				if not rs_prev.EOF Then
					cd_obs = rs_prev("cd_codigo")%>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manut_obs.asp?cd_obs=<%=cd_obs%>','nome','width=400,height=230'); return false;>&nbsp;<%=rs_prev("nm_obs")%></a><img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="javascript:JsDelObs('<%=cd_obs%>')" id="no_print">
				<%else%>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<a href=javascript:viod(0); onclick=window.open('patrimonio/patrimonio_cad_manut_obs.asp?ano_atribuido=<%=ano_sel%>&tipo_manut=<%=i%>&cd_unidade=<%=cd_unidade%>','nome','width=400,height=230'); return false;>&nbsp;<img src='../imagens/ic_novo.gif' border='0' id="no_print"></a>
				<%end if%>
		</td>
	</tr>
	<tr>		
		<td>&nbsp;</td>
		<td align="left" colspan="14"></td>
		<td align="center" colspan="3"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>	


</table><%if i <= 2 then%><!--p style="page-break-after:always;"></p--><%'end if%>
<div style="page-break-after:always;">&nbsp;</div><%end if%>
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
<tr><td colspan="17"><%=replace(gravar_txt,"$","<br>")%></td></tr>
	
<form  name="gravarbkp">
	<input type="hidden" name="conteudo" value="<%=gravar_txt%>"  id="no_print">
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
<script language="JavaScript">
	 function grava_bkp(){
		 conteudo = document.gravarbkp.conteudo.value;
		 window.open('patrimonio/patrimonio_grava.asp?conteudo='+conteudo+'','nome','width=400,height=250');
	 }
</script>
</body>
</html>
