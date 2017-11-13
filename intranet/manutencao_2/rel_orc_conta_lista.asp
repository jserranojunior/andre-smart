<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<!-- #include file="../css/geral.htm"-->
<body onunload="window.opener.location.reload(true);">

<%cd_user = session("cd_codigo")
pasta_loc = "manutencao_2\"
arquivo_loc = "rel_orc_conta_lista.asp"

tipo = request("tipo")

dt_mes = month(now)
dt_ano = year(now)

		largura_relatorio = 700
		qtd_linhas_pagina = 65
		num_linha = 1
		num = 1
ordem_sel = 2

%>


<!--#include file="../includes/arquivo_loc.asp"-->


<table align="center" border="1" style="border-collapse:collapse; width:<%=largura_relatorio%>px;">
	<%
	maior_qtd_parcelas = 0
	cancela_subtotal = 0
	ultima_dt_entrega = NULL
		if cd_user = 46  then
			limite_i = 0
			limite_f = 4
		else
			limite_i = 3
			limite_f = 4
		end if
		
	compromissos = 4
	aprovados_recentes = 3
	pgto_suspensos = 2
	sem_aprovacao = 1
	erro_processo = 0
	
	qtd_meses_mostrar = 6
	
	'strsql ="sp_orcamento_s01 @funcao=S, @dt_inicial='"&dt_ano&"-"&dt_mes&"-1', @dt_final='"&dt_ano&"-"&dt_mes&"-"&dia_final&"', @limite_i="&limite_i&", @limite_f="&limite_f&""
	dt_ano_fim = dt_ano
	dt_mes_fim = dt_mes + 1
		if dt_mes_fim > 12 then
			dt_mes_fim = 1
			dt_ano_fim = dt_ano + 1
		end if
	dia_final = ultimodiames(dt_mes_fim,dt_ano_fim)
	total_valores = 0
	
	strsql ="sp_orcamento_s01 @funcao=SS"
		Set rs = dbconn.execute(strsql)
			while not rs.EOF
				cd_orcamento = rs("cd_codigo")
				cd_fornecedor = rs("cd_fornecedor")
				nm_fornecedor = rs("nm_fornecedor")
				nr_orcamento = rs("nr_orcamento")
				'dt_orcamento = rs("dt_orcamento")
				dt_aprov_orc = rs("dt_aprov_orc")
				dt_entrega = rs("dt_entrega")
					if isnull(dt_entrega) Then
						dt_entrega = NULL
					else
						dt_entrega = zero(day(dt_entrega))&"/"&zero(month(dt_entrega))&"/"&right(year(dt_entrega),2)
					end if
				nm_valor = rs("nm_valor")
				'	total_nm_valor = abs(total_nm_valor) + abs(nm_valor)
					if nm_valor <> "" Then
						total_valores = total_valores + nm_valor
					end if
				nr_parcela = rs("nr_parcela")
					if nr_parcela >= 1 and nm_valor <> "" then
						nr_valor_parcela = nm_valor / nr_parcela
					else
						nr_valor_parcela = 0
					end if
				dt_inicio_suspensao = rs("dt_inicio_suspensao")
					if not ISNULL(dt_inicio_suspensao) Then dt_inicio_suspensao = zero(day(dt_inicio_suspensao))&"/"&zero(month(dt_inicio_suspensao))&"/"&right(year(dt_inicio_suspensao),2)
						
				dt_fim_suspensao = rs("dt_fim_suspensao")
					if not ISNULL(dt_fim_suspensao) Then 
						dt_fim_suspensao = zero(day(dt_fim_suspensao))&"/"&zero(month(dt_fim_suspensao))&"/"&right(year(dt_fim_suspensao),2)
					else
						dt_fim_suspensao = " X "
					end if
				
				if nr_parcela > maior_qtd_parcelas then
					maior_qtd_parcelas = nr_parcela
				end if
				
				dt_inicio = rs("dt_i")
				dt_fim = rs("dt_f")
				
				coluna_divisora = rs("coluna_divisora")
				cd_status = rs("cd_status")
					
					
					pos_atual = right(dt_ano,2)&zero(dt_mes)
					pos_aprov = right(year(dt_aprov_orc),2)&zero(month(dt_aprov_orc))
					pos_dif = pos_atual - pos_aprov
					qtd_parc_real = (nr_parcela - pos_dif)
					
					IF coluna_divisora = sem_aprovacao Then
						difereca_dt = DATEDIFF("m","1/"&month(dt_orcamento)&"/"&year(dt_orcamento),"1/"&zero(dt_mes)&"/"&dt_ano)
						cab_relatorio = "Orçamentos Aguardando Aprovação - "
						cor_relat = "#ec9997"
						
					ELSEIF coluna_divisora = pgto_suspensos Then
						difereca_dt = DATEDIFF("m","1/"&month(dt_aprov_orc)&"/"&year(dt_aprov_orc),"1/"&zero(dt_mes)&"/"&dt_ano)
						cab_relatorio = "Pagamentos Suspensos e/ou não cobrados -"
						cor_relat = "#e9e9e9"
					ELSEIF coluna_divisora = aprovados_recentes Then
						difereca_dt = DATEDIFF("m","1/"&month(dt_aprov_orc)&"/"&year(dt_aprov_orc),"1/"&zero(dt_mes)&"/"&dt_ano)
						cab_relatorio = "Orçamentos Aprovados recentes -"
						cor_relat = "#f1f5ba"
					ELSEIF coluna_divisora = compromissos Then
						difereca_dt = DATEDIFF("m","1/"&month(dt_entrega)&"/"&year(dt_entrega),"1/"&zero(dt_mes)&"/"&dt_ano)
						cab_relatorio = "Compromissos - Pagamentos no &quot;Contas a Pagar&quot; - "
						'cor_relat = "#d1e9d8"
						if cd_status = 0 Then cor_relat = "#d1f1c7" else cor_relat = "#d1e9d8"  end if
					ELSE
						diferenca_dt = 0
						cab_relatorio = "Listagem de orçamentos novos "
						cor_relat = "#bafcc9"
					END IF
				
			'*****************************************************
			'*** Cabeçalho e Rodapé ******************************
			'*****************************************************
			if num_linha = 1 then%>
				
				<tr>
					<td align="center" colspan="12" style="background-color:gray; color:white; font-size:14px;"><b><%=cab_relatorio%><%'&mes_selecionado(dt_mes)&"/"&dt_ano%><%'=dt_ano&"-"&dt_mes&"-"&dt_dia%></b></td>
				</tr>
				<tr style="background-color:silver;">
					<td><b>N°</b><br><img src='../imagens/px.gif' width='10' height='1'></td>
					<td><b>Nº Orç.</b><br><img src='../imagens/px.gif' width='60' height='1'></td>
					<td><b>Fornecedor</b><br><img src='../imagens/px.gif' width='85' height='1'></td>
					<%if coluna_divisora <> 4 Then%>
						<td><b>Dt aprov</b><br><img src='../imagens/px.gif' width='50' height='1'></td>
					<%else%>
						<td><b>Dt entrega</b><br><img src='../imagens/px.gif' width='50' height='1'></td>
					<%end if%>
					<!--td><b>Data I/F</b><br><img src='../imagens/px.gif' width='70' height='1'></td-->
					<td><b>Valor Total</b><br><img src='../imagens/px.gif' width='60' height='1'></td>
					<%'if coluna_divisora <> pgto_suspensos Then%>
						<td><b>Qtd<br>parc</b><br><img src='../imagens/px.gif' width='25' height='1'></td>
						<%'dt_mes_parcela = dt_mes -1
						'dt_ano_parcela = dt_ano
						'for i = 1 to qtd_meses_mostrar
							'dt_mes_parcela = dt_mes_parcela + 1
								'if dt_mes_parcela > 12 then
									'dt_mes_parcela = 1
									'dt_ano_parcela = dt_ano + 1
								'end if%>
									<td align="center"><b><%=left(mes_selecionado(dt_mes),3)&"/"&right(dt_ano,2)%></b><br>
									<img src='../imagens/px.gif' width='50' height='1'></td>
						<%'next%>
					
					
					<!--td>O.S. - DataEntrega<br><img src='../imagens/px.gif' width='150' height='1'></td-->
					
				</tr>
			<%array_subtotal=""
				if num_linha > qtd_linhas_pagina then 
					'num = 1
					num_linha = 1
				else	
					num = 1
					num_linha = 1
				end if
				
			end if%>
	<%'if coluna_divisora < 4 AND coluna_divisora > 1 Then%>
			<tr style="background-color:<%=cor_relat%>;" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>" onClick="window.open('../financeiro/diario_cad3.asp?mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>&cd_origem=6.<%=cd_orcamento%>&acao=inserir','Cadastro','width=500,height=480,scrollbars=1,status=1')" return false;">
				<td align="center"><%=num%></td>
				<td><%=nr_orcamento%></td>
				<td><%if len(nm_fornecedor) > 15 then response.write(left(nm_fornecedor,15)&" ...") Else response.write(nm_fornecedor) end if%> 
					<%if coluna_divisora = compromissos then 
						'response.write(" - "&nr_parcela&" - "&difereca_dt&" = "&qtd_parc_rest&" >>>("&nr_parcela-difereca_dt&")")
					else
						'response.write(" - "&nr_parcela)
					end if%> <%'=cd_status%>
				</td>
				<%if coluna_divisora <> 4 Then%>
					<td align="center"><%if not ISNULL(dt_aprov_orc) then response.write(zero(day(dt_aprov_orc))&"/"&zero(month(dt_aprov_orc))&"/"&right(year(dt_aprov_orc),2))%></td>
				<%else%>
					<td align="center"><%=dt_entrega%></td>
				<%end if%>
				<!--td align="center"><%=dt_inicio&" - "&dt_fim%></td-->
				
				<td align="right"><%if tipo = "financ" then%><a href="javascript:void(0);" return false;" onclick="window.open('../financeiro/diario_cad3.asp?cd_origem=6.<%=cd_orcamento%>&cd_forn=<%=cd_fornecedor%>&cd_equip=<%=cd_equipamento%>&cd_valor_orc=<%=replace(nm_valor,".","")%>&cd_tipo_orc=<%=cd_gestao%>&cd_unidade=<%=cd_unidade%>&campo=cd_codigo&visual=1&jan=1', 'pagamento_<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%if nm_valor <> "" Then response.write(FormatNumber(nm_valor,2))%></a><%else%><%if nm_valor <> "" Then response.write(formatnumber(nm_valor,2))%><%end if%></td>
				<td align="center"><%=nr_parcela%></td>
				
					<%'******************************************
					'*** Ditribuição das Parcelas - Divisão 4 ***
					'********************************************
					nr_parcela_max = 5
					coluna_princ = dt_mes
					
					if coluna_divisora	= compromissos then
							qtd_parc_rest = nr_parcela-difereca_dt
							coluna_registro = month(dt_inicio)
						for i = 0 to qtd_parc_rest
							if int(coluna_registro) > int(coluna_princ) then
								coluna_princ = coluna_princ + 1
									if coluna_princ > 12 then coluna_princ = 1%>
								<td>&nbsp;<%'=i&"-"&coluna_registro&"-"&coluna_princ%></td>
							<%else
								array_subtotal = array_subtotal&";"&i&":"&nr_valor_parcela%>
								<td align="right"><%if i = 0 then response.write("<b>")%><%'=i&"-"&coluna_registro&"-"&coluna_princ&"<br>"%> <%=formatnumber(nr_valor_parcela,2)%><%if i = 0 then response.write("</b>")%></td>
							<%end if
						next
						coluna_preeenchimento = (nr_parcela_max - qtd_parc_rest)-1
					
					'********************************************
					'*** Ditribuição das Parcelas - Divisão 2 ***
					'********************************************
					elseif coluna_divisora = pgto_suspensos then
							if nr_parcela > nr_parcela_max+1 Then nr_parcela = nr_parcela_max+1
						
						'for i = 0 to nr_parcela-1
						'	array_subtotal = array_subtotal&";"&i&":"&nr_valor_parcela
							%>
							<td align="right"><%if i = 0 then response.write("<b>")%><%'=i%> <%=formatnumber(nr_valor_parcela,2)%><%if i = 0 then response.write("</b>")%></td>
							<td align="center"><%=dt_inicio_suspensao%></td>
							<td align="center"><%=dt_fim_suspensao%></td>
							<td align="center" colspan="3">&nbsp;</td>
						<%'next
						coluna_preeenchimento = -1'(nr_parcela_max - nr_parcela)
					
					'************************************************
					'*** Ditribuição das Parcelas - Divisão < = 3 ***
					'************************************************
					elseif coluna_divisora	<= aprovados_recentes then
							if nr_parcela > nr_parcela_max+1 Then nr_parcela = nr_parcela_max+1
						
						for i = 0 to nr_parcela-1
							array_subtotal = array_subtotal&";"&i&":"&nr_valor_parcela
							%>
							<td align="right"><%if i = 0 then response.write("<b>")%><%'=i%> <%=formatnumber(nr_valor_parcela,2)%><%if i = 0 then response.write("</b>")%></td>
						<%next
						coluna_preeenchimento = (nr_parcela_max - nr_parcela)
						
					end if
						
						'*************************************
						'*** Preenche as colunas em branco ***
						'*************************************
						for i = 0 to coluna_preeenchimento%>
							<td align="right"><b><%'=i%></b> <%'=coluna_preeenchimento&"<br>"%>&nbsp;</td>
						<%next%>
				
				
				
				<!--td>	
						<%if coluna_divisora < compromissos then
						
						strsql = "SELECT * FROM VIEW_os_lista2 WHERE num_os_assist = '"&nr_orcamento&"' ORDER BY cd_unidade_os,dt_resposta_orc"
						Set rs_2 = dbconn.execute(strsql)
						num_os_2 = ""
						conta_os = 1
						ultima_entrega_os = ""
						marca_inconsistencia = 0
			
						while not rs_2.EOF
							cd_unidade_2 = rs_2("cd_unidade_os")
							num_os_2 = rs_2("num_os")
							dt_retorno_2 = rs_2("dt_retorno")
							dt_resposta_orc = rs_2("dt_resposta_orc")
							
							'****** Verifica Inconsistencias *******
							'if ultima_resposta_orc <> dt_resposta_orc  AND conta_os > 1 THEN
							'	marca_inconsistencia = 1
							
							if ultima_entrega_os <> dt_retorno_2  AND conta_os > 1 THEN
							'	marca_inconsistencia = 1
								
							ELSEIF  ISNULL(dt_retorno_2) THEN
								marca_inconsistencia = 1
								
							ELSEIF  ISNULL(dt_resposta_orc) THEN
								marca_inconsistencia = 1
								
							ELSEIF num_os_2 = "" Then
								marca_inconsistencia = 1
							else
								'marca_inconsistencia = 0
							end if
							
							'response.write("("&conta_os&")")
							%>
							
							<%'="<b>"&zero(cd_unidade_2)&"."&num_os_2&"</b> - "&dt_retorno_2%><%'=":"&ultima_entrega_os%><br>
						<%
						ultima_entrega_os = dt_retorno_2
						ultima_resposta_orc = dt_resposta_orc
						conta_os = conta_os +1
						
						rs_2.movenext
						wend
						
							IF num_os_2 = "" Then
								marca_inconsistencia = 1
						'		response.write("<b>-->> </b>")
							end if
							
							'if marca_inconsistencia = 1 AND not ISNULL(dt_retorno_2) AND num_os_2 <> "" then
							if marca_inconsistencia = 1 then
								'response.write("<br><b>xxxXXXxxx</b>")
								
							'elseif marca_inconsistencia = 0 AND not ISNULL(dt_retorno_2) AND num_os_2 <> "" then
							elseif marca_inconsistencia = 0 then
								'********* Efetua a alteração da data de entrega no Orçamento****
								'response.write("@cd_orcamento='"&cd_orcamento&"'<br> @dt_entrega="&year(dt_retorno_2)&"-"&month(dt_retorno_2)&"-"&day(dt_retorno_2)&"")
								'xsql = "up_orcamento_baixa @cd_orcamento='"&cd_orcamento&"', @dt_entrega='"&year(dt_retorno_2)&"-"&month(dt_retorno_2)&"-"&day(dt_retorno_2)&"'"
								'response.write("dt_aprov_orc ='"&year(dt_resposta_orc)&"-"&month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"'  cd_codigo="&cd_orcamento)
								'xsql = "UPDATE TBL_manutencao_orcamento SET dt_aprov_orc ='"&year(dt_resposta_orc)&"-"&month(dt_resposta_orc)&"-"&day(dt_resposta_orc)&"' WHERE  cd_codigo="&cd_orcamento&""
								
								'dbconn.execute(xsql)
						'		response.write("^^^^^^^^^^")
							else
						'		response.write("...")
							end if
						marca_inconsistencia = 0
						ultima_entrega_os = ""
						end if
						
						conta_os = 1%>
				</td -->
				
				
				
						
			</tr>
<%'end if%>			
			<%conta_os = 0
			dt_mes_parcela = 0
			coluna_preeenchimento = 0
			
			
			'pos_aprov = 0
			pos_dif = 0
			qtd_parc_real = 1
			
			grava_info = compara1
			num = num + 1
			num_linha = num_linha + 1
			ultima_divisao = coluna_divisora 
			xptop = ""
			rs.movenext
			wend
			
			'response.write(listaParcelas_total%>
			
			
					
					<tr><td colspan="7"><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td></tr>
					<tr>
						<td colspan="3">&nbsp;</td>
						<td align="right">&nbsp; Subtotal:</td>
						<td align="right"><b><%if total_valores > 0 then response.write(formatnumber(total_valores,2))%></b></td>
						<td>&nbsp;</td>
					<%for i = 0 to 5
						array_parcela = split(array_subtotal,";")
							for each parc_item in array_parcela
								'array_subtotal&";"&i&":"&nr_valor_parcela
								if instr(1,parc_item,i&":",1) then
									if parc_item <> "" then
										parc_item = replace(parc_item,i&":","")
											subtotal = subtotal+abs(parc_item)
									end if
								end if
							next%>
						<td align="right"><%if i = 1 then response.write("<b>")%><%if subtotal > 0 Then response.write(formatnumber(subtotal,2))%><%if i = 1 then response.write("</b>")%></td>
					<%subtotal = 0
					next%>
				</tr>
		
				<!--tr><td colspan="12">&nbsp;</td></tr>
			<tr style="background-color:silver;">
				<td colspan="3">&nbsp;</td>
				<td align="center"><i>TOTAL</i></td>
				<td align="right"><i><%if total_nm_valor <> "" Then response.write(formatNumber(total_nm_valor,2))%></i></td>
				<td colspan="7">&nbsp;</td>
				
			</tr-->
					</table>
				</td>
			</tr>
			<%=response.write(listaParcelas_total)%>			
</table>
<br>
<br>
</body>

<%if tipo <> "financ" then%>			
	<SCRIPT LANGUAGE="javascript">
		formataMoeda(document.forms.form.nm_valor, 2);	
		
		function JsDelete(cod)
		   {
			if (confirm ("Tem certeza que deseja excluir o orçamento?"))
		  {
		  document.location.href('manutencao_2/acoes/manutencao_acao.asp?cd_orcamento='+cod+'&acao=3&filtro=orc_cad_del');
			}
			  }
	</SCRIPT>
<%end if%>
