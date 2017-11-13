<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<!-- #include file="../css/geral.htm"-->
<body>
<%tipo = request("tipo")
cd_orcamento = request("cd_orcamento")%>
<%if tipo = "cad_orc_aprov" then%>
	<table align="center">
	
		<tr>
			<td align="center" colspan="4" style="background-color:gray; color:white; font-">Cadastrar Orçamentos Aprovados</td>
		</tr>
		<%if cd_orcamento <> "" Then
			strsql ="SELECT * FROM View_manutencao_orcamento WHERE cd_codigo = '"&cd_orcamento&"' ORDER BY dt_orcamento"
			Set rs_orc = dbconn.execute(strsql)
				while not rs_orc.EOF
					nm_fornecedor = rs_orc("nm_fornecedor")
					nr_orcamento = rs_orc("nr_orcamento")
					dt_orcamento = rs_orc("dt_orcamento")
						if dt_orcamento <> "" Then	
							dia_orc = zero(day(dt_orcamento))
							mes_orc = zero(month(dt_orcamento))
							ano_orc = year(dt_orcamento)
						end if
						
					nm_valor = rs_orc("nm_valor")
						'total_nm_valor = total_nm_valor + nm_valor
					nr_parcela = rs_orc("nr_parcela")
						if nr_parcela > 1 then
							nr_valor_parcela = nm_valor / nr_parcela
						else
							nr_valor_parcela = 0
						end if
				rs_orc.movenext
				wend
			filtro = "orc_cad_upd"
			nm_botao = "Altera"
		else
			filtro = "orc_cad"
			nm_botao = "Cadastra"
		end if%>
		<tr>
			<td>&nbsp;
			<form name="form" action="manutencao_2/acoes/manutencao_acao.asp" method="get" id="form">
			<input type="hidden" name="filtro" value="<%=filtro%>">
			<input type="hidden" name="cd_orcamento" value="<%=cd_orcamento%>">
			<input type="hidden" name="acao" value="1">
			<input type="hidden" name="jan" value="0">
			</td>
		</tr>
		<tr>
			<td><b>Nº Orç.</b><br>
				<img src="../imagens/px.gif" width="120" height="1"><br>
				<input type="text" name="nr_orcamento" size="10" maxlength="20" value="<%=nr_orcamento%>"></td>
			<%strsql ="TBL_Fornecedor order by nm_fornecedor"
			Set rs_forn = dbconn.execute(strsql)%>
			<td colspan="2"><b>Assist./Fornecedor.</b><br>
			<img src="../imagens/px.gif" width="120" height="1"><br>
				<span id='divForn'>
					<select name="cd_fornecedor" class="inputs" onFocus="nextfield ='observacoes';">
						<option value="">
						<%Do while Not rs_forn.EOF%><%fornecedor=rs_forn("nm_fornecedor")%><%if nm_fornecedor=fornecedor Then%><%ck_forn="selected"%><%else ck_forn=""%><%end if%>
						<option value="<%=rs_forn("cd_codigo")%>" <%=ck_forn%>><%=rs_forn("nm_fornecedor")%></option><%rs_forn.movenext
						Loop%>
					</select>
				</span>
			</td>
							
		</tr>				
		<tr>
			<td><b>Data</b><br>
				<img src="../imagens/px.gif" width="170" height="1"><br>
				<input type="text" name="dt_dia" size="2" maxlength="2" value="<%=dia_orc%>">/
				<input type="text" name="dt_mes" size="2" maxlength="2" value="<%=mes_orc%>">/
				<input type="text" name="dt_ano" size="4" maxlength="4" value="<%=ano_orc%>"></td>
			<td><b>valor</b><br>
				<img src="../imagens/px.gif" width="120" height="1"><br>
				<input type="text" name="nm_valor" size="10" maxlength="20" value="<%=nm_valor%>" onKeyup="valores();" style="text-align:right"></td>
			<td><b>Qtd. Parcelas.</b><br>
				<img src="../imagens/px.gif" width="80" height="1"><br>
				<input type="text" name="nr_parcela" size="2" maxlength="3" value="<%=nr_parcela%>"></td>
		</tr>
		<tr>
			<td><input type="submit" value="<%=nm_botao%>"></td>
			<td><input type="reset" name="Reset" value="Reset" onclick="javascript:window.location.replace('manutencao2.asp?tipo=cad_orc_aprov')"></td>
		</tr>
			</form>
	
	</table>
<%end if%>
<br class="no_print">
<br class="no_print">
<table align="center" border="1" style="">
	<tr>
		<td align="center" colspan="9" style="background-color:gray; color:white; font-size:14px;">Lista de orçamentos cadastrados</td>
	</tr>
	<%num = 1
	strsql ="SELECT * FROM View_manutencao_orcamento Where cd_pagto is null ORDER BY dt_orcamento"
		Set rs = dbconn.execute(strsql)
			while not rs.EOF
				cd_orcamento = rs("cd_codigo")
				cd_fornecedor = rs("cd_fornecedor")
				nm_fornecedor = rs("nm_fornecedor")
				nr_orcamento = rs("nr_orcamento")
				dt_orcamento = rs("dt_orcamento")
				nm_valor = rs("nm_valor")
					'total_nm_valor = total_nm_valor + nm_valor
					total_nm_valor = abs(total_nm_valor) + abs(nm_valor)
				nr_parcela = rs("nr_parcela")
					if nr_parcela > 1 then
						nr_valor_parcela = nm_valor / nr_parcela
					else
						nr_valor_parcela = 0
					end if
			
			if num = 1 then%>
				<tr style="background-color:silver">
					<td><b>Num.</b><br><img src='../imagens/px.gif' width='20' height='1'></td>
					<td><b>Fornecedor</b><br><img src='../imagens/px.gif' width='200' height='1'></td>
					<td><b>Nº Orçamento</b><br><img src='../imagens/ox.gif' width='90' height='1'></td>
					<td><b>Data</b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<td><b>Valor Total</b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<td><b>Qtd parc.</b><br><img src='../imagens/px.gif' width='50' height='1'></td>
					<td><b>Valor parcela</b><br><img src='../imagens/px.gif' width='70' height='1'></td>
					<td colspan="2">&nbsp;</td>
				</tr>
			<%end if%>
			
			<tr>
				<td align="center"><%=num%></td>
				<td><%=nm_fornecedor%></td>
				<td><a href="javascript:void(0);" onclick="window.open('<%if tipo = "financ" then response.write("../")%>manutencao_2/orcamentos_aprovados_os.asp?cd_orcamento=<%=cd_orcamento%>', 'orcamento<%=cd_orcamento%>', 'width=520, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=nr_orcamento%></a></td>
				<td align="center"><%=zero(day(dt_orcamento))&"/"&zero(month(dt_orcamento))&"/"&right(year(dt_orcamento),2)%></td>
				<td align="right"><%if tipo = "financ" then%><a href="javascript:void(0);" return false;" onclick="window.open('../financeiro/diario_cad3.asp?cd_origem=6.<%=cd_orcamento%>&cd_forn=<%=cd_fornecedor%>&cd_equip=<%=cd_equipamento%>&cd_valor_orc=<%=replace(nm_valor,".","")%>&cd_tipo_orc=<%=cd_gestao%>&cd_unidade=<%=cd_unidade%>&campo=cd_codigo&visual=1&jan=1', 'pagamento_<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%if nm_valor <> "" Then response.write(FormatNumber(nm_valor,2))%></a><%else%><%if nm_valor <> "" Then response.write(formatnumber(nm_valor,2))%><%end if%></td>
				<td align="center"><%=nr_parcela%></td>
				<td align="right"><%if nr_valor_parcela > 0 then response.write(formatnumber(nr_valor_parcela,2)) Else response.write("&nbsp;") end if%></td>
				<%if tipo <> "financ" then%>
					<td><img src="../imagens/ic_editar.gif" alt="" width="13" height="14" border="0" onclick="javascript:window.location.replace('manutencao2.asp?tipo=cad_orc_aprov&cd_orcamento=<%=cd_orcamento%>')"></td>
					<td><img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="javascript:JsDelete(<%=cd_orcamento%>)"></td>
				<%else%>
					<td colspan="2">&nbsp;</td>
				<%end if%>
			</tr>
			<%num = num + 1
			rs.movenext
			wend%>
			<tr style="background-color:silver;">
				<td colspan="3">&nbsp;</td>
				<td align="center">Total</td>
				<td align="right"><b><%if total_nm_valor <> "" Then response.write(formatNumber(total_nm_valor,2))%></b></td>
				<td colspan="4">&nbsp;</td>
			</tr>
</table>
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
