<!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->

<style type="text/css">
<!--
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; }  
.textopadrao {font-family: arial;font-size: 14px;}
.textopequeno {font-family: verdana;font-size: 8px;}
.textocinza {font-family: verdana;font-size:8px;color:#363636;}
-->
</style>

<%
mes = request("mes")
ano = request("ano")
unit = "x" 'Para futura possibilidade de escolha por unidade

relatorio = request("relatorio")
	
		if relatorio = "compl" Then
		txt_soma_parcial = ""
		txt_tipo_relatorio = "TOTAL FINAL"
		colunas_mostra = "a"
		largura_col = "760"
		border = "0"
		border2 = "0"
		
		Elseif relatorio = "prod_ben" Then
		txt_soma_parcial = ""
		txt_tipo_relatorio = "DETALHADO DE PRODUTIVIDADE"
		colunas_mostra = "b"
		largura_col = "640"
		border = "0"
		border2 = "0"
		
		Elseif relatorio = "beneficios" then
		txt_soma_parcial = ""
		txt_tipo_relatorio = "DETALHADO DE BENEFÍCIOS"
		colunas_mostra = "c"
		largura_col = "640"
		border = "0"
		border2 = "0"
		
		Elseif relatorio = "consolidado" Then
		txt_soma_parcial = ""
		txt_tipo_relatorio = "CONSOLIDADO"
		colunas_mostra = "d"
		largura_col = "400"
		border = "0"
		border2 = "0"
		
		Else
		txt_soma_parcial = "Soma Parcial"
		colunas_mostra = "a"
		border = "0"
		border2 = "0"
		end If%>
<html>
<body>


<table width="" cellspacing="0" cellpadding="0" border="<%=border%>" class="textoPadrao" align="center">
	<tr>
		<td align="center">
			<%if relatorio = "" Then%>
			<table width="500" cellspacing="1" cellpadding="1" border="0" class="textoPadrao" bgcolor="#fbfbfb">
				<tr>
					<td colspan="100%" textoPadrao><b>RELATÓRIOS - <font color="red">RELATÓRIO FINAL</font></b></td>
				</tr>
				<tr bgcolor=#cococo>
					<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
				</tr>
			<%end if%>
<%If ano = "" AND mes = "" Then
'*********************************************************
'*				      1ª Parte							'*
'*********************************************************
%>
	<form action="relatorio_coop.asp">	
		<tr>
			<td>&nbsp;</td>
			<td colspan="5">&nbsp;<!--[<a href="produtividade.asp">Mudar Ano</a>]  &nbsp;&nbsp;/  &nbsp;&nbsp;[<a href="produtividade.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>">Mudar Cooperado]--></td>
		</tr>
		<tr bgcolor=#cococo>
			<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td><input type="hidden" name="relatorio" value="<%=relatorio%>">
			<td><select name="ano">
					<option value="">Selecione </option>
					<%'Mostra meses disponíveis no Banco de dados
			campo = "ano"
			nm_tbl = "View_relatorio_cooperativa"
			ordem = "ano"
		
			xsql_meses = "up_meses_disponiveis @campo='"&campo&"', @nm_tbl='"&nm_tbl&"', @condicao='"&condicao&"', @ordem='"&ordem&"'"
			Set rs_meses = dbconn.execute(xsql_meses)
			Do While not rs_meses.EOF
			'mes = rs_meses("mes")
			ano = rs_meses("ano")%>
					<option value="<%=ano%>"><%=ano%></option>
					<%
					'umavez = "1"
					rs_meses.movenext
					loop%>
				</select><input type="submit" value="ok"></td>
			<td>&nbsp; Informe o ano </td>
	</form>		
		</tr>
		<tr><td>&nbsp;</td>
			<td>&nbsp;<br><br></td>
		</tr>
<%ElseIf ano <> "" AND mes = "" Then
'*********************************************************
'*				      2ª Parte							'*
'*********************************************************

%>		
		
	<form action="relatorio_coop.asp">
		<tr>
			<td>&nbsp;</td>
			<td colspan="5">[<a href="relatorio_coop.asp?relatorio=<%=relatorio%>">Mudar Ano</a>]  &nbsp;&nbsp;</td>
		</tr>
		<tr bgcolor=#cococo>
			<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td align="">ANO: <b><%=ano%></b></td>
		</tr>
		<tr>
			<td>&nbsp;</td><input type="hidden" name="relatorio" value="<%=relatorio%>">
			<td><select name="mes">
					<option value="">Selecione</option>
					<%'Mostra meses disponíveis no Banco de dados
			campo = "ano, mes"
			nm_tbl = "View_relatorio_cooperativa"
			condicao = " WHERE ano="&ano&" "
			ordem = "ano, mes"
		
			xsql_meses = "up_meses_disponiveis @campo='"&campo&"', @nm_tbl='"&nm_tbl&"', @condicao='"&condicao&"', @ordem='"&ordem&"'"
			Set rs_meses = dbconn.execute(xsql_meses)
			Do While not rs_meses.EOF
			mes = rs_meses("mes")
			ano = rs_meses("ano")%>
					<option value="<%=mes%>"><%=mes_selecionado(mes)%></option>
					<%
					'umavez = "1"
					rs_meses.movenext
					loop%>
				</select>
				<input type="hidden" name="ano" value="<%=ano%>">
				<input type="submit" value="ok"></td>
			<td>&nbsp; Informe o mês </td>
			
		</tr>
	</form>
		<tr><td>&nbsp;</td>
			<td>&nbsp;<br><br></td>
		</tr>
			
<%Elseif mes <> "" AND Ano <> "" Then
'*********************************************************
'*				      3ª Parte							'*
'*********************************************************

	'**** inclui 0 ao mes menor que 10 *****
	If mes < 10 Then
	mes1 = "0"&mes
	Else
	mes1 = mes
	End if
	dt_messel =  ano & mes1
	'dt_messel =  ano & mes
	
	'unidade = request("unidade")
	
	Function prod_horas(min) 'funcionarios
	Hora = Int(min/60)   '60 minutos
	Minutos = min Mod 60 '60 segundos
	NovaHora = Right(Hora,4) & ":" & Right("0" & Minutos,2)
	prod_horas = novahora
	End Function
	
	Function prod_valor(valor)
	Hora = Int(valor/60)   '60 minutos
	Minutos = valor Mod 60 '60 segundos
	hora_valor = Right(Hora,6) & "," & Right("0" & Minutos,2)
	hora_valor = hora_valor
	money = round(hora_valor,2)
	'money = hora_valor
	prod_valor = money 
	End Function
			
		dt_ano = YEAR(Now)
		dt_mes = Month(Now)
		dt_dia = Day(now)
		'str_condicoes = "AND cd_codigo= "
		
		'*** Conta os registros ***
		xsql1 = "up_relatorio_cooperativa_count @mes="&mes&", @ano="&ano&", @uni="&unit&""
		Set rs_count = dbconn.execute(xsql1)
		
		'*** Lista os registros ***
		xsql = "up_relatorio_cooperativa @mes="&mes&", @ano="&ano&", @str_condicoes='"&str_condicoes&"'"
		Set rs = dbconn.execute(xsql)
		
		strmes= rs("mes")
		strano= rs("ano")
			if strmes < 10 Then
			strmes = 0 & strmes
			End if
				cd_alteracao = strano & strmes
		
		'*** lista os valores dos benefícios ***
		'*** Valores Cooperados ***
		xsql_valores ="up_valores_lista @cd_alteracao='"&cd_alteracao&"'"
		Set rs_valores = dbconn.execute(xsql_valores)
		%>
		<%if relatorio = "" Then%>
		<tr>
			<td>&nbsp;</td>
			<td colspan="5">[<a href="relatorio_coop.asp">Mudar Ano</a>]  &nbsp;&nbsp;/  &nbsp;&nbsp;[<a href="relatorio_coop.asp?ano=<%=ano%>">Mudar Mês]</td>
		</tr>
		<tr bgcolor=#cococo>
			<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
		</tr>
	</table>
		<%end if%>
		<%If rs.EOF Then ' *** Sem registros sobre o mes solicitado ***%>
		<Table width="780" border="0" cellspacing="1" cellpadding="0" class="textopequeno">
		<tr>
			<td class="textopequeno" rowspan="2000" bgcolor="000000"></td>
			<td colspan=16 class="textopequeno"><b>Relatórios &raquo; <font color="red">Cooperativa</font></b><br><br><%=mes_selecionado(mes)%>/<%=ano%><br></td>
		</tr>
		<tr class="textopequeno"><td>Ainda não existem registros. </td></tr>
		<%Else '*** Mostra a tablea com os valores do mes solicitado ***%>
		<Table width="<%=largura_col%>" border="<%=border2%>" cellspacing="1" cellpadding="0" class="textopequeno">
		<tr>
		<!--tr>	<td rowspan="2000" bgcolor="000000"></td-->
			<%if relatorio <> "" Then%>
			<td class="textopequeno" colspan=21 align="center"><b><font color="red">VD LAP CIRURGICA LTDA</font> <br>
								RELATÓRIO <%=txt_tipo_relatorio%><br><br>
								&nbsp;VENCIMENTO 15/<%if mes + 1 = 13 Then%>
								<%mes_pag = 1%>
								<%ano_pag = ano + 1%>
								<%else%>
								<%mes_pag = mes + 1%><%ano_pag = ano%>
								<%end if%>
								<%=mes_pag%>/<%=ano_pag%> <br>
								Referente ao mês de <%=mes_selecionado(mes)%> - <%=ano%><br></b><br>
			</td>
			<%else%>
			<td colspan=16><b>Relatórios &raquo; <font color="red">Cooperativa</font></b><br><br><%=mes_selecionado(mes)%>/<%=ano%>  <%'=cd_alteracao%><br></td>
			<%end if%>
		</tr>
			<td colspan="21"><img src="px_black.jpg" width="100%" height="1px"></td>
		<tr><!--Linha 2-->
			<td bgcolor="#B5B5B5" width="3"><b>&nbsp;</b></td>
			<td bgcolor="#B5B5B5" width="5"><b>&nbsp;</b></td>
			<td bgcolor="#B5B5B5" width="20"><b>&nbsp;</b></td>
			<!--td bgcolor="#B5B5B5" width="20"><b>&nbsp;</b></td-->
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td bgcolor="#B5B5B5" align="center" valign="top" colspan=3><b>&nbsp;PRODUTIVIDADE</b></td>
			<%elseif colunas_mostra = "d" Then%>
			<td bgcolor="#B5B5B5" align="center" valign="top" colspan=3><b>&nbsp;<!--sem legenda--></b></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "c" Then%>
			<td bgcolor="#E8E8E8" align="center" valign="top" rowspan="100%"><b>&nbsp;<img src="../../imagens/px.gif" width="8"><font color="#FFFFFF""></font></b></td>
			<td bgcolor="#B5B5B5" align="center" valign="top" colspan=2><b>&nbsp;TRANSPORTE&nbsp;</b></td>
			<td bgcolor="#E8E8E8" align="center" valign="top" rowspan="100%"><b>&nbsp;<img src="../../imagens/px.gif" width="4"><font color="#FFFFFF""></font></b></td>
			<td bgcolor="#B5B5B5" align="center" valign="top" colspan=2><b>&nbsp;REFEI&Ccedil;&Atilde;O&nbsp;</b></td>
			<td bgcolor="#E8E8E8" align="center" valign="top" rowspan="100%"><b>&nbsp;<img src="../../imagens/px.gif" width="4"><font color="#FFFFFF""></font></b></td>
			<td bgcolor="#B5B5B5" align="center" valign="top" colspan=2><b>&nbsp;CONV&Ecirc;NIO&nbsp;</b></td>
			<td bgcolor="#E8E8E8" align="center" valign="top" rowspan="100%"><b>&nbsp;<img src="../../imagens/px.gif" width="4"><font color="#FFFFFF""></font></b></td>
			<td bgcolor="#B5B5B5" align="right" valign="top"><b>&nbsp;INSS&nbsp;</b></td>
			<%end if%>
			<%if colunas_mostra <> "d" Then%>
			<td bgcolor="#E8E8E8" align="center" valign="top" rowspan="100%"><b>&nbsp;<img src="../../imagens/px.gif" width="8"><font color="#FFFFFF""></font></b></td>
			<td bgcolor="#B5B5B5" align="right" valign="top"><b>&nbsp;BENEFÍCIOS&nbsp;</b></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td bgcolor="#E8E8E8" align="center" valign="top" rowspan="100%"><b>&nbsp;<img src="../../imagens/px.gif" width="8"><font color="#FFFFFF""></font></b></td>
			<td bgcolor="#B5B5B5" align="right" valign="top"><b>&nbsp;TOTAL&nbsp;</b></td>
			<%end if%>
		</tr>
		<td colspan="21"><img src="px_black.jpg" width="100%" height="1px"></td>
		<tr bgcolor="#B5B5B5"><!--Linha 3-->
		<%
		'*** Valores dos benefícios ***
		If Not rs_valores.EOF Then
		nr_transporte = rs_valores("nr_transporte")
		nr_refeicao = rs_valores("nr_refeicao")
		nr_convenio = rs_valores("nr_convenio")
		nr_inss = rs_valores("nr_inss")
		nr_inss = rs_valores("nr_inss")
		End if
		' ********* CABEÇALHO **********
		%>
			<td><b>&nbsp;</b></td>
			
			<td align="center" valign="top"><b>&nbsp;Mat. </b></td>
			<td><b>&nbsp;Nome </b></td>
			<!--td><b>&nbsp; </b></td-->
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td align="right" valign="top"><b>&nbsp;Horas&nbsp;</b></td>
			<td><b>&nbsp;&nbsp;$/hr&nbsp;&nbsp;</b></td>
			<td align="right" valign="top"><b>&nbsp;Reais&nbsp;</b></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "c" Then%>
			<td align="center" valign="top"><b>&nbsp;dias</b></td>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(nr_transporte,2)%></b></td>
			<td align="center" valign="top"><b>&nbsp;dias</b></td>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(nr_refeicao,2)%></b></td>
			<td align="center" valign="top"><b>&nbsp;Qtd</b></td>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(nr_convenio,2)%></b></td>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(nr_inss,2)%></b></td>
			<%end if%>
			<td><b>&nbsp;</b></td>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td><b>&nbsp;</b></td>
			<%end if%>
				
		</tr> 
		<td colspan="21"><img src="px_black.jpg" width="100%" height="1px"></td>
		<%
		contagem = 0
		Do While Not rs.EOF
		
		'*** Insere o nome da Unidde ***
		nm_unidade = rs("nm_unidade")
		status = rs("nm_status")
		
		if status = 3 then
		status_unidade = "CONVÊNIO - "
		Elseif status = 5 Then
		status_unidade = "TREINAMENTO - "
		End if
		
		If contagem = 0 then
		%>
		<tr><!--Linha 4-->
			<td colspan="2"><b>&nbsp;</b></td>
			<td Colspan=5><b>&nbsp;<%=status_unidade%><%=nm_unidade%></b></td>
		</tr>
		<%
		End if
		
		'*** Dados do cooperado ***
		mat = rs("cd_matricula")
		cd_codigo = rs("cd_codigo")
		nome = rs("nm_nome")
		sobrenome = rs("nm_sobrenome")
		
		funcao = rs("cd_funcao")
		dt_demissao = int(rs("dt_desliga"))
		'status = rs("nm_status") *** Posicionado na inserção da unidade
		unidade = rs("cd_unidades")
		
		'*** Verifica se o funcionário deve receber INSS ***
		dt_mes_beneficio = dt_messel + 1
		
		if dt_demissao = dt_mes_beneficio OR dt_demissao < dt_mes_beneficio Then 
		valor_inss = "0"
		Else
		valor_inss = nr_inss
		t="inss"
		End if
		
		mes_desl = Month(rs("dt_desliga"))
		data_selecionada = mes
		
		'*** Produtividade ***
		horas_trabalhadas = prod_valor(rs("total_trabalhado"))
		
		'if funcao = 1 AND status = 1 Then
		'valor_hora = rs_valores("nr_valor")
		'Elseif funcao = 3 AND Status = 1 Then
		'valor_hora = rs_valores("nr_tecnico")
		'Elseif funcao = 1 AND status = 4 Then
		'valor_hora = rs_valores("nr_tecnico")
		'Elseif funcao = 4 AND status = 1 Then
		'valor_hora = rs_valores("nr_enfermeiro")
		'Elseif funcao = 4 AND status = 4 Then
		'valor_hora = rs_valores("nr_enfermeiro")
		
		'Elseif funcao = 1 AND status = 5 Then
		'valor_hora = rs_valores("nr_valor_trei")
		'valor_inss = 0
		'Elseif funcao = 3 AND status = 5 Then
		'valor_hora = rs_valores("nr_tecnico_trei")
		'valor_inss = 0
		
		'***** Valores organizados *****
		if funcao = 1 AND status = 1 Then
		valor_hora = rs_valores("nr_valor")
		Elseif funcao = 1 AND status = 3 Then
		valor_hora = rs_valores("nr_valor")
		valor_inss = 0
		Elseif funcao = 1 AND status = 4 Then
		valor_hora = rs_valores("nr_valor")
		Elseif funcao = 1 AND status = 5 Then
		valor_hora = rs_valores("nr_valor_trei")
		valor_inss = 0
		
		Elseif funcao = 3 AND Status = 1 Then
		valor_hora = rs_valores("nr_tecnico")
		Elseif funcao = 3 AND Status = 3 Then
		valor_hora = rs_valores("nr_tecnico")
		valor_inss = 0
		Elseif funcao = 3 AND Status = 4 Then
		valor_hora = rs_valores("nr_tecnico")
		Elseif funcao = 3 AND status = 5 Then
		valor_hora = rs_valores("nr_tecnico_trei")
		valor_inss = 0
		
		Elseif funcao = 4 AND status = 1 Then
		valor_hora = rs_valores("nr_enfermeiro")
		Elseif funcao = 4 AND status = 3 Then
		valor_hora = rs_valores("nr_enfermeiro")
		valor_inss = 0
		Elseif funcao = 4 AND status = 4 Then
		valor_hora = rs_valores("nr_enfermeiro")
		Elseif funcao = 4 AND status = 5 Then
		valor_hora = rs_valores("nr_enfermeiro")
		valor_inss = 0
		
		
		Else
		valor_hora = 0
		End if
		
		'*** Valor gestores ***
		valor_gestor = rs_valores("nr_valor_gestores")
		
		'*** Operaçoes da Produtividade ***
		reais = horas_trabalhadas * valor_hora
		
		'*** Contador de registros por unidade/mes
		conta = rs_count("conta")
		
		'*** verifica se o benefício já foi atribuído ao cooperado ***
		analise = Instr(1,codigos,cd_codigo,1)
		codigos = codigos & "," & cd_codigo
		
		%>
		<!--Linha 5-->
		<tr bgcolor="<%=strlinha%>" onmouseover="mOvr(this,'#eaeaea');" onmouseout="mOut(this,'FFFFFF');" <%if relatorio = "" Then%> onclick="javascript:window.open('produtividade/prod_cartao_ajuste.asp?cd_codigo=<%=cd_codigo%>&mes=<%=mes%>&ano=<%=ano%>','AJUSTE','width=350,height=350');"<%end if%>>
			<td align="right" valign="top"><b>&nbsp;<%=contagem + 1%></b></td>
			
			<td align="center" valign="top">&nbsp;<%'=cd_codigo&"-"%><%=mat%></td>
			<td>&nbsp;<%=nome%>&nbsp;<%=sobrenome%>&nbsp;<font color="<%=strcores%>"></font></td>
			<!--td><b>&nbsp;<%=funcao%>-<%=status%>-<%=unidade%></b><font color="<%=strcores%>"></font></td-->
			
			<!--Dados dos Cooperados-->
			<%if colunas_mostra = "a" OR colunas_mostra = "b" OR colunas_mostra = "d" Then%>
			<td align="right" valign="top">&nbsp;<%=tiraacento(FormatNumber(horas_trabalhadas,2))%></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td align="center" valign="top">&nbsp;<%=FormatNumber(valor_hora,2)%></td>
			<td align="right" valign="top">&nbsp;<b><%=FormatNumber(reais,2)%></b></td>
			<%end if%>
			<!--Benefícios-->
			<%'*** Mostra os benefícios do cooperado ***
			mes_beneficio = mes + 1
				if mes_beneficio = 13 then
				mes_beneficio = 1
				End if
			
			If mes = 12 then
			ano_beneficio = ano + 1
			Else
			ano_beneficio = ano
			End if
			
			xsql_beneficios ="up_relatorio_beneficios @cd_codigo='"&cd_codigo&"', @dt_mes='"&mes_beneficio&"', @dt_ano='"&ano_beneficio&"'"
			Set rs_beneficios = dbconn.execute(xsql_beneficios)
			
			'*** Mostra as quantidades de benefícios de cada cooperado ***
			if rs_beneficios.EOF Then
			Else
			nr_qtd_transporte = rs_beneficios("nr_qtd_trans")
			nr_qtd_refeicao = rs_beneficios("nr_qtd_ref")
			nr_qtd_convenio = rs_beneficios("nr_qtd_conv")
			nr_qtd_inss = rs_beneficios("nr_qtd_inss")
			End if
			
			'*** Verifica se o INSS está em caso especial ***
			If nr_qtd_inss = "1" Then
			valor_inss = nr_inss
			End if
			
			'*** Apaga os valores de benefícios caso já tenha sido atribuído
			If analise <> 0 Then
			nr_qtd_transporte = int("0")
			nr_qtd_refeicao = int("0")
			nr_qtd_convenio = int("0")
			nr_qtd_inss = int("0")
			valor_inss = "0"
			end If
			
			'*** Multiplica a quantidade x valor do benefício ***
			valor_transporte = nr_transporte * nr_qtd_transporte
			valor_refeicao = nr_refeicao * nr_qtd_refeicao
			valor_convenio = nr_convenio * nr_qtd_convenio
			valor_inss = valor_inss 'Não há multiplicação
			%>
			
			
			<%if colunas_mostra = "a" OR colunas_mostra = "c" Then%>
			<td align="right" valign="top">&nbsp;<%=nr_qtd_transporte%></td>
			<td align="right" valign="top">&nbsp;<%=FormatNumber(valor_transporte,2)%></td>
			<td align="right" valign="top">&nbsp;<%=nr_qtd_refeicao%></td>
			<td align="right" valign="top">&nbsp;<%=FormatNumber(valor_refeicao,2)%></td>
			<td align="right" valign="top">&nbsp;<%=nr_qtd_convenio%></td>
			<td align="right" valign="top">&nbsp;<%=FormatNumber(valor_convenio,2)%></td>
			<td align="right" valign="top">&nbsp;<%=FormatNumber(valor_inss,2)%></td>
			<%end if%>
		<%' *************** Início das somas dos valores pagos ***************
		
		'*** Produtividade ***
		'*** Soma parcial de cada unidade ***
		soma_h_trab_parc = soma_h_trab_parc + horas_trabalhadas
		soma_reais_parc = soma_reais_parc + reais
		
		'*** Soma as horas de todos colaboradores ***
		soma_h_trab = soma_h_trab + horas_trabalhadas
		soma_reais = soma_reais + reais
		'*** Fim ---------------------------------***
		
		'*** Benefícios***
		'*** Soma Parcial de benefícios por coluna ***
		soma_nr_qtd_transporte_parc = soma_nr_qtd_transporte_parc + nr_qtd_transporte
		soma_valor_transporte_parc = soma_valor_transporte_parc + valor_transporte
		soma_nr_qtd_refeicao_parc = soma_nr_qtd_refeicao_parc + nr_qtd_refeicao
		soma_valor_refeicao_parc = soma_valor_refeicao_parc + valor_refeicao
		soma_nr_qtd_convenio_parc = soma_nr_qtd_convenio_parc + nr_qtd_convenio
		soma_valor_convenio_parc = soma_valor_convenio_parc + valor_convenio
		soma_valor_inss_parc = soma_valor_inss_parc + valor_inss
		
		'*** Soma Parcial de benefícios por linha ***
		soma_nr_beneficios_parc = valor_transporte + valor_refeicao + valor_convenio + valor_inss
		
		'*** Soma Parcial do total de benefícios por coluna ***
		soma_nr_total_beneficios_parc = soma_nr_total_beneficios_parc + soma_nr_beneficios_parc
		
		'*** Soma Parcial do total de produtividade + benefícios
		soma_total_bruto_parc = reais + soma_nr_beneficios_parc
		
		'*** Soma os benefícios de todos os colaboradores
		soma_nr_qtd_transporte = soma_nr_qtd_transporte + nr_qtd_transporte
		soma_valor_transporte = soma_valor_transporte + valor_transporte
		soma_nr_qtd_refeicao = soma_nr_qtd_refeicao + nr_qtd_refeicao
		soma_valor_refeicao = soma_valor_refeicao + valor_refeicao
		soma_nr_qtd_convenio = soma_nr_qtd_convenio + nr_qtd_convenio
		soma_valor_convenio = soma_valor_convenio + valor_convenio
		soma_valor_inss = soma_valor_inss + valor_inss
		
		'*** Soma o total de benefícios ***
		soma_nr_total_beneficios = soma_nr_total_beneficios + soma_nr_beneficios_parc
		
		
		
		'*** Soma do total de produtividade + benefícios ***
		soma_total_bruto_parc = reais + soma_nr_beneficios_parc
		
		'*** Soma parcial do total bruto de cada unidade ***
		soma_total_parc = soma_total_parc + soma_total_bruto_parc
		
		'*** Soma total do bruto das unidades ***
		soma_total = soma_total + soma_total_bruto_parc
		
		
		' *************** Fim das somas dos valores pagos ***************
		
		'**** Apaga os valores das variáveis ***
		funcao= ""
		status= ""
		unidade= ""
		soma_unidade = 0
		
		contagem = contagem + 1
		If conta = contagem then
		contagem = 0
		rs_count.movenext
		soma_unidade = 1
		End if
		%>
			<!--Soma parcial de valores-->
			<%if colunas_mostra <> "d" Then%>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(soma_nr_beneficios_parc,2)%></b></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(soma_total_bruto_parc,2)%></b></td>
			<%end if%>
		<%'*** Limpa as quantidades de benefícios ao final da soma do Cooperado ***
			nr_qtd_transporte = int("0")
			nr_qtd_refeicao = int("0")
			nr_qtd_convenio = int("0")
			nr_qtd_inss = int("0")
			valor_inss = "0"
			'End if
			%>
		</tr>
		<%'*** Soma valores Parciais de cada Unidade ***
		if soma_unidade = 1 Then
		%>
		<%'if status_unidade <> "TREINAMENTO - " Then%>
		
		<%if colunas_mostra <> "d" Then%>
		<tr bgcolor="#CFCFCF">
			<td colspan="2">&nbsp;</td>
			<%if colunas_mostra = "c" Then%>
			<td>&nbsp;</td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td valign="top">&nbsp;<%=txt_soma_parcial%></td>
			<td align="right" valign="top"><b><%=FormatNumber(soma_h_trab_parc,2)%></b></td>
			<td align="right" valign="top">&nbsp;</td>
			<td align="right" valign="top"><b><%=FormatNumber(soma_reais_parc,2)%></b></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "c" Then%>
			<td align="right" valign="top">&nbsp;<b><%=soma_nr_qtd_transporte_parc%></b></td>
			<td align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_valor_transporte_parc,2)%></b></td>
			<td align="right" valign="top">&nbsp;<b><%=soma_nr_qtd_refeicao_parc%></b></td>
			<td align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_valor_refeicao_parc,2)%></b></td>
			<td align="right" valign="top">&nbsp;<b><%=soma_nr_qtd_convenio_parc%></b></td>
			<td align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_valor_convenio_parc,2)%></b></td>
			<td align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_valor_inss_parc,2)%></b></td>
			<%end if%>
			<td align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_nr_total_beneficios_parc,2)%></b></td>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_total_parc,2)%></b></td>
			<%end if%>
		</tr>
		<%end if%>
		<%'end if
		%>
		<%'if status_unidade <> "TREINAMENTO - " Then%>
		<tr>	
			<td colspan="7"><%'=status_unidade%></td>
		</tr>
		<%'end if
		
		'*** Zera os valores parciais de cada unidade ***
		soma_h_trab_parc = 0
		soma_reais_parc = 0
		
		soma_nr_qtd_transporte_parc = 0
		soma_valor_transporte_parc = 0
		soma_nr_qtd_refeicao_parc = 0
		soma_valor_refeicao_parc = 0
		soma_nr_qtd_convenio_parc = 0
		soma_valor_convenio_parc = 0
		soma_valor_inss_parc = 0
		soma_nr_total_beneficios_parc = 0
		
		soma_total_bruto_parc = 0
		soma_total_parc = 0
		
		end if
		
		rs.movenext
		Loop
		
		'*** Soma Total das Unidades ***
		%>
		<tr bgcolor="#CFCFCF">
			<td colspan="2">&nbsp;</td>
			<td><b>&nbsp;Soma Total das Unidades</b></td>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" OR colunas_mostra = "d" Then%>
			<td align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_h_trab,2)%></b></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td>&nbsp;</td>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(soma_reais,2)%></b></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "c" Then%>
			<td align="right" valign="top"><b><%=soma_nr_qtd_transporte%></b></td>
			<td align="right" valign="top"><b><%=FormatNumber(soma_valor_transporte,2)%></b></td>
			<td align="right" valign="top"><b><%=soma_nr_qtd_refeicao%></b></td>
			<td align="right" valign="top"><b><%=FormatNumber(soma_valor_refeicao,2)%></b></td>
			<td align="right" valign="top"><b><%=soma_nr_qtd_convenio%></b></td>
			<td align="right" valign="top"><b><%=FormatNumber(soma_valor_convenio,2)%></b></td>
			<td align="right" valign="top"><b><%=FormatNumber(soma_valor_inss,2)%></b></td>
			<%end if%>
			<%if colunas_mostra <> "d" Then%>
			<td align="right" valign="top"><b><%=FormatNumber(soma_nr_total_beneficios,2)%></b>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td align="right" valign="top"><b><%=FormatNumber(soma_total,2)%></b>
			
		</tr>
		
		<tr>	
			<td colspan="7">&nbsp;</td>
			<%end if%>
		</tr>
		<%
		
		
		%>
		<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
		<tr>
		<%
		'mes_gest = mes - 1
		'mes_inicio = mes '+ 1
		'ano_inicio = ano
		
		
		'if mes = 2 Then
		dia = UltimoDiaMes(mes,ano)
		'Else 
		'dia = 30
		'end if
		
		
		
		data_inicio = mes&"/1/"&ano
		data_fim = mes&"/"&dia&"/"&ano
		
		'xsql_gest = "up_relatorio_gestores @mes_inicio='"&mes_inicio&"',@mes='"&mes&"', @ano_inicio='"&ano_inicio&"',@ano='"&ano&"'"
		xsql_gest = "up_relatorio_gestores @data_inicio='"&data_inicio&"', @data_fim='"&data_fim&"'"
		Set rs_gest = dbconn.execute(xsql_gest)
		
		contagem = 1
		%>
		
		
			<td><b>&nbsp;</b></td>
			<td><b>&nbsp;</b></td>
			<td colspan=4><b>&nbsp;GESTORES</b></td>
		</tr> 
		<%
		
		Do While Not rs_gest.EOF
		
		'*** Dados dos Gestores ***
		g_cod = rs_gest("cd_codigo")
		g_mat = rs_gest("cd_matricula")
		g_nome = rs_gest("nm_nome")
		g_sobrenome = rs_gest("nm_sobrenome")
		g_unidades = rs_gest("cd_unidades")
		g_funcao = rs_gest("cd_funcao")
		g_dt_inicio = rs_gest("dt_inicio")
		g_dt_fim = rs_gest("dt_fim")
		g_valor_gestor = valor_gestor
		
		g_horas = rs_gest("cd_horas")
		
		'*** Horas X Valores dos Gestores ***
		g_reais = g_horas * g_valor_gestor
		
		%>
		<tr bgcolor="<%=strlinha%>" onmouseover="mOvr(this,'#eaeaea');" onmouseout="mOut(this,'FFFFFF');">
			<td align="right" valign="top"><b>&nbsp;<%=contagem%></b></td>
			<td align="right" valign="top"><b>&nbsp;</b><%=g_mat%></td>
			<td><b>&nbsp;&nbsp;</b><%=g_nome%>&nbsp;<%=g_sobrenome%></td>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td align="right" valign="top"><b>&nbsp;</b><%=TiraAcento(FormatNumber(g_horas,2))%></td>
			<td align="center" valign="top">&nbsp;<%=FormatNumber(g_valor_gestor,2)%></td>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(g_reais,2)%></b></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "c" Then%>
			<td align="center" valign="top"><b>&nbsp;</b> - </td>
			<td align="center" valign="top"><b>&nbsp;</b> - </td>
			<td align="center" valign="top"><b>&nbsp;</b> - </td>
			<td align="center" valign="top"><b>&nbsp;</b> - </td>
			<td align="center" valign="top"><b>&nbsp;</b> - </td>
			<td align="center" valign="top"><b>&nbsp;</b> - </td>
			<td align="center" valign="top"><b>&nbsp;</b> - </td>
			<%end if%>
			<td align="center" valign="top"><b>&nbsp;</b> - </td>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(g_reais,2)%></b></td>
		
		</tr>
		<%
		'*** Soma das horas e valores dos gestores ***
		soma_g_horas = soma_g_horas + ABS(g_horas)
		soma_g_reais = soma_g_reais + g_reais
		
		contagem = contagem + 1
		rs_gest.movenext
		Loop
		
		%>
		<tr>
			<td colspan="3" bgcolor="#CFCFCF"><b>&nbsp;</b></td>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td align="right" valign="top"  bgcolor="#CFCFCF"><b>&nbsp;<%=TiraAcento(FormatNumber(soma_g_horas,2))%></b></td>
			<td align="right" valign="top" bgcolor="#CFCFCF">&nbsp;<%=g1%></td>
			<td align="right" valign="top" bgcolor="#CFCFCF"><b>&nbsp;<%=FormatNumber(soma_g_reais)%></b></td>
			<%end if%>
			<%if colunas_mostra = "a" Then%>
			<td align="center" valign="top" bgcolor="#CFCFCF"><b>&nbsp;</b>-</td>
			<td align="center" valign="top" bgcolor="#CFCFCF"><b>&nbsp;</b>-</td>
			<td align="center" valign="top" bgcolor="#CFCFCF"><b>&nbsp;</b>-</td>
			<td align="center" valign="top" bgcolor="#CFCFCF"><b>&nbsp;</b>-</td>
			<td align="center" valign="top" bgcolor="#CFCFCF"><b>&nbsp;</b>-</td>
			<td align="center" valign="top" bgcolor="#CFCFCF"><b>&nbsp;</b>-</td>
			<td align="center" valign="top" bgcolor="#CFCFCF"><b>&nbsp;</b>-</td>
			<%end if%>
			<td align="center" valign="top" bgcolor="#CFCFCF"><b>&nbsp;</b>-</td>
			<td align="right" valign="top" bgcolor="#CFCFCF"><b>&nbsp;<%=FormatNumber(soma_g_reais,2)%></b></td>
			
			
		</tr>
		<tr>
			<td colspan=16>&nbsp;
		</tr>
		<tr><td colspan="21"><img src="px_black.jpg" width="100%" height="1px"></td></tr>
		<tr>
		<%
		'*** Soma das horas e valores de produtividade e gestores ***
		soma_geral_horas = soma_h_trab + soma_g_horas
		soma_geral_reais = soma_reais + soma_g_reais
		
		'*** Soma do total de valores de horas, benefícios e horas dos gestores ***
		total_nr_geral = soma_total + soma_g_reais
		%>
		
			<td align="right" valign="top">&nbsp;</td>
			<td Colspan=2 class="textopequeno" align="center" valign="top"><b>&nbsp;Total Geral:</b></td>
			<%if colunas_mostra = "a" OR colunas_mostra = "b" Then%>
			<td align="right" valign="top"><b>&nbsp;<%=TiraAcento(FormatNumber(soma_geral_horas,2))%></b></td>
			<td align="right" valign="top">&nbsp;</td>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(soma_geral_reais,2)%></b></td>
			<%end if%>
			<%if colunas_mostra = "a" OR colunas_mostra = "c" Then%>
			<td align="right" valign="top"><b>&nbsp;<%=soma_nr_qtd_transporte%></b></td>
			<td align="right" valign="top"><b>&nbsp;<%=FormatNumber(soma_valor_transporte,2)%></b></td>
			<td align="right" valign="top"><b>&nbsp;<%=soma_nr_qtd_refeicao%></b></td>
			<td bgcolor="#FFFFFF"align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_valor_refeicao,2)%></b></td>
		
			<td bgcolor="#FFFFFF"align="right" valign="top">&nbsp;<b><%=soma_nr_qtd_convenio%></b></td>
			<td bgcolor="#FFFFFF"align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_valor_convenio,2)%></b></td>
		
			<td bgcolor="#FFFFFF"align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_valor_inss,2)%></td>
			<%end if%>
			<td bgcolor="#FFFFFF"align="right" valign="top">&nbsp;<b><%=FormatNumber(soma_nr_total_beneficios,2)%></b></td>
			<td bgcolor="#FFFFFF"align="right" valign="top">&nbsp;<b><%=FormatNumber(total_nr_geral,2)%></b></td>
			
		</tr>
		<tr><td colspan="21"><img src="px_black.jpg" width="100%" height="1px"></td></tr>
		</table>
		<Table width="<%=largura_col%>" border="0" cellspacing="1" cellpadding="0" >
		<tr>
			<td  colspan=18 width="98%" height="2" bgcolor=#FFFFFF><td>
		</tr>
		
		<tr onmouseover="mOvr(this,'#eaeaea');" onmouseout="mOut(this,'FFFFFF');">
		<%
		'Serviços Prestados: Produtividade + Benefícios
		serv_pres = total_nr_geral
		
		'Taxa Administrativa: 10% sobre a Produtividade
		tax = soma_geral_reais * 0.1
		
		'Imposto de Renda: 1,50% sobre (Serviços Prestados + Taxa Administrativa) 
		ir = serv_pres + tax
		ir = ir * 0.015
		
		'Valor Liquido da Nota Fiscal: Serviços Prestados + Taxa administrativa - Imposto de Renda
		nf = serv_pres + tax
		nf = nf - ir
		%>
			<td></td>
			<td colspan=12 class="textopequeno" align="center" valign="top"><b>&nbsp;VALOR SERVIÇOS PRESTADOS (produtividade + benefícios)</b></td>
			<td colspan=7 class="textopequeno">&nbsp;</td>
			<td colspan=2 class="textopequeno" align="right" valign="top"><b>&nbsp;<%=FormatNumber(serv_pres,2)%></b></td>
		</tr>
		<tr onmouseover="mOvr(this,'#eaeaea');" onmouseout="mOut(this,'FFFFFF');">
			<td></td>
			<td colspan=12 class="textopequeno" align="center" valign="top"><b>&nbsp;TAXA ADMINISTRATIVA 10% (sobre a produtividade)</b></td>
			<td colspan=7 class="textopequeno"><b>&nbsp;</b></td>
			<td colspan=2 class="textopequeno" align="right" valign="top"><b><%=FormatNumber(tax,2)%></b></td>
		</tr>
		
		<tr onmouseover="mOvr(this,'#eaeaea');" onmouseout="mOut(this,'FFFFFF');">
			<td></td>
			<td colspan=12 class="textopequeno" align="center" valign="top"><b>&nbsp;IMPOSTO DE RENDA 1,50%</b></td>
			<td colspan=7 class="textopequeno"><b>&nbsp;</b></td>
			<td colspan=2 class="textopequeno" align="right" valign="top"><b><%=FormatNumber(ir,2)%></b></td>
		</tr>
		<!--tr>
			<td  colspan=16 width="98%" height="2" bgcolor=#FFFFFF>&nbsp;<td>
		</tr-->
		<tr><td colspan="100%"><img src="px_black.jpg" width="100%" height="1px"></td></tr>
		<tr><td colspan="100%"><img src="imagens/px.gif" width="100%" height="1px"></td></tr>
		<tr onmouseover="mOvr(this,'#eaeaea');" onmouseout="mOut(this,'FFFFFF');">
			<td><b>&nbsp;</b></td>
			<td colspan=12 class="textopequeno" align="center" valign="top"><b>&nbsp;VALOR LÍQUIDO DA NOTA FISCAL</b></td>
			<td colspan=7 class="textopequeno"><b>&nbsp;</b></td>
			<td colspan=2 class="textopequeno" align="right" valign="top"><b><%=FormatNumber(nf,2)%></b></td>
		</tr>
		<!--tr>
			<td Colspan=16 height="15"><img src="px.gif" alt="" width="1" height="26" border="0"></td>
		</tr>
		<tr>
			<td  colspan=21 width="100%" bgcolor=#c0c0c0><td>
		</tr-->
		<%if relatorio = "" Then%>
		<tr>
			
			<td>&nbsp;</td>
			<td colspan=10 class="textopadrao">
			<a href="relatorios/cooperativa/cooperativa.asp?mes=<%=mes%>&ano=<%=ano%>&unit=k&relatorio=compl" target="blanck">Imprimir - Relatório completo</a><br>
			<a href="relatorios/cooperativa/cooperativa.asp?mes=<%=mes%>&ano=<%=ano%>&unit=k&relatorio=prod_ben" target="blanck">Imprimir - Produtividade e Benefícios</a><br>
			<a href="relatorios/cooperativa/cooperativa.asp?mes=<%=mes%>&ano=<%=ano%>&unit=k&relatorio=beneficios" target="blanck">Imprimir - Benefícios em Detalhes</a><br><br><br>
			<a href="relatorios/cooperativa/cooperativa.asp?mes=<%=mes%>&ano=<%=ano%>&unit=k&relatorio=consolidado" target="blanck">Imprimir - Consolidado</a><br><br>
			</td>
			<td>&nbsp;<br><br><br></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<%end If%>
		<%end If%>
		<%end if%>
<%end if%>
		<%
		'rs.close
		'Set rs = Nothing
		%>
</body>
</html>