<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<%'*** Datas selecionadas ***
dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")
	if dia_sel = "" Then
		dia_sel = zero(day(now))
	end if
	
	if mes_sel = "" Then
		mes_sel = zero(month(now))
	end if
	
	if ano_sel = "" Then
		ano_sel = year(now)
	end if

	dia_final = ultimodiames(mes_sel,ano_sel)

data_selecionada = dia_sel&"/"&mes_sel&"/"&ano_sel
periodo_sel = ano_sel&mes_sel

	dia_hoje = day(now)
	mes_hoje = month(now)
	ano_hoje = year(now)
data_hoje = zero(dia_hoje)&"/"&zero(mes_hoje)&"/"&ano_hoje

mes_anterior = ano_sel&(mes_sel-1)

'****************************************************
cd_situacao = request("cd_situacao")
ordem_funcionarios = request("ordem_funcionarios")
	if ordem_funcionarios = "" Then
		ordem_funcionarios = 1
	end if

campos = request("campos")
	campo_sex = Instr(1,campos,"sexo",0)
	campo_endereco = Instr(1,campos,"endereco",0)
	campo_rg = Instr(1,campos,"rg",0)
	campo_cpf = Instr(1,campos,"cpf",0)
	campo_rgp = Instr(1,campos,"rgp",0)
	campo_cargo = Instr(1,campos,"cargo",0)
	campo_hora = Instr(1,campos,"hora",0)
	campo_unidade = Instr(1,campos,"unidade",0)
	campo_situacao = Instr(1,campos,"situacao",0)
	campo_ctps = Instr(1,campos,"ctps",0)
	campo_contratos = Instr(1,campos,"contratos",0)
	campo_tempo_casa = Instr(1,campos,"tempo_casa",0)
nome = request("nome")
	if nome <> "" Then
		busca_nome = "AND nm_nome like '"&nome&"%'"
	end if 
sexo = request("sexo")
	if sexo > "0" Then
		busca_sexo = "AND cd_sexo = '"&sexo&"'"
	end if 
	

 
%>
<head>
	<title>Listagem de funcionários</title>
</head>

<body><br>

<table align="center" border="1" id="no_print">
	<tr><td colspan="100%" align="center" style="background-color:gray; color:white; font-weight:bold;">LISTAGEM</td></tr>
	<form action="../../empresa.asp" method="post">
	<input type="hidden" name="tipo" value="mov_anual">
	<tr id="no_print">
		<td valign="top" align="center" colspan="2">ANO: 
			<input type="text" name="ano_sel" size="4" maxlength="4" value="<%=ano_sel%>">
		</td>		
	</tr>		
	<tr id="no_print">
		<td><input type="submit" name="OK" value="Mostrar"></td>
		<td><img src="../../imagens/ic_print.gif" alt="" width="24" height="24" border="0" onClick="window.print();">&nbsp;
		<img src="../../imagens/ic_print_view.gif" alt="" width="24" height="26" border="0" onclick="visualizarImpressao();">
		</td>
	</tr>
	</form>
</table><br>
<table style="border-color:white;" align="center">
	<tr>
		<td id="ok_print"><img src="../../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
		<td>
<table align="center" border="1" width="400" style="border-collapse:collapse;">
	
		<tr><td colspan="100%" style="text-align:center; font-weight:bold; font-size:12px; background-color:gray; color:white;">Movimentação anual de colaboradores - <%=ano_sel%></td></tr>
		<tr><td colspan="100%">&nbsp;</td></tr>
		<%'*** verifica quais contratos estavam ativos ***
				xsql = "SELECT DISTINCT cd_contrato, nm_regime_trab FROM VIEW_funcionario_contrato_lista WHERE year(dt_admissao) <= '"&ano_sel&"' AND year(dt_demissao) > '"&ano_sel&"' OR year(dt_admissao) <= '"&ano_sel&"' AND year(dt_demissao) IS NULL GROUP BY cd_contrato,nm_regime_trab"
					Set rs_contr = dbconn.execute(xsql)
						while not rs_contr.EOF
						cd_contrato = rs_contr("cd_contrato")
						nm_regime_trab = rs_contr("nm_regime_trab")%>
		<tr><td colspan="100%" style="background-color:silver; color:black; font-weight:bold; text-align:center;"><%=nm_regime_trab%></td></tr>
					
		<tr>			
		<%for i = 1 to  12%>
		<%if i = 1 then%>
			<td valign="top">&nbsp;<br>
				Saldo<br>
				Admissão<br>
				Recisão<br>
				Total
			</td>
		<%end if%>
		<td align="right" valign="top">
			<!--table align="center" width="10"-->
				<%'*** conta os funcionários de cada contrato ***
						'*** seleciona o mês anterior ***
							if i = 1 then
								mes_anterior = 12
								ano_anterior = ano_sel - 1
							else
								mes_anterior = i - 1
								ano_anterior = ano_sel
							end if
							
						data_sel = ano_anterior&zero(mes_anterior)
				
				total_funcionarios = 0
				
				%>
				<!--tr-->
					<%if i = 0 then%>
						<!--td valign="top"-->&nbsp;<br>
						S<br>
						A<br>
						R<br>
						T
						<!--/td-->
					<%end if%>
					<!--td style="background-color:white; color:black;" colspan="3"--><b><%'=nm_regime_trab%><%=left(ucase(mes_selecionado(i)),3)%></b><br><!--/td>
				</tr-->
					<%	xsql = "SELECT COUNT(admissao) AS conta, cd_contrato FROM VIEW_funcionario_contrato_lista WHERE (admissao <= '"&data_sel&"') AND (demissao > '"&data_sel&"') AND (cd_contrato = "&cd_contrato&") OR (admissao <= '"&data_sel&"') AND (dt_demissao IS NULL) AND (cd_contrato = "&cd_contrato&") GROUP BY cd_contrato"
							Set rs_conta = dbconn.execute(xsql)
								while not rs_conta.EOF
									conta = rs_conta("conta")
								rs_conta.movenext
								wend
								
									xsql = "SELECT COUNT(cd_contrato) AS conta_admissao FROM VIEW_funcionario_contrato_lista WHERE (admissao = '"&ano_sel&zero(i)&"') AND (cd_contrato = "&cd_contrato&")"
									Set rs_admiss = dbconn.execute(xsql)
										if not rs_admiss.EOF Then
											total_admissao = rs_admiss("conta_admissao")
										end if
									
									xsql = "SELECT COUNT(cd_contrato) AS conta_recisao FROM VIEW_funcionario_contrato_lista WHERE (demissao  = '"&ano_sel&zero(i)&"') AND (cd_contrato = "&cd_contrato&")"
									Set rs_recisao = dbconn.execute(xsql)
										if not rs_recisao.EOF Then
											total_recisao = rs_recisao("conta_recisao")
										end if%>
					 
					<!--tr>						
						<td colspan="3" align="right"-->
						<b><%=int(conta)%></b> <%'="("&data_sel&")"%><br>
						<b style="color:green;"><%=total_admissao%></b><br>
						<b style="color:red;"><%=total_recisao%></b><br>
						<b><%total_geral=(conta+total_admissao)-total_recisao%> <%=total_geral%></b>
					<!--tr><td><%=data_sel%></td></tr-->
				<%total_funcionarios = total_funcionarios + total_geral
				
				conta = 0
				%>
				<!--tr><td colspan="3"><hr></td></tr>
				<tr><td colspan="3"><b><%=total_funcionarios%></b></td></tr>
				<tr>
					<td><%'=lista_funcionario%><%'=remanescentes%></td>
				</tr-->			
		<!--/table>	
	</td-->
					
	<%if i= 1 then
		t_jan = t_jan + total_funcionarios
	elseif i = 2 then
		t_fev = t_fev + total_funcionarios
	elseif i = 3 then
		t_mar = t_mar + total_funcionarios
	elseif i = 4 then
		t_abr = t_abr + total_funcionarios
	elseif i = 5 then
		t_mai = t_mai + total_funcionarios
	elseif i = 6 then
		t_jun = t_jun + total_funcionarios
	elseif i = 7 then
		t_jul = t_jul + total_funcionarios
	elseif i = 8 then
		t_ago = t_ago + total_funcionarios
	elseif i = 9 then
		t_set = t_set + total_funcionarios
	elseif i = 10 then
		t_out = t_out + total_funcionarios
	elseif i = 11 then
		t_nov = t_nov + total_funcionarios
	elseif i = 12 then
		t_dez = t_dez + total_funcionarios
	end if
	
	t_admissao = t_admissao + total_admissao
	t_recisao = t_recisao + total_recisao
	t_anual = t_anual + total_funcionarios
	
	next%>
	<td align="right">
		&nbsp;<br>
		&nbsp;<br>		
		&nbsp;<b style="color:green;"><%=t_admissao%></b><br>
		&nbsp;<b style="color:red;"><%=t_recisao%></b><br>
		&nbsp;<%'=t_anual%>		
	</td>
	</tr>
	
	<tr><td colspan="100%">&nbsp;</td></tr>
	<%rs_contr.movenext
	wend%>	
	<tr style="font-weight:bold;">
		<td>Head Count<br><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td align="center"><%=t_jan%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_fev%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_mar%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_abr%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_mai%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_jun%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_jul%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_ago%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_set%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_out%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_nov%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center"><%=t_dez%><br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td align="center">&nbsp;<br><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
	</tr>
	<!--tr>
		<td><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<%for i=1 to 12%><td><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td><%next%>
	</tr-->
</table>

</td></tr></table>
</body>
</html>
