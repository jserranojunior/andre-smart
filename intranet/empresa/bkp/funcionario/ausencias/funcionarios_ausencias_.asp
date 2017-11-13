<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<%
dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

if dia_sel = "" or mes_sel = "" or ano_sel = "" then
	dia_sel = day(now)
	mes_sel = month(now)
	ano_sel = year(now)
end if

data_selecionada = dia_sel&"/"&mes_sel&"/"&ano_sel

'****************************************************
cd_situacao = request("cd_situacao")
ordem_funcionarios = request("ordem_funcionarios")

'data_selecionada = request("data_selecionada")
'data_selecionada = month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)
%>
<head>
	<title>Untitled</title>
</head>

<body><br>

<table align="center" border="">
	<form action="../../empresa.asp" method="post">
	<input type="hidden" name="tipo" value="lista">
	<tr id="no_print">
		<td colspan="2">
			<input type="text" name="dia_sel" size="2" maxlength="2" value="<%=zero(day(data_selecionada))%>">
			<input type="text" name="mes_sel" size="2" maxlength="2" value="<%=zero(month(data_selecionada))%>">
			<input type="text" name="ano_sel" size="4" maxlength="4" value="<%=year(data_selecionada)%>">
		</td>
		<%if cd_situacao = "" Then
			ck_situacao_1 = "checked"
		end if%>
		<td>
			<input type="radio" name="cd_situacao" value="1" <%=ck_situacao_1%>>Ativo<br>
			<input type="radio" name="cd_situacao" value="2">Inativo
		</td>
		<%'*** ORDEM ***
		if ordem_funcionarios = "" or ordem_funcionarios = 1 Then
			ordem_funcionarios = "cd_contrato,nm_nome,dt_admissao"
			ordem_nome = "checked"
		elseif ordem_funcionarios = 2 then
			ordem_funcionarios = "cd_contrato,dt_admissao,nm_nome"
			ordem_admissao = "checked"
		end if%>
		<td>Ordenar por:<br>
			<input type="radio" name="ordem_funcionarios" value="1" <%=ordem_nome%>>Nome<br>
			<input type="radio" name="ordem_funcionarios" value="2" <%=ordem_admissao%>>Admissao
		</td>
				
	</tr>
	<tr id="no_print">
		<td><input type="submit" name="OK" value="OK"></td>
	</tr>
	</form>	
		<tr>
						<%
						'SELECT * FROM VIEW_funcionario_contrato_lista WHERE (dt_admissao <= '11/8/2010') AND (dt_demissao >= '11/8/2010') OR (dt_admissao <= '11/8/2010') AND (dt_demissao IS NULL) ORDER BY cd_contrato, nm_nome
					'if ordem_funcionarios_ativos = "" then
					'	ordem_funcionarios_ativos = "cd_contrato,nm_nome,dt_admissao"
					'end if
					'ordem_funcionarios_ativos = "cd_contrato,dt_admissao,nm_nome"
					'ordem_funcionarios_inativos = "nm_nome,dt_demissao"
					
					if cd_situacao = 1 OR cd_situacao = "" then
						xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL ORDER BY "&ordem_funcionarios&""
					elseif cd_situacao = 2 then
						'xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao >= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL ORDER BY cd_contrato, nm_nome"
						xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_demissao is not NULL ORDER BY "&ordem_funcionarios&""
					end if
						Set rs_func = dbconn.execute(xsql)
							
							while not rs_func.EOF
								conta_func = conta_func + 1	
								
								cd_funcionario = rs_func("cd_funcionario")
								nm_funcionario = rs_func("nm_nome")
								dt_admissao = zero(day(rs_func("dt_admissao")))&"/"&zero(month(rs_func("dt_admissao")))&"/"&year(rs_func("dt_admissao"))
								dt_demissao = zero(day(rs_func("dt_demissao")))&"/"&zero(month(rs_func("dt_demissao")))&"/"&year(rs_func("dt_demissao"))
								
								cd_contrato = rs_func("cd_contrato")
								
								'xsql_regprof = "Select * From TBL_funcionario Where cd_codigo = "&cd_funcionario&""
								'Set rs_regprof = dbconn.execute(xsql_regprof)
								'	if not rs_regprof.EOF then
								'		cd_numreg = rs_regprof("cd_numreg")
								'	end if
								
								'*** CARGO ***
								xsql_cargo = "Select * From View_funcionario_cargo Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 ORDER BY dt_atualizacao desc"
								Set rs_cargo = dbconn.execute(xsql_cargo)
									if not rs_cargo.EOF then
										nm_cargo = rs_cargo("nm_cargo")
									end if
								
								'*** HORÁRIO ***
								xsql = "Select * From View_funcionario_horario Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 order by dt_atualizacao DESC"
								Set rs_hora = dbconn.execute(xsql)
									if not rs_hora.EOF Then
										hr_entrada = zero(hour(rs_hora("hr_entrada")))&":"&zero(minute(rs_hora("hr_entrada")))
										hr_saida = zero(hour(rs_hora("hr_saida")))&":"&zero(minute(rs_hora("hr_saida")))
										nm_intervalo = rs_hora("nm_intervalo")
									end if
								
								'*** UNIDADE ***
								xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 order by dt_atualizacao DESC"
								Set rs_unid = dbconn.execute(xsql)
									if not rs_unid.EOF Then
										nm_unidade = rs_unid("nm_unidade")
										nm_sigla = rs_unid("nm_sigla")
									end if
									
								'*** STATUS ***
								xsql = "Select * From View_funcionario_status Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 order by dt_atualizacao DESC"
								Set rs_status = dbconn.execute(xsql)
									if not rs_status.EOF then
										nm_status = rs_status("nm_status")
									end if
								
								'*** CONTRATO (EMPRESA) ***
								xsql = "Select * From TBL_tipo_contrato Where cd_codigo = "&cd_contrato&""
								Set rs_contrato = dbconn.execute(xsql)
									if not rs_contrato.EOF then
										nm_regime_trab = rs_contrato("nm_regime_trab")
									end if
								
								'*** Mostra o cabeçalho ****
									if cabeca_ativo <>  cd_contrato then%>
										<tr id="no_print">
											<td colspan="7">&nbsp;</td>											
										</tr>
										<tr>
											<td colspan="100%" align="center"><%'=cd_contrato%><b><%=nm_regime_trab%></b></td>
										</tr>
										<tr bgcolor=#c0c0c0>
											<td>&nbsp;</td>
											<td>&nbsp;<b>Funcionário</b></td>
											<td>&nbsp;<b>Cargo</b></td>
											<td>&nbsp;<b>Horário</b></td>
											<td>&nbsp;<b>Unidade</b></td>
											<td>&nbsp;<b>Situação</b></td>
											<%if cd_situacao = 2 then%>
											<td>&nbsp;<b>Empresa</b></td>
											<%end if%>
											<td>&nbsp;<b>Admissão</b></td>
											<%if cd_situacao = 2 then%>
											<td>&nbsp;<b>demissão</b></td>
											<%end if%>
										</tr>
									<%end if%>
								
								<%'*** Lista os funcionarios ***%>
								<!--tr-->
								<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa.asp?tipo=cad_ausencia&cod=<%=rs_func("cd_funcionario")%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px;" bgcolor="<%=cor_linha%>">
									<td align="center"><%=zero(conta_func)%></td>
									<td>&nbsp;<%=nm_funcionario%></td>
									<td>&nbsp;<%=nm_cargo%></td>
									<td>&nbsp;<%=hr_entrada%> - <%=hr_saida%></td>
									<td>&nbsp;<%=nm_sigla%></td>
									<td>&nbsp;<%=nm_status%></td>
								<%if cd_situacao = 2 then%>
									<td>&nbsp;<%=nm_regime_trab%></td>
								<%end if%><td>&nbsp;<%=dt_admissao%></td>
								<%if cd_situacao = 2 then%>
									<td>&nbsp;<%=dt_demissao%></td>
								<%end if%>
								</tr>							
								
								<%'if cd_situacao = 2 and pulo_linha_inativos = cd_funcionario then%>
								<!--tr>
									<td>&nbsp;</td>
								</tr-->
								<%cabeca_ativo = cd_contrato

							rs_func.movenext
							wend%>				
					<%'cabeca = ""
					
					
					lista_empresa = ""
'*******************************************************%>
							<!--tr>
								<td><%=lista_empresa%>...</td>
							</tr-->
							
					
						

		<%'next
		'rs.movenext
		'wend%>
		<tr>
			<td><img src="../../imagens/blackdot.gif" alt="" width="25" height="1" border="0"></td>
			<td><img src="../../imagens/blackdot.gif" alt="" width="200" height="1" border="0"></td>
			<td><img src="../../imagens/blackdot.gif" alt="" width="115" height="1" border="0"></td>
			<td><img src="../../imagens/blackdot.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../../imagens/blackdot.gif" alt="" width="45" height="1" border="0"></td>			
			<td><img src="../../imagens/blackdot.gif" alt="" width="75" height="1" border="0"></td>
			<td><img src="../../imagens/blackdot.gif" alt="" width="75" height="1" border="0"></td>
		</tr>
		<tr>
			<td><%'=lista_funcionario%><%'=remanescentes%></td>
		</tr>
		<td></td>
	</tr>
</table>


</body>
</html>
