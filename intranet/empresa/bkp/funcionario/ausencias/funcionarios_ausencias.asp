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

nome = request("nome")
	if nome <> "" Then
		busca_nome = "AND nm_nome like '"&nome&"%'"
	end if 
sexo = request("sexo")
	if sexo <> "0" Then
		busca_sexo = "AND cd_sexo = '"&sexo&"'"
	end if 

%>


<head>
	<title>Listagem de funcionários</title>
</head>

<body><br>

<table align="center" border="1">
	<tr><td colspan="100%" align="center" style="background-color:gray; color:white; font-weight:bold;">AUSÊNCIAS</td></tr>
	<form action="../../empresa.asp" method="get">
	<input type="hidden" name="tipo" value="ausencia">
	<tr>
		<td>Data</td>
		<td>Contratos</td>
		<td>Mostrar</td>
		<td>Detalhes</td>
		<td>Ordem</td>
	</tr>
	<tr id="no_print">
		<td valign="top">
			<input type="text" name="dia_sel" size="2" maxlength="2" value="<%=zero(day(data_selecionada))%>">
			<input type="text" name="mes_sel" size="2" maxlength="2" value="<%=zero(month(data_selecionada))%>">
			<input type="text" name="ano_sel" size="4" maxlength="4" value="<%=year(data_selecionada)%>">
		</td>
		<%if cd_situacao = "" Then
			ck_situacao_1 = "checked"
		end if%>
		<td valign="top">
			<input type="radio" name="cd_situacao" value="1" <%=ck_situacao_1%> style=border:0px;>Ativo<br>
			<input type="radio" name="cd_situacao" value="2" style=border:0px;>Inativo
		</td>
		<td valign="top">&nbsp;<br>
			<input type="checkbox" name="campos" value="sexo" <%if campo_sex <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Sexo<br>
			<input type="checkbox" name="campos" value="endereco" <%if campo_endereco <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Endereço<br>
			<input type="checkbox" name="campos" value="rg" <%if campo_rg <> 0 Then response.write("CHECKED") end if%> style=border:0px;> RG<br>
			<input type="checkbox" name="campos" value="cpf" <%if campo_cpf <> 0 Then response.write("CHECKED") end if%> style=border:0px;> CPF<br>
			<input type="checkbox" name="campos" value="rgp" <%if campo_rgp <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Registro Prof.<br>
			<input type="checkbox" name="campos" value="cargo" <%if campo_cargo <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Cargo<br>
			<!--input type="checkbox" name="campos" value="unidade" <%if campo_unidade <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Unidade<br-->
			<input type="checkbox" name="campos" value="situacao" <%if campo_situacao <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Situacao
		</td>
		<td valign="top">
			<input type="text" name="nome" size="20" maxlength="100" value="<%=nome%>">Funcionário<br>
			<%if sexo = 1 Then%><%ck_sx1 = "checked"%><%Elseif sexo = 2 then%><%ck_sx2 = "checked"%><%else%><%ck_sx = "checked"%><%End if%>
			<input type="radio" name="sexo" value="1" <%=ck_sx1%>>M&nbsp;<input type="radio" name="sexo" value="2" <%=ck_sx2%>>F&nbsp;<input type="radio" name="sexo" value="0" <%=ck_sx%>>Ambos<br>
		</td>
		<%'*** ORDEM ***
		if ordem_funcionarios = "" or ordem_funcionarios = 1 Then
			ordem_funcionarios = "cd_contrato,nm_nome,dt_admissao"
			ordem_nome = "checked"
		elseif ordem_funcionarios = 2 then
			ordem_funcionarios = "cd_contrato,dt_admissao,nm_nome"
			ordem_admissao = "checked"
		end if%>
		<td><input type="radio" name="ordem_funcionarios" value="1" <%=ordem_nome%> style=border:0px;>Nome<br>
			<input type="radio" name="ordem_funcionarios" value="2" <%=ordem_admissao%> style=border:0px;>Admissao
		</td>
	</tr>
		
	
	<!--tr><td></td></tr>
	<tr><td></td></tr>
	<tr><td></td></tr>
	<tr><td></td></tr-->
	
		
		
		
				
	</tr>
	
	
	
	<tr id="no_print">
		<td><input type="submit" name="OK" value="OK"></td>
		<td>&nbsp;<a href="empresa.asp?tipo=novo"><img src="../../imagens/ic_novo.gif" alt="" width="10" height="12" border="0">Novo colaborador</a></td>
	</tr>
	</form>
</table>	
<table align="center">
	<tr>
					<%if cd_situacao = 1 OR cd_situacao = "" then
						xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL "&busca_nome&" "&busca_sexo&" ORDER BY "&ordem_funcionarios&""
					elseif cd_situacao = 2 then
						'xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao >= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL ORDER BY cd_contrato, nm_nome"
						xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_demissao is not NULL ORDER BY "&ordem_funcionarios&""
					end if
						Set rs_func = dbconn.execute(xsql)
							
							while not rs_func.EOF
								conta_func = conta_func + 1	
								
								cd_funcionario = rs_func("cd_funcionario")
								nm_funcionario = rs_func("nm_nome")
								
								cd_sexo = rs_func("cd_sexo")
									if cd_sexo = 1 then
										nm_sexo = "Masc"
									elseif cd_sexo = 2 then
										nm_sexo = "Fem"
									else
										nm_sexo = "Indeterminado"
									end if
									
								nm_endereco = rs_func("nm_endereco")
								nr_numero = rs_func("nr_numero")
								nm_bairro = rs_func("nm_bairro")
								nm_cidade = rs_func("nm_cidade")
								nm_estado = rs_func("nm_estado")
								nm_cep = rs_func("nm_cep")
								
								nm_rg = rs_func("nm_rg")
								nm_cpf = rs_func("nm_cpf")
								
								dt_admissao = zero(day(rs_func("dt_admissao")))&"/"&zero(month(rs_func("dt_admissao")))&"/"&year(rs_func("dt_admissao"))
								dt_demissao = zero(day(rs_func("dt_demissao")))&"/"&zero(month(rs_func("dt_demissao")))&"/"&year(rs_func("dt_demissao"))
								
								
								
								cd_contrato = rs_func("cd_contrato")
								
								xsql_regprof = "Select * From TBL_funcionario Where cd_codigo = "&cd_funcionario&""
								Set rs_regprof = dbconn.execute(xsql_regprof)
									if not rs_regprof.EOF then
										cd_numreg = rs_regprof("cd_numreg")
									end if
								
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
											<td colspan="100%">&nbsp;</td>										
										</tr>
										<tr>
											<td colspan="100%" align="center" style="background-color:gray; color:white;"><%'=cd_contrato%><b><%=nm_regime_trab%></b></td>
										</tr>
										<tr bgcolor=#c0c0c0>
											<td>&nbsp;</td>
											<td><b>Funcionário</b></td>
											<%if campo_sex <> 0 Then%><td><b>Sexo</b></td> <%end if%>
											<%if campo_endereco <> 0 Then%><td><b>Endereço</b></td><%end if%>
											<%if campo_rg <> 0 Then%><td><b>RG</b></td><%end if%>
											<%if campo_cpf <> 0 Then%><td><b>CPF</b></td><%end if%>
											<%if campo_rgp <> 0 Then%><td><b>Reg. Profissional</b></td><%end if%>
											<%if campo_cargo <> 0 Then%><td><b>Cargo</b></td><%end if%>
											<%if campo_hora <> 0 Then%><td><b>Horário</b></td><%end if%>
											<%'if campo_unidade <> 0 Then%><td><b>Unidade</b></td><%'end if%>
											<%if campo_situacao <> 0 Then%><td><b>Situação</b></td><%end if%>
											<%if cd_situacao = 2 then%>
											<td><b>Empresa</b></td>
											<%end if%>
											<td><b>Admissão</b></td>
											<%if cd_situacao = 2 then%>
											<td><b>demissão</b></td>
											<%end if%>
										</tr>
									<%end if%>
								
								<%'*** Lista os funcionarios ***%>
								<!--tr-->
								<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa.asp?tipo=cad_ausencia&cod=<%=rs_func("cd_funcionario")%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px;" bgcolor="<%=cor_linha%>">
									<td align="center"><%=zero(conta_func)%></td>
									<td valign="top"><%=nm_funcionario%></td>
									<%if campo_sex <> 0 Then%><td><%=nm_sexo%></td><%end if%>
									<%if campo_endereco <> 0 Then%><td><%=nm_endereco%>, <%=nr_numero%> - <%=nm_bairro%> - <%=nm_cidade%>-<%=nm_estado%>&nbsp;</td><%end if%>
									<%if campo_rg <> 0 Then%><td><%=nm_rg%></td><%end if%>
									<%if campo_cpf <> 0 Then%><td><%=nm_cpf%></td><%end if%>
									<%if campo_rgp <> 0 Then%><td><%=cd_numreg%></td><%end if%>
									<%if campo_cargo <> 0 Then%><td><%=nm_cargo%></td><%end if%>
									<%if campo_hora <> 0 Then%><td><%=hr_entrada%> - <%=hr_saida%></td><%end if%>
									<%'if campo_unidade <> 0 Then%><td><%=nm_sigla%></td><%'end if%>
									<%if campo_situacao <> 0 Then%><td><%=nm_status%></td><%end if%>
								<%if cd_situacao = 2 then%>
									<td><%=nm_regime_trab%></td>
								<%end if%><td><%=dt_admissao%></td>
								<%if cd_situacao = 2 then%>
									<td><%=dt_demissao%></td>
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
			<td><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
			<td><img src="../../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
			<%if campo_sex <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td><%end if%>	
			<%if campo_endereco <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="250" height="1" border="0"></td><%end if%>
			
			<%if campo_rg <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td><%end if%>
			<%if campo_cpf <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td><%end if%>
			
			<%if campo_rgp <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="115" height="1" border="0"></td><%end if%>
			<%if campo_cargo <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="115" height="1" border="0"></td><%end if%>
			<%if campo_hora <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td><%end if%>
			<%'if campo_unidade <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="45" height="1" border="0"></td><%'end if%>	
			<%if campo_situacao <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="75" height="1" border="0"></td><%end if%>
			<td><img src="../../imagens/px.gif" alt="" width="70" height="3" border="0"></td>
		</tr>
		<tr>
			<td><%'=lista_funcionario%><%'=remanescentes%></td>
		</tr>
		<td></td>
	</tr>
</table>


</body>
</html>



