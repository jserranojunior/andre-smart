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
	<input type="hidden" name="tipo" value="movimentacao">
	<tr>
		<td>Data</td>
		<td>Contratos</td>
		<td colspan="2">Mostrar</td>
		<td>Detalhes</td>
		<td>Ordem</td>
	</tr>
	<tr id="no_print">
		<td valign="top" align="center">DE: 
			<!--input type="text" name="dia_inicio" size="2" maxlength="2" value="<%=dia_inicio%>">
			<input type="text" name="mes_inicio" size="2" maxlength="2" value="<%=mes_inicio%>">
			<input type="text" name="ano_inicio" size="4" maxlength="4" value="<%=ano_inicio%>"><br><br>
			ATÉ
			<input type="text" name="dia_final" size="2" maxlength="2" value="<%=dia_final%>">
			<input type="text" name="mes_final" size="2" maxlength="2" value="<%=mes_final%>">
			<input type="text" name="ano_final" size="4" maxlength="4" value="<%=ano_final%>"><br>
			<br-->
			<!--input type="text" name="dia_sel" size="2" maxlength="2" value="<%=dia_sel%>"-->
			<input type="text" name="mes_sel" size="2" maxlength="2" value="<%=mes_sel%>">
			<input type="text" name="ano_sel" size="4" maxlength="4" value="<%=ano_sel%>"><br><br>
			
			
			<%=dia_sel%>/<%=mes_sel%>/<%=ano_sel%> <br>
			(<%=ano_sel&mes_sel%>)
		</td>
		<%if cd_situacao = "" OR cd_situacao = 1 Then
			ck_situacao_1 = "checked"
		elseif cd_situacao = 2 Then
			ck_situacao_2 = "checked"
		end if%>
		<td valign="top">
			<input type="radio" name="cd_situacao" value="1" <%=ck_situacao_1%> style=border:0px;>Ativo<br>
			<input type="radio" name="cd_situacao" value="2" <%=ck_situacao_2%> style=border:0px;>Inativo
		</td>
		<td valign="top">
			<input type="checkbox" name="campos" value="sexo" <%if campo_sex <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Sexo<br>
			<input type="checkbox" name="campos" value="endereco" <%if campo_endereco <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Endereço<br>
			<input type="checkbox" name="campos" value="rg" <%if campo_rg <> 0 Then response.write("CHECKED") end if%> style=border:0px;> RG<br>
			<input type="checkbox" name="campos" value="cpf" <%if campo_cpf <> 0 Then response.write("CHECKED") end if%> style=border:0px;> CPF<br>
			<input type="checkbox" name="campos" value="rgp" <%if campo_rgp <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Registro Prof.<br>
			<input type="checkbox" name="campos" value="cargo" <%if campo_cargo <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Cargo<br>
			<input type="checkbox" name="campos" value="unidade" <%if campo_unidade <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Unidade<br>
			<input type="checkbox" name="campos" value="situacao" <%if campo_situacao <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Situacao<br>
			<input type="checkbox" name="campos" value="ctps" <%if campo_ctps <> 0 Then response.write("CHECKED") end if%> style=border:0px;> CTPS<br>			
		</td>
		<td valign="top">
			<input type="checkbox" name="campos" value="tempo_casa" <%if campo_tempo_casa <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Tempo de casa<br>
			<input type="checkbox" name="campos" value="contratos" <%if campo_contratos <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Contratos<br>
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
		elseif ordem_funcionarios = 3 then
			ordem_funcionarios = "tempo_casa DESC"
			ordem_tempo = "checked"
		end if%>
		<td><input type="radio" name="ordem_funcionarios" value="1" <%=ordem_nome%> style=border:0px;>Nome<br>
			<input type="radio" name="ordem_funcionarios" value="2" <%=ordem_admissao%> style=border:0px;>Admissao<br>
			<input type="radio" name="ordem_funcionarios" value="3" <%=ordem_tempo%> style=border:0px;>Tempo de casa<br>
			<br>
			<%=ordem_funcionarios%>
		</td>
	</tr>		
	<tr id="no_print">
		<td><input type="submit" name="OK" value="Mostrar"></td>
		<td>&nbsp;<a href="empresa.asp?tipo=novo"><img src="../../imagens/ic_novo.gif" alt="" width="10" height="12" border="0">Novo colaborador</a></td>
	</tr>
	</form>
</table>	
<table align="center">
	<tr>
					<%cor_linha = "#FFFFFF"
					cor = 1
					
					'if cd_situacao = 1 OR cd_situacao = "" then
					'xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE admissao <= '"&ano_sel&mes_sel&"' AND demissao >= '"&ano_sel&mes_sel&"' "&busca_nome&" "&busca_sexo&" OR admissao <= '"&ano_sel&mes_sel&"' AND dt_demissao is NULL "&busca_nome&" "&busca_sexo&" ORDER BY "&ordem_funcionarios&""
					
					'SELECT * FROM View_funcionario_contrato_lista 
					'WHERE admissao <= '"&ano_sel&mes_sel&"' AND demissao >= '"&ano_sel&mes_sel&"' OR admissao <= '"&ano_sel&mes_sel&"' AND dt_demissao is NULL  
					'ORDER BY "&ordem_funcionarios&""
					
					xsql = "up_funcionario_contrato_movimentacao @dt_data='"&ano_sel&mes_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
					'dbconn.execute(xsql)
					'xsql = "SELECT * FROM VIEW_funcionario_contrato_lista WHERE (admissao <= '"&zero(int(ano_sel))&mes_sel&"') AND dt_demissao >= '"&zero(int(mes_sel))&"/"&dia_final&"/"&ano_sel&"' OR (admissao <= '"&ano_sel&mes_sel&"') AND dt_demissao IS NULL ORDER BY "&ordem_funcionarios&""'nm_regime_trab,nm_nome, dt_admissao"
					
						Set rs_func = dbconn.execute(xsql)
							
							while not rs_func.EOF
								cd_matricula = rs_func("cd_matricula")
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
								cd_ctps = rs_func("cd_ctps")
								cd_ctps_serie = rs_func("cd_ctps_serie")
								
								nm_rg = rs_func("nm_rg")
								nm_cpf = rs_func("nm_cpf")
								
								dt_admissao = zero(day(rs_func("dt_admissao")))&"/"&zero(month(rs_func("dt_admissao")))&"/"&year(rs_func("dt_admissao"))
								dt_demissao = zero(day(rs_func("dt_demissao")))&"/"&zero(month(rs_func("dt_demissao")))&"/"&year(rs_func("dt_demissao"))
								
								admissao = rs_func("admissao")
								demissao = rs_func("demissao")
								
								cd_contrato = rs_func("cd_contrato")
								
								'tempo_casa = rs_func("tempo_casa")
								
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
								'xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 order by dt_atualizacao DESC"
								'xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 AND dt_atualizacao between '"&mes_sel&"/1/"&ano_sel&"' AND '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"'order by dt_atualizacao DESC"
								xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 AND dt_atualizacao <= '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"'order by dt_atualizacao DESC"
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
											<td><b>Mat.</b></td>
											<td><b>Funcionário</b></td>
											<%if campo_sex <> 0 Then%><td><b>Sexo</b></td> <%end if%>
											<%if campo_endereco <> 0 Then%><td><b>Endereço</b></td><%end if%>
											<%if campo_rg <> 0 Then%><td><b>RG</b></td><%end if%>
											<%if campo_cpf <> 0 Then%><td><b>CPF</b></td><%end if%>
											<%if campo_ctps <> 0 Then%><td><b>CTPS</b></td><%end if%>
											<%if campo_rgp <> 0 Then%><td><b>Reg. Profissional</b></td><%end if%>
											<%if campo_cargo <> 0 Then%><td><b>Cargo</b></td><%end if%>
											<%if campo_hora <> 0 Then%><td><b>Horário</b></td><%end if%>
											<%if campo_unidade <> 0 Then%><td><b>Unidade</b></td><%end if%>
											<%if campo_situacao <> 0 Then%><td><b>Situação</b></td><%end if%>
											<%if cd_situacao = 2 then%><td><b>Empresa</b></td><%end if%>
											<!--td><b>Admissão</b></td>
											<%'if cd_situacao = 2 then%>
											<td><b>demissão</b></td-->
											<%'end if%>
											<td><b>Admissão</b></td>
											<td><b>Recisão</b></td>
											<%if campo_tempo_casa <> 0 then%><td><b>Tempo de casa</b></td><%end if%>
											<%if campo_contratos <> 0 Then%><td colspan="2" align="center"><b>Todos os contratos</b></td><%end if%>
										</tr>
									<%conta_func = 0
									end if%>
								
								<%'*** Lista os funcionarios ***%>
								<!--tr-->
								<%if int(periodo_sel)=int(demissao) then
									cor_registro = "gray"
								else
									cor_registro = "black"
								end if%>
								<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px; color:<%=cor_registro%>;" bgcolor="<%=cor_linha%>">
									<td align="center" valign="top" style="color:gray;">
										<%if int(periodo_sel) = int(demissao) then
										else
											conta_func = conta_func + 1
											response.write(zero(conta_func))
										end if%>
										<%'=%></td>
									<td align="right" valign="top"><%=cd_matricula%></td>
									<td valign="top"><%=nm_funcionario%></td>
									<%if campo_sex <> 0 Then%><td valign="top"><%=nm_sexo%></td><%end if%>
									<%if campo_endereco <> 0 Then%><td valign="top"><%=nm_endereco%>, <%=nr_numero%> &nbsp; <br id="mok_print"><%=nm_bairro%> - <%=nm_cidade%>-<%=nm_estado%>&nbsp;</td><%end if%>
									<%if campo_rg <> 0 Then%><td valign="top"><%=nm_rg%></td><%end if%>
									<%if campo_cpf <> 0 Then%><td valign="top"><%=nm_cpf%></td><%end if%>
									<%if campo_ctps <> 0 Then%><td valign="top"><%=cd_ctps&"-"&cd_ctps_serie%></td><%end if%>
									<%if campo_rgp <> 0 Then%><td valign="top"><%=cd_numreg%></td><%end if%>
									<%if campo_cargo <> 0 Then%><td valign="top"><%=nm_cargo%></td><%end if%>
									<%if campo_hora <> 0 Then%><td valign="top"><%=hr_entrada%> - <%=hr_saida%></td><%end if%>
									<%if campo_unidade <> 0 Then%><td valign="top"><%=nm_sigla%></td><%end if%>
									<%if campo_situacao <> 0 Then%><td valign="top"><%=nm_status%></td><%end if%>
									<%'if cd_situacao = 2 then%><!--td valign="top"><%'=nm_regime_trab%></td--><%'end if%>
									
									<td valign="top">
										<%if int(periodo_sel)=int(admissao) then
											cor_admiss = "green;"
											entrada = entrada + 1
										end if%>
											<b style="color:<%=cor_admiss%>;"><%=dt_admissao%></b></td>
								<%'if cd_situacao = 2 then%>
									<td valign="top">
										<%if int(periodo_sel)=int(demissao) then
											cor_demiss = "red;"
											saida = saida + 1
										end if%>
											<b style="color:<%=cor_demiss%>;"><%=dt_demissao%></b></td>
								<%'end if%>
								<%if campo_tempo_casa <> 0 then%><td align="left"><b style="color:blue;"><%'=tempo_casa&" &nbsp"%></b> 
									
										
									<%anos = 365
									mes = 30'.5
									
										t_casa_a = int(tempo_casa / anos)
											dif_tempo_a = anos * t_casa_a
											tempo_casa = tempo_casa - dif_tempo_a
										t_casa_m = int(tempo_casa / mes)
											dif_tempo_m = mes * t_casa_m
											tempo_casa = tempo_casa - dif_tempo_m
										t_casa_d = int(tempo_casa)
									
									%>
										
										
										
									<%'=tempo_anos&espaco_data&tempo_meses%>
									<b><%=zero(t_casa_a)%></b>a <b><%=zero(t_casa_m)%></b>m <b><%=zero(t_casa_d)%></b>d &nbsp;<!--(<%=tempo_casa%>)-->
								</td><%end if%>
								
								
								<%if campo_contratos <> 0 Then%>
									<%xsql = "SELECT * FROM  VIEW_funcionario_contrato_lista WHERE (cd_funcionario = '"&cd_funcionario&"') order by dt_admissao"
										Set rs_contr = dbconn.execute(xsql)
											while not rs_contr.EOF
												nm_regime_trab = rs_contr("nm_regime_trab")
												dt_admissao = rs_contr("dt_admissao")
												dt_demissao = rs_contr("dt_demissao")%>
												<td valign="top"><%response.write("<b>"&ucase(left(nm_regime_trab,5))&":</b>"&zero(day(dt_admissao))&"/"&zero(month(dt_admissao))&"/"&year(dt_admissao)&" - "&zero(day(dt_demissao))&"/"&zero(month(dt_demissao))&"/"&year(dt_demissao)&"<br>")%><!--br-->
													<%if dt_demissao <> "" Then
													else
														dt_demissao = now
													end if%>
												<%'=datediff("m",dt_admissao,dt_demissao)%></td>
												
											<%rs_contr.movenext
											wend%>									
								<%end if%>
								</tr>
								<%if cabeca_ativo <> cd_contrato then%>
								<tr>
									<td></td>
								</tr>						
								<%end if%>
								<%'if cd_situacao = 2 and pulo_linha_inativos = cd_funcionario then%>
								
								<%cabeca_ativo = cd_contrato
								if cor > 0 then
									cor_linha = "#d7d7d7"
									cor = 0
								else
									cor_linha = "#FFFFFF"
									cor = 1
								end if
								
							
							espaco_data = ""
							conta_func_total = conta_func_total + 1	
							cor_admiss = "black"
							cor_demiss = "black"
							
							rs_func.movenext
							wend%>	
								
					<%'cabeca = ""
					
					lista_empresa = ""
'*******************************************************%>
		<tr>
			<td colspan="100%"><img src="../../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td>				
		<tr>
			<td></td>
			<td colspan="3">Total de colaboradores <b><%=conta_func_total-saida%></b></td>
		</tr>
		
		<%'*** conta os funcionários de cada contrato ***
				'*** seleciona o mês anterior ***
					if mes_sel = 1 then
						mes_anterior = 12
						ano_anterior = ano_sel - 1
					else
						mes_anterior = mes_sel - 1
						ano_anterior = ano_sel
					end if
					
				data_sel = ano_anterior&zero(mes_anterior)
		
		total_funcionarios = 0
		
		'*** verifica quais contratos estavam ativos ***
		xsql = "SELECT DISTINCT cd_contrato, nm_regime_trab FROM VIEW_funcionario_contrato_lista WHERE admissao <= '"&ano_sel&mes_sel&"' AND demissao > '"&ano_sel&mes_sel&"' OR admissao <= '"&ano_sel&mes_sel&"' AND dt_demissao IS NULL GROUP BY cd_contrato,nm_regime_trab"
			Set rs_contr = dbconn.execute(xsql)
				while not rs_contr.EOF
				cd_contrato = rs_contr("cd_contrato")
				nm_regime_trab = rs_contr("nm_regime_trab")%>
		<tr><td colspan="3"><hr></td></tr>
		<tr>
			<td style="background-color:gray; color:white;" colspan="3">CONTRATO: <b><%=nm_regime_trab%></b></td>
		</tr>
			<%'*** conta os funcionários de cada contrato ***
				'*** seleciona o mês anterior ***
					if mes_sel = 1 then
						mes_anterior = 12
						ano_anterior = ano_sel - 1
					else
						mes_anterior = mes_sel - 1
						ano_anterior = ano_sel
					end if
					
				data_sel = ano_anterior&zero(mes_anterior)
				
				
					xsql = "SELECT COUNT(admissao) AS conta, cd_contrato FROM VIEW_funcionario_contrato_lista WHERE (admissao <= '"&data_sel&"') AND (demissao > '"&data_sel&"') AND (cd_contrato = "&cd_contrato&") OR (admissao <= '"&data_sel&"') AND (dt_demissao IS NULL) AND (cd_contrato = "&cd_contrato&") GROUP BY cd_contrato"
					Set rs_conta = dbconn.execute(xsql)
						while not rs_conta.EOF
							conta = rs_conta("conta")
						rs_conta.movenext
						wend
						
							xsql = "SELECT COUNT(cd_contrato) AS conta_admissao FROM VIEW_funcionario_contrato_lista WHERE (admissao = '"&ano_sel&mes_sel&"') AND (cd_contrato = "&cd_contrato&")"
							Set rs_admiss = dbconn.execute(xsql)
								if not rs_admiss.EOF Then
									total_admissao = rs_admiss("conta_admissao")
								end if
							
							xsql = "SELECT COUNT(cd_contrato) AS conta_recisao FROM VIEW_funcionario_contrato_lista WHERE (demissao  = '"&ano_sel&mes_sel&"') AND (cd_contrato = "&cd_contrato&")"
							Set rs_recisao = dbconn.execute(xsql)
								if not rs_recisao.EOF Then
									total_recisao = rs_recisao("conta_recisao")
								end if
							%> 
			<tr><td colspan="3">Saldo inicio Mês: <b><%=int(conta)%></b> <%'="("&data_sel&")"%></td></tr>
			<tr><td colspan="3">Admissões: <b style="color:green;"><%=total_admissao%></b></td></tr>
			<tr><td colspan="3">Recisões: <b style="color:red;"><%=total_recisao%></b></td></tr>
			<tr><td colspan="3">TOTAL: <b><%total_geral=(conta+total_admissao)-total_recisao%> <%=total_geral%></b></td></tr>			
			<!--tr><td><%=data_sel%></td></tr-->
		<%total_funcionarios = total_funcionarios + total_geral
		
		conta = 0
		rs_contr.movenext
		wend%>
		<tr><td colspan="3"><hr></td></tr>
		<tr><td colspan="3"><b>HEAD COUNT: <%=total_funcionarios%></b> <%if total_funcionarios > 1 then%>colaboradores<%else%>Colaborador<%end if%> </td></tr>
		<tr>
			<td><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
			<td><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
			<td><img src="../../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
			<td><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>	
			<%if campo_endereco <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="240" height="1" border="0"></td><%end if%>
			
			<%if campo_rg <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="90" height="1" border="0"></td><%end if%>
			<%if campo_cpf <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td><%end if%>
			
			<%if campo_rgp <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"></td><%end if%>
			<%if campo_cargo <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"></td><%end if%>
			<%if campo_hora <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td><%end if%>
			<%'if campo_unidade <> 0 Then%><td><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td><%'end if%>	
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
