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
	campo_rg = Instr(1,campos,"rg_",0)
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
	if sexo > "0" Then
		busca_sexo = "AND cd_sexo = '"&sexo&"'"
	end if 
strcd_unidade = request("cd_unidade")
	if strcd_unidade = "" Then	strcd_unidade = 0
'*** ORDEM ***
	ordem_funcionarios = "dt_rgproval,dt_admissao,nm_nome"
	
%>


<head>
	<title>Listagem de funcionários</title>
</head>

<body><br id="no_print">

<table align="center" width="40">
	<tr><td colspan="4"><a href="empresa.asp?tipo=resumo_venc">Gestão de enfermagem >></a></td></tr>
	<form action="../../empresa.asp" name="filtro" id="filtro">
	<input type="hidden" name="tipo" value="coren_venc">
	<tr>
		<td colspan="6"><%strsql ="TBL_unidades where cd_hospital > 0 and cd_status = 1"
			  	Set rs_uni = dbconn.execute(strsql)%>
				<select name="cd_unidade" class="inputs"> 
					<option value="">Todas</option>
					<%Do While Not rs_uni.eof%>
					<option value="<%=rs_uni("cd_codigo")%>" <%if int(strcd_unidade) = rs_uni("cd_codigo") then response.write("SELECTED")%>><%=rs_uni("nm_Unidade")%></option>
					<%rs_uni.movenext
					loop
					rs_uni.close
					Set rs_uni = nothing%>
				</select>
				<input type="submit" name="filtra" value="filtrar">
		</td>
	</tr>
	</form>
	<tr><td>&nbsp;</td></tr>
		<%cor_linha = "#FFFFFF"
		cor = 1
		cabeca_coren = 0
		
		if cd_situacao = 1 OR cd_situacao = "" then
			'xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_rgproval) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND (DATEDIFF(m, GETDATE(),dt_rgproval) <= '1') "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL AND (DATEDIFF(m, GETDATE(),dt_rgproval) <= '1')"&busca_nome&" "&busca_sexo&" ORDER BY "&ordem_funcionarios&""
			xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_rgproval) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND cd_conspro <> 1 "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL AND cd_conspro <> 1 "&busca_nome&" "&busca_sexo&" ORDER BY "&ordem_funcionarios&""
		elseif cd_situacao = 2 then
			xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_demissao is not NULL "&busca_nome&" "&busca_sexo&" ORDER BY "&ordem_funcionarios&""
		end if
			Set rs_func = dbconn.execute(xsql)
				
				while not rs_func.EOF
					'conta_func = conta_func + 1	
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
						
						dt_rgproinscr = rs_func("dt_rgproinscr")
						dt_rgproval = rs_func("dt_rgproval")
						obs_rgpro = rs_func("obs_rgpro")
						obs_rgpro = rs_func("obs_rgpro")
						
						validade = rs_func("validade")
						
						if validade <> "" then
							val_coren = validade
							if validade < (-12) then
								if validade < 13 then
									tempo_ano = "ano"
								else
									tempo_ano = "anos"
								end if 
									validade_mes = mid((validade/12),instr(validade/12,",")+1,left(len(validade/12),1))
									validade = mid(validade/12,1,instr(validade/12,","))&validade_mes&" "&tempo_ano
									validade = replace(validade,"-","")
							else
								validade = replace(validade,"-","")
								if validade = 1 then
									tempo_mes = "mes"
								else
									tempo_mes = "meses"
								end if 
								
								validade = replace(validade,"-","")&" "&tempo_mes
								'validade = validade&"meses"
							end if
						else
							val_coren = NULL
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
							cd_unidade = rs_unid("cd_unidade")
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
						'if cabeca_ativo <> cd_contrato then
						if cabeca_ativo = 0 Then%>
							<tr><td colspan="100%" align="center" style="background-color:gray; color:white; font-size:14px;"><b>VENCIMENTO DOS CORENS</b></td></tr>
						<%cabeca_ativo = 1
						end if%>
						
							<%if cabeca_coren = 0 AND isnull(val_coren) Then%>
								<tr><td align="center" colspan="100%" style="color:white;font-size:13px;" bgcolor="#808080">SEM DATA DE VENCIMENTO</td></tr>
							<%cabeca_coren = 1
							cabeca_coren1 = 1
							conta_vazios = 0
							
							elseif cabeca_coren = 1 AND val_coren > -120 Then%>
								<tr><td align="center" colspan="100%" style="color:white;font-size:13px;" bgcolor="#808080">CORENS VENCIDOS</td></tr>
							<%cabeca_coren = 2
							cabeca_coren1 = 1
							conta_vencidos = 0
							
							elseif cabeca_coren = 2 AND val_coren >= 1 then%>
								<tr><td align="center" colspan="100%" style="color:white;" bgcolor="#808080">CORENS VENCENDO</td></tr>
							<%cabeca_coren = 3
							cabeca_coren1 = 1
							conta_vencendo = 0
							
							elseif cabeca_coren = 3 AND val_coren >= 3 then%>
								<tr><td align="center" colspan="100%" style="color:white;" bgcolor="#808080">CORENS REGULARES</td></tr>
							<%cabeca_coren = 4
							cabeca_coren1 = 1
							conta_ok = 0
							end if%>
							
						<%if cabeca_coren1 = 1 then%>
							<tr bgcolor=#c0c0c0>
								<td><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
								<td align="center"><b>Mat.</b><br><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
								<td align="center"><b>Funcionário</b><br><img src="../../imagens/px.gif" alt="" width="190" height="1" border="0"></td>
								<td align="center"><b>CPF</b><br><img src="../../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
								<td align="center"><b>Coren</b><br><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
								<td align="center"><b>Inscrição</b><br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
								<td align="center"><b>Validade</b><br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>											
								<td align="center"><b>Tempo</b><br><img src="../../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
								<td align="center"><b>Cargo</b><br><img src="../../imagens/px.gif" alt="" width="115" height="1" border="0"></td>
								<td align="center"><b>Unidade</b><br><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
								<td align="center"><b>Admissão</b><br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
								<td align="center"><b>Observação</b><br><img src="../../imagens/px.gif" alt="" width="190" height="1" border="0"></td>
							</tr>
						<%cabeca_coren1 = 2
						
						end if%>
				<%if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
					conta_func = conta_func + 1	%>
					<%'*** Lista os funcionarios ***%>
					<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');" onclick="javascript:window.open('empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>#coren','Cadastro_<%=rs_func("cd_funcionario")%>','width=760,height=700,scrollbars=1');" style="height:20px;" bgcolor="<%=cor_linha%>">
					<!-- tr onmouseover="mOvr(this,'#CFC8FF');" onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px;" bgcolor="<%=cor_linha%>"-->
						<td align="center" valign="top" style="color:gray;"><%=zero(conta_func)%></td>
						<td align="right" valign="top"><%=cd_matricula%></td>
						<td valign="top"><%=nm_funcionario%></td>
						<td align="right" valign="top" ><%=nm_cpf%></td>
						<td align="right" valign="top"><%=cd_numreg%></td>
						<td valign="top"><%=zero(day(dt_rgproinscr))&"/"&zero(month(dt_rgproinscr))&"/"&year(dt_rgproinscr)%></td>
						<td valign="top"><%=dt_rgproval%></td>						
						<td valign="top" <%if val_coren <= 0 then response.write("style='color:red;'")%>><%=validade%><%'=" - "&cabeca_coren&"/"&val_coren%></td>
						<td valign="top"><%=nm_cargo%></td>
						<td valign="top"><%=nm_sigla%></td>					
					<%if cd_situacao = 2 then%>
						<td valign="top"><%=nm_regime_trab%>***</td>
					<%end if%><td valign="top"><%=dt_admissao%></td>
					<%if cd_situacao = 2 then%>
						<td> valign="top"<%=dt_demissao%></td>
					<%end if%>
						<td valign="middle" align="center">&nbsp;<%=obs_rgpro%></td>
					</tr>
					
					<%'cabeca_coren = cabeca_coren + 1
					cabeca_ativo = cd_contrato
					
					if conta_vencidos = "" AND conta_vencendo = "" AND conta_ok = "" then
						conta_vazios = conta_vazios + 1
					elseif conta_vazios <> "" AND conta_vencendo = "" AND conta_ok = "" then
						conta_vencidos = conta_vencidos + 1
					elseif conta_vazios <> "" AND conta_vencidos <> "" AND conta_ok = "" then
						conta_vencendo = conta_vencendo + 1
					elseif conta_vazios <> "" AND conta_vencidos <> "" AND conta_vencendo <> "" then
						conta_ok = conta_ok + 1
					end if
					
					if cor > 0 then
						cor_linha = "#d7d7d7"
						cor = 0
					else
						cor_linha = "#FFFFFF"
						cor = 1
					end if
				end if
				rs_func.movenext
				wend%>				
		<%'cabeca = ""
		
		lista_empresa = ""
'*******************************************************%>
	<tr><td colspan="12"><hr></td></tr>
	<tr>
		<td colspan="2">&nbsp;</td>
		<td>Vazios</td>
		<td align="right"><%=conta_vazios%></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
		<td>Vencidos</td>
		<td align="right"><%=conta_vencidos%></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
		<td>Vencendo</td>
		<td align="right"><%=conta_vencendo%></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
		<td>Ok</td>
		<td align="right"><%=conta_ok%></td>
	</tr>
	<tr>
		<td><%'=lista_funcionario%><%'=remanescentes%></td>
	</tr>
		<td></td>
	</tr>
</table>


</body>
</html>
