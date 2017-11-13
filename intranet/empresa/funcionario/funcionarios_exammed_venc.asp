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
	'ordem_funcionarios = "dt_rgproval,dt_admissao,nm_nome"
	'ordem_funcionarios = "validade, dt_exame_medico desc,nm_nome desc,dt_admissao desc"
	ordem_funcionarios = "nm_nome,validade, dt_exame_medico desc,dt_admissao desc"
	
%>


<head>
	<title>Listagem de funcionários</title>
</head>

<body><br id="no_print">

<table align="center" width="40">
	<tr>
		<td colspan="4"><a href="empresa.asp?tipo=resumo_venc">Gestão de enfermagem >></a></td>
		<form action="../../empresa.asp" name="filtro" id="filtro">
	<input type="hidden" name="tipo" value="exammed_venc">
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
	</tr>
	<tr><td>&nbsp;</td></tr>
		<%cor_linha = "#FFFFFF"
		cor = 1
		cabeca_exam = 0
		'MONTH(dt_exame_medico) + '-' + DAY(dt_exame_medico) + '-' + YEAR(dt_exame_medico)
		if cd_situacao = 1 OR cd_situacao = "" then
			'xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_exame_medico) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND (DATEDIFF(m, GETDATE(),dt_exame_medico) <= '1') "&busca_nome&" "&busca_sexo&" AND dt_demissao is null OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL AND (DATEDIFF(m, GETDATE(),dt_exame_medico) <= '1')"&busca_nome&" "&busca_sexo&"  AND dt_demissao is null ORDER BY "&ordem_funcionarios&""
			'xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_exame_medico) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND (DATEDIFF(m, GETDATE(),dt_exame_medico) <= '1') "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL AND (DATEDIFF(m, GETDATE(),dt_exame_medico) <= '1') "&busca_nome&" "&busca_sexo&"  ORDER BY "&ordem_funcionarios&""
			xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_exame_medico) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL "&busca_nome&" "&busca_sexo&"  ORDER BY "&ordem_funcionarios&""
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
						
						'dt_rgproinscr = rs_func("dt_rgproinscr")
						dt_exame_medico = rs_func("dt_exame_medico")
						dt_exame_ultimo = zero(day(dt_exame_medico))&"/"&zero(month(dt_exame_medico))&"/"&year(dt_exame_medico)
							
						dt_exame_medico_val = zero(day(dt_exame_medico))&"/"&zero(month(dt_exame_medico))&"/"&year(dt_exame_medico)+1
						validade = rs_func("validade")+12
						val = validade
						
						if validade <> "" then
							val_exam = validade
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
							val_exam = NULL
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
							<tr><td colspan="100%" align="center" style="background-color:gray; color:white; font-size:14px;"><b>VENCIMENTO DOS EXAMES MÉDICOS</b></td></tr>
						<%cabeca_ativo = 1
						end if%>
							
							<%if cabeca_exam = 0 AND isnull(val_exam) Then%>
								<tr><td align="center" colspan="6" style="color:white;" bgcolor="#808080"><span style="color:#ff0000;font-size:12px;"><<<</span> Exames não preenchidos <span style="color:#ff0000;">>>></span></td></tr>
							<%cabeca_exam = 1
							cabeca_exam1 = 0
							cor_val = "#000000"
							conta_vazios = 0
							
							elseif cabeca_exam = 1 AND val_exam > -120 Then%>
								<tr><td align="center" colspan="6" style="color:white;" bgcolor="#808080"><span style="color:#ff0000;font-size:12px;"><<<</span> Exames já vencidos <span style="color:#ff0000;">>>></span></td></tr>
							<%cabeca_exam = 2
							cabeca_exam1 = 0
							cor_val = "#ff0000"
							conta_vencidos = 0
							
							elseif cabeca_exam = 2 AND val_exam > 0 Then%>
								<tr><td align="center" colspan="6" style="color:white;" bgcolor="#808080"><span style="color:#fa9814;font-size:12px;"><<<</span> Exames vencendo <span style="color:#fa9814;">>>></span></td></tr>
							<%cabeca_exam = 3
							cabeca_exam1 = 0
							cor_val = "#c18509"
							conta_vencendo = 0
							
							elseif cabeca_exam = 3 AND val_exam > 3 Then%>
								<tr><td align="center" colspan="6" style="color:white;" bgcolor="#808080"><span style="color:#00ff40;font-size:12px;"><<<</span> Exames OK <span style="color:#00ff40;">>>></span></td></tr>
							<%cabeca_exam = 4
							cabeca_exam1 = 0
							cor_val = "#000000"
							conta_ok = 0
							end if%>
							
							<%if cabeca_exam1 = 0 then%><tr bgcolor=#c0c0c0>
								<td>&nbsp;</td>
								<td align="center"><b>Mat.</b></td>
								<td align="center"><b>Funcionário</b></td>
								<td align="center"><b>Último Exame<br>Periódico</b></td>
								<td align="center"><b>Validade</b></td>											
								<td align="center"><b>Tempo</b></td>
							</tr>
							<%cabeca_exam1 = 1
							end if%>
					<%if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
					conta_func = conta_func + 1	%>
					<%'*** Lista os funcionarios ***%>
					<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');" onclick="javascript:window.open('empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>#exammed','Cadastro_<%=rs_func("cd_funcionario")%>','width=760,height=700,scrollbars=1');" style="height:20px;" bgcolor="<%=cor_linha%>">
					<!-- tr onmouseover="mOvr(this,'#CFC8FF');" onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px;" bgcolor="<%=cor_linha%>"-->
						<td align="center" valign="top" style="color:gray;"><%=zero(conta_func)%></td>
						<td align="right" valign="top"><%=zero(cd_matricula)%></td>
						<td valign="top"><%=nm_funcionario%></td>
						<td valign="top" align="center"><%=dt_exame_ultimo%></td>
						<td valign="top" align="center"><%=dt_exame_medico_val%></td>						
						<td valign="top" align="center" style="color:<%=cor_val%>;"><%=validade%></td>								
					<%if cd_situacao = 2 then%>
						<td valign="top"><%=nm_regime_trab%>***</td>
					<%end if%>
						
					<%if cd_situacao = 2 then%>
						<td> valign="top"<%=dt_demissao%></td>
					<%end if%>
					</tr>
					
					<%'cabeca_ativo = cd_contrato
						
						if conta_vencidos = "" AND conta_vencendo = "" AND conta_ok = "" then
							conta_vazios = conta_vazios + 1
						elseif conta_vazios <> "" AND conta_vencendo = "" AND conta_ok = "" then
							conta_vencidos = conta_vencidos + 1
						elseif conta_vazios <> "" AND conta_vencidos <> "" AND conta_ok = "" then
							conta_vencendo = conta_vencendo + 1
						elseif conta_vazios <> "" AND conta_vencidos <> "" AND conta_vencendo <> "" then
							conta_ok = conta_ok + 1
						end if
					cabeca_ativo = cd_contrato
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
	<tr></tr>
	<tr>
		<td><img src="../../imagens/blackdot.gif" alt="" width="25" height="1" border="0"></td>
		<td><img src="../../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td>
		<td><img src="../../imagens/blackdot.gif" alt="" width="220" height="1" border="0"></td>
		<td><img src="../../imagens/blackdot.gif" alt="" width="90" height="1" border="0"></td>	
		<td><img src="../../imagens/blackdot.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../../imagens/blackdot.gif" alt="" width="70" height="1" border="0"></td>
	</tr>
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
