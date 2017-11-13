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
%>
<SCRIPT LANGUAGE="javascript">

function hepatb(){
		document.location= "empresa.asp?tipo=hepatb_vacina";	
	}
	
function dupla(){
		document.location= "empresa.asp?tipo=dupla_venc";	
	}
	
function scr(){
		document.location= "empresa.asp?tipo=scr_vacina";	
	}
	
function exame_med(){
		document.location= "empresa.asp?tipo=exammed_venc";	
	}
	
function coren(){
		document.location= "empresa.asp?tipo=coren_venc";
	}

</script>

<head>
	<title>Gestão de Enfermagem</title>
</head>

<body><br id="no_print">

<table>
	<tr>
		<td align="center" colspan="3" bgcolor="#808080" style="color:white;">GESTÃO DE ENFERMAGEM</td>
	</tr>
	<tr>
		<td>
		<%'****************************
		'***	Vacina - Hepatite B	***
		'******************************
		ordem_funcionarios = "nr_hepatite_b, nm_nome desc,dt_admissao desc"%>
			<table align="center" width="40" onclick="hepatb();">
					<%cor_linha = "#FFFFFF"
					cor = 1
					cabeca_exam = 0
					if cd_situacao = 1 OR cd_situacao = "" then
						xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL "&busca_nome&" "&busca_sexo&"  ORDER BY "&ordem_funcionarios&""
					end if
						Set rs_func = dbconn.execute(xsql)
							
							while not rs_func.EOF
								conta_func = conta_func + 1	
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
									nr_hepatite_b = rs_func("nr_hepatite_b")
										if nr_hepatite_b = 1 then
											x_vacina1 = "X"
										elseif nr_hepatite_b = 2 then
											x_vacina1 = "X"
											x_vacina2 = "X"
										elseif nr_hepatite_b = 3 then
											x_vacina1 = "X"
											x_vacina2 = "X"
											x_vacina3 = "X"
										end if%>
										
										<%'if cabeca_exam = 0 AND isnull(nr_hepatite_b) Then
										if cabeca_exam = 0 AND nr_hepatite_b = 0 Then
										cabeca_exam = 1
										cabeca_exam1 = 0
										cor_val = "#000000"
										conta_0dose = 0
										
										elseif cabeca_exam = 1 AND nr_hepatite_b = 1 Then
										cabeca_exam = 2
										cabeca_exam1 = 0
										cor_val = "#ff0000"
										conta_1dose = 0
										
										elseif cabeca_exam = 2 AND nr_hepatite_b = 2 Then
										cabeca_exam = 3
										cabeca_exam1 = 0
										cor_val = "#c18509"
										conta_2dose = 0
										
										elseif cabeca_exam = 3 AND nr_hepatite_b =  3 Then
										cabeca_exam = 4
										cabeca_exam1 = 0
										cor_val = "#000000"
										conta_3dose = 0
										end if%>
										
								<%'cabeca_ativo = cd_contrato
									
									if conta_1dose = "" AND conta_2dose = "" AND conta_3dose = "" then
										conta_0dose = conta_0dose + 1
									elseif conta_0dose <> "" AND conta_2dose = "" AND conta_3dose = "" then
										conta_1dose = conta_1dose + 1
									elseif conta_0dose <> "" AND conta_1dose <> "" AND conta_3dose = "" then
										conta_2dose = conta_2dose + 1
									elseif conta_0dose <> "" AND conta_1dose <> "" AND conta_2dose <> "" then
										conta_3dose = conta_3dose + 1
									end if
									
								x_vacina1 = ""
								x_vacina2 = ""
								x_vacina3 = ""
							rs_func.movenext
							wend%>				
					<%'cabeca = ""
						
					
					lista_empresa = ""
			'*******************************************************%>
				<tr>
					<td align="center" colspan="2" bgcolor="silver"><b>VACINA - Hepatite B</b></td>
				</tr>
				<tr>
					<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>		
				</tr>
				
				<tr>
					<td>Nenhuma</td>
					<td align="right"><%=conta_0dose%></td>
				</tr>
				<tr>
					<td>1ª dose</td>
					<td align="right"><%=conta_1dose%></td>
				</tr>
				<tr>
					<td>2ª dose</td>
					<td align="right"><%=conta_2dose%></td>
				</tr>
				<tr>
					<td>3ª dose</td>
					<td align="right"><%=conta_3dose%></td>
				</tr>
				<tr style="background-color:#e0e0e0;">
					<td><b>Total</b></td>
					<td align="right"><b><%=conta_0dose+conta_1dose+conta_2dose+conta_3dose%></b></td>
				</tr>
				<%conta_0dose = ""
				conta_1dose = ""
				conta_2dose = ""
				conta_3dose = ""%>
			</table>
		</td>
		<td>
			<%'********************************
			'***	Vacinas - Dupla Adulto	***
			'**********************************
			ordem_funcionarios = "validade,nr_dupla_adulto, dt_dupla_adulto_validade desc,nm_nome desc,dt_admissao desc"%>
			<table align="center" width="40" onclick="dupla();">
				<%cor_linha = "#FFFFFF"
				cor = 1
				cabeca_exam = 0
				'MONTH(dt_exame_medico) + '-' + DAY(dt_exame_medico) + '-' + YEAR(dt_exame_medico)
				if cd_situacao = 1 OR cd_situacao = "" then
					'xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_exame_medico) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND (DATEDIFF(m, GETDATE(),dt_exame_medico) <= '1') "&busca_nome&" "&busca_sexo&" AND dt_demissao is null OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL AND (DATEDIFF(m, GETDATE(),dt_exame_medico) <= '1')"&busca_nome&" "&busca_sexo&"  AND dt_demissao is null ORDER BY "&ordem_funcionarios&""
					'xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_exame_medico) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND (DATEDIFF(m, GETDATE(),dt_exame_medico) <= '1') "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL AND (DATEDIFF(m, GETDATE(),dt_exame_medico) <= '1') "&busca_nome&" "&busca_sexo&"  ORDER BY "&ordem_funcionarios&""
					xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_dupla_adulto_validade) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL "&busca_nome&" "&busca_sexo&"  ORDER BY "&ordem_funcionarios&""
				end if
					Set rs_func = dbconn.execute(xsql)
						
						while not rs_func.EOF
							conta_func = conta_func + 1	
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
								nr_dupla_adulto = rs_func("nr_dupla_adulto")
									if nr_dupla_adulto = 1 then
										x_vacina1 = "X"
									elseif nr_dupla_adulto = 2 then
										x_vacina1 = "X"
										x_vacina2 = "X"
									elseif nr_dupla_adulto = 3 then
										x_vacina1 = "X"
										x_vacina2 = "X"
										x_vacina3 = "X"
									end if
								dt_dupla_adulto_validade = rs_func("dt_dupla_adulto_validade")
									
								dt_dupla_adulto_val = zero(day(dt_dupla_adulto_validade))&"/"&zero(month(dt_dupla_adulto_validade))&"/"&year(dt_dupla_adulto_validade)
								validade = rs_func("validade")+12
								val = validade
								
								if validade <> "" then
									val_vac = validade
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
									val_vac = NULL
								end if%>
									
									<%if cabeca_exam = 0 AND isnull(val_vac) Then
									cabeca_exam = 1
									cabeca_exam1 = 0
									cor_val = "#000000"
									conta_vazios = 0
									
									elseif cabeca_exam = 1 AND val_vac > -120 Then
									cabeca_exam = 2
									cabeca_exam1 = 0
									cor_val = "#ff0000"
									conta_vencidos = 0
									
									elseif cabeca_exam = 2 AND val_vac > 0 Then
									cabeca_exam = 3
									cabeca_exam1 = 0
									cor_val = "#c18509"
									conta_vencendo = 0
									
									elseif cabeca_exam = 3 AND val_vac > 3 Then
									cabeca_exam = 4
									cabeca_exam1 = 0
									cor_val = "#000000"
									conta_ok = 0
									end if%>								
									
							
							<%'*** Lista os funcionarios ***%>
							
							
							<%if conta_vencidos = "" AND conta_vencendo = "" AND conta_ok = "" then
								conta_vazios = conta_vazios + 1
							elseif conta_vazios <> "" AND conta_vencendo = "" AND conta_ok = "" then
								conta_vencidos = conta_vencidos + 1
							elseif conta_vazios <> "" AND conta_vencidos <> "" AND conta_ok = "" then
								conta_vencendo = conta_vencendo + 1
							elseif conta_vazios <> "" AND conta_vencidos <> "" AND conta_vencendo <> "" then
								conta_ok = conta_ok + 1
							end if
								
							x_vacina1 = ""
							x_vacina2 = ""
							x_vacina3 = ""
						rs_func.movenext
						wend%>				
				
				<%lista_empresa = ""
			'*******************************************************%>
				<tr>
				<td align="center" colspan="2" bgcolor="silver"><b>VACINA - Dupla adulto</b></td>
			</tr>
			<tr>
				<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
				<td><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
			</tr>
				<tr>
					<td>Vazios</td>
					<td align="right"><%=conta_vazios%></td>
				</tr>
				<tr style="color:red">
					<td>Vencidos</td>
					<td align="right"><%=conta_vencidos%></td>
				</tr>
				<tr>
					<td>Vencendo</td>
					<td align="right"><%=conta_vencendo%></td>
				</tr>
				<tr>
					<td>Ok</td>
					<td align="right"><%=conta_ok%></td>
				</tr>
				<tr style="background-color:#e0e0e0;">
					<td><b>Total</b></td>
					<td align="right"><b><%=conta_vazios+conta_vencidos+conta_vencendo+conta_ok%></b></td>
				</tr>
			</table>
			<%conta_vazios = ""
			conta_vencidos = ""
			conta_vencendo = ""
			conta_ok = ""%>
		</td>
		<td>
			<%'************************
			'***	VACINA - SCR	***
			'**************************
			'*** ORDEM ***
			ordem_funcionarios = "nr_scr, nm_nome desc,dt_admissao desc"%>
			<table align="center" width="40" onclick="scr();">
	
				<%cor_linha = "#FFFFFF"
				cor = 1
				cabeca_scr = 0
				'MONTH(dt_exame_medico) + '-' + DAY(dt_exame_medico) + '-' + YEAR(dt_exame_medico)
				if cd_situacao = 1 OR cd_situacao = "" then
					xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL "&busca_nome&" "&busca_sexo&"  ORDER BY "&ordem_funcionarios&""
				end if
					Set rs_func = dbconn.execute(xsql)
						
						while not rs_func.EOF
							conta_func = conta_func + 1	
							cd_matricula = rs_func("cd_matricula")
							cd_funcionario = rs_func("cd_funcionario")
							nm_funcionario = rs_func("nm_nome")
							
							
							xsql_regprof = "Select * From TBL_funcionario Where cd_codigo = "&cd_funcionario&""
							Set rs_regprof = dbconn.execute(xsql_regprof)
								if not rs_regprof.EOF then
									cd_numreg = rs_regprof("cd_numreg")
								end if
								
								nr_scr = rs_func("nr_scr")
									if nr_scr = 0 then
										scr_nao = "X"
									elseif nr_scr = 1 then
										scr_sim = "X"
									end if%>
									
									<%'if cabeca_exam = 0 AND isnull(nr_hepatite_b) Then
									if cabeca_scr = 0 AND nr_scr = 0 Then
									cabeca_scr = 1
									cabeca_scr1 = 0
									cor_val = "#ff0000"
									conta_nao = 0
									
									elseif cabeca_scr = 1 AND nr_scr = 1 Then
									cabeca_scr = 2
									cabeca_scr1 = 0
									cor_val = "#c18509"
									conta_sim = 0
									end if%>
							
							<%	if conta_sim = "" then
									conta_nao = conta_nao + 1
								elseif conta_nao <> "" then
									conta_sim = conta_sim + 1
								end if
								
							scr_nao = ""
							scr_sim = ""
						rs_func.movenext
						wend%>				
				<%'cabeca = ""
				
				lista_empresa = ""
			'*******************************************************%>
				
				<tr>
					<td align="center" colspan="2" bgcolor="silver"><b>VACINA - SCR</b></td>
				</tr>
				<tr>
					<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
				</tr>
				<tr style="color:red">
					<td>NÃO</td>
					<td align="right"><%=conta_nao%></td>
				</tr>
				<tr>
					<td>SIM</td>
					<td align="right"><%=conta_sim%></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td align="right">&nbsp;</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td align="right">&nbsp;</td>
				</tr>
				<tr style="background-color:#e0e0e0;">
					<td><b>Total</b></td>
					<td align="right"><b><%=conta_nao+conta_sim%></b></td>
				</tr>
			</table>
		</td>
	</tr>
	
	
	
	
	
	
	<tr><td>&nbsp;</td></tr>
	<tr>
		<td>
			<table align="center" width="40" onclick="exame_med();">
	
				<%'************************
				'***	EXAMES MÉDICOS	***
				'**************************
				'*** ORDEM ***
				ordem_funcionarios = "validade, dt_exame_medico desc,nm_nome desc,dt_admissao desc"
				
				cor_linha = "#FFFFFF"
				cor = 1
				cabeca_exam = 0
				if cd_situacao = 1 OR cd_situacao = "" then
					xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_exame_medico) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL "&busca_nome&" "&busca_sexo&"  ORDER BY "&ordem_funcionarios&""
				end if
					Set rs_func = dbconn.execute(xsql)
						
						while not rs_func.EOF
							conta_func = conta_func + 1	
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
							
							'*** Mostra o cabeçalho ****%>
								
									
									<%if cabeca_exam = 0 AND isnull(val_exam) Then
									cabeca_exam = 1
									cabeca_exam1 = 0
									conta_vazios = 0
									
									elseif cabeca_exam = 1 AND val_exam > -120 Then
									cabeca_exam = 2
									cabeca_exam1 = 0
									conta_vencidos = 0
									
									elseif cabeca_exam = 2 AND val_exam > 0 Then
									cabeca_exam = 3
									cabeca_exam1 = 0
									conta_vencendo = 0
									
									elseif cabeca_exam = 3 AND val_exam > 3 Then
									cabeca_exam = 4
									cabeca_exam1 = 0
									conta_ok = 0
									end if
									
									if cabeca_exam1 = 0 then
									cabeca_exam1 = 1
									end if
							
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
								
							
							
						rs_func.movenext
						wend%>				
				<%'cabeca = ""
				
				lista_empresa = ""
		'*******************************************************%>
			<tr>
				<td align="center" colspan="2" bgcolor="silver"><b>EXAME MÉDICO</b></td>
			</tr>
			<tr>
				<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
				<td><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
			</tr>
			<tr>
				<td>Ñ Preenchidos</td>
				<td align="right"><%=conta_vazios%></td>
			</tr>
			<tr style="color:red">
				<td>Vencidos</td>
				<td align="right"><%=conta_vencidos%></td>
			</tr>
			<tr>
				<td>Vencendo</td>
				<td align="right"><%=conta_vencendo%></td>
			</tr>
			<tr>
				<td>Ok</td>
				<td align="right"><%=conta_ok%></td>
			</tr>
			<tr style="background-color:#e0e0e0;">
				<td><b>Total</b></td>
				<td align="right"><b><%=conta_vazios+conta_vencidos+conta_vencendo+conta_ok%></b></td>
			</tr>
			<%conta_vazios = ""
			conta_vencidos = ""
			conta_vencendo = ""
			conta_ok = ""%>
			
		</table>
		</td>
		<td>
			<table align="center" width="40" onclick="coren();">
				
				<%'****************
				'***	COREN	***
				'******************
				ordem_funcionarios = "dt_rgproval,dt_admissao,nm_nome"
					
					cor_linha = "#FFFFFF"
					cor = 1
					cabeca_coren = 0
					conta_func = 0
					if cd_situacao = 1 OR cd_situacao = "" then
						xsql = "SELECT *,DATEDIFF(m,GETDATE(),dt_rgproval) AS validade FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao > '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND cd_conspro <> 1 "&busca_nome&" "&busca_sexo&" OR dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao is NULL AND cd_conspro <> 1 "&busca_nome&" "&busca_sexo&" ORDER BY "&ordem_funcionarios&""
					end if
						Set rs_func = dbconn.execute(xsql)
							
							while not rs_func.EOF
								conta_func = conta_func + 1	
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
									end if%>
									
										<%if cabeca_coren = 0 AND isnull(val_coren) Then
										cabeca_coren = 1
										cabeca_coren1 = 1
										conta_vazios = 0
										
										elseif cabeca_coren = 1 AND val_coren > -120 Then
										cabeca_coren = 2
										cabeca_coren1 = 1
										conta_vencidos = 0
										
										elseif cabeca_coren = 2 AND val_coren >= 1 then
										cabeca_coren = 3
										cabeca_coren1 = 1
										conta_vencendo = 0
										
										elseif cabeca_coren = 3 AND val_coren >= 3 then
										cabeca_coren = 4
										cabeca_coren1 = 1
										conta_ok = 0
										end if%>
								
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
							rs_func.movenext
							wend%>				
					<%'cabeca = ""
					
					lista_empresa = ""
			'*******************************************************%>
				<tr>
					<td align="center" colspan="2" bgcolor="silver"><b>CORENS</b></td>
				</tr>
				<tr>
					<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>		
				</tr>
				<tr>
					<td>Ñ Preenchidos</td>
					<td align="right"><%=conta_vazios%></td>
				</tr>
				<tr style="color:red">
					<td>Vencidos</td>
					<td align="right"><%=conta_vencidos%></td>
				</tr>
				<tr>
					<td>Vencendo</td>
					<td align="right"><%=conta_vencendo%></td>
				</tr>
				<tr>
					<td>Ok</td>
					<td align="right"><%=conta_ok%></td>
				</tr>
				<tr style="background-color:#e0e0e0;">
					<td><b>Total</b></td>
					<td align="right"><b><%=conta_vazios+conta_vencidos+conta_vencendo+conta_ok%></b></td>
				</tr>
				<%conta_vazios = ""
				conta_vencidos = ""
				conta_vencendo = ""
				conta_ok = ""%>
			</table>
		</td>
	</tr>
	
	<tr><td>&nbsp;</td></tr>

	
	
	
</table>
</body>
</html>
