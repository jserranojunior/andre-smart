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
	
'*** ORDEM ***
	ordem_funcionarios = "nr_scr, nm_nome desc,dt_admissao desc"
	
%>


<head>
	<title>Listagem de funcionários</title>
</head>

<body><br id="no_print">

<table align="center" width="40">
	<tr><td colspan="4"><a href="empresa.asp?tipo=resumo_venc">Gestão de enfermagem >></a></td></tr>
	<tr><td>&nbsp;</td></tr>
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
						nr_scr = rs_func("nr_scr")
							if nr_scr = 0 then
								scr_nao = "X"
							elseif nr_scr = 1 then
								scr_sim = "X"
							end if
						
					'*** Mostra o cabeçalho ****
						'if cabeca_ativo <> cd_contrato then
						if cabeca_ativo = 0 Then%>
							<tr><td colspan="100%" align="center" style="background-color:gray; color:white; font-size:14px;"><b>VACINAÇÃO - SCR </b><span style="font-size:9px; color:silver;">(Sarampo, Caxumba, Rubéola)</span></td></tr>
						<%end if%>
							
							<%'if cabeca_exam = 0 AND isnull(nr_hepatite_b) Then
							if cabeca_scr = 0 AND nr_scr = 0 Then%>
								<tr><td align="center" colspan="8" style="color:white;" bgcolor="#808080"><span style="color:#ff0000;font-size:12px;"><<<</span> NÃO <span style="color:#ff0000;">>>></span></td></tr>
							<%cabeca_scr = 1
							cabeca_scr1 = 0
							cor_val = "#ff0000"
							conta_nao = 0
							
							elseif cabeca_scr = 1 AND nr_scr = 1 Then%>
								<tr><td align="center" colspan="8" style="color:white;" bgcolor="#808080"><span style="color:#00ff40;font-size:12px;"><<<</span> SIM <span style="color:#00ff40;">>>></span></td></tr>
							<%cabeca_scr = 2
							cabeca_scr1 = 0
							cor_val = "#c18509"
							conta_sim = 0
							end if%>
							
							<%if cabeca_scr1 = 0 then%><tr bgcolor=#c0c0c0>
								<td>&nbsp;</td>
								<td align="center"><b>Mat.</b></td>
								<td align="center"><b>Funcionário</b></td>
								<!--td align="center"><b>Admissão</b></td-->
								<td align="center">NÃO</td>
								<td align="center">SIM</td>
							</tr>
							<%cabeca_scr1 = 1
							end if%>
					
					<%'*** Lista os funcionarios ***%>
					<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');" onclick="javascript:window.open('empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>','Cadastro_<%=rs_func("cd_funcionario")%>','width=760,height=700,scrollbars=1');" style="height:20px;" bgcolor="<%=cor_linha%>">
					<!-- tr onmouseover="mOvr(this,'#CFC8FF');" onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px;" bgcolor="<%=cor_linha%>"-->
						<td align="center" valign="top" style="color:gray;"><%=zero(conta_func)%></td>
						<td align="right" valign="top"><%=zero(cd_matricula)%></td>
						<td valign="top"><%=nm_funcionario%></td>
						<!--td valign="top" align="center"><%=dt_admissao%></td-->
						<td align="center" style="color:red;"><b><%=scr_nao%></b></td>
						<td align="center" style="color:green;"><b><%=scr_sim%></b></td>
					</tr>
					
					<%cabeca_ativo = cd_contrato
						
						if conta_sim = "" then
							conta_nao = conta_nao + 1
						elseif conta_nao <> "" then
							conta_sim = conta_sim + 1
						end if
						
					if cor > 0 then
						cor_linha = "#d7d7d7"
						cor = 0
					else
						cor_linha = "#FFFFFF"
						cor = 1
					end if
					
					scr_nao = ""
					scr_sim = ""
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
		<!--td><img src="../../imagens/blackdot.gif" alt="" width="90" height="1" border="0"></td-->
		<td><img src="../../imagens/blackdot.gif" alt="" width="25" height="1" border="0"></td>
		<td><img src="../../imagens/blackdot.gif" alt="" width="25" height="1" border="0"></td>
	</tr>
	
	<tr>
		<td colspan="2">&nbsp;</td>
		<td>NÃO</td>
		<td align="right"><%=conta_nao%></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
		<td>SIM</td>
		<td align="right"><%=conta_sim%></td>
	</tr>
	
</table>


</body>
</html>
