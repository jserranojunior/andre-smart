<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<%
dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

if dia_sel = "" or mes_sel = "" or ano_sel = "" then
	dia_sel = zero(day(now))
	mes_sel = zero(month(now))
	ano_sel = year(now)
end if

data_selecionada = dia_sel&"/"&mes_sel&"/"&ano_sel

	dia_hoje = day(now)
	mes_hoje = month(now)
	ano_hoje = year(now)
data_hoje = zero(dia_hoje)&"/"&zero(mes_hoje)&"/"&ano_hoje

'data_comparativa_demissao = mes_hoje&"/"&dia_hoje&"/"&ano_hoje
	dem_dia = dia_hoje
	dem_mes = mes_hoje-1
	dem_ano = ano_hoje
	
	if dem_mes < 1 then
		dem_mes = 12
		dem_ano = dem_ano -1
	end if
data_comparativa_demissao = dem_mes&"/"&dem_dia&"/"&dem_ano

'****************************************************
cd_situacao = request("cd_situacao")
ordem_funcionarios = request("ordem_funcionarios")

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
	if sexo > "0" Then
		busca_sexo = "AND cd_sexo = '"&sexo&"'"
	end if 
%>
<head>
	<title>Listagem de funcionários</title>
</head>

<body><br>

<br id="no_print">
<table align="center">
	<tr><td colspan="7"><a href="empresa.asp?tipo=dependentes"><img src="../../imagens/reload6.gif" alt="" width="10" height="12" border="0" id="no_print"></a></td></tr>
	<tr><td colspan="7" align="center"><b>LISTAGEM DE DEPENDENTES</b></td></tr>
	<%'ultima_faixa_idade = 0
	'strsql = "SELECT * FROM View_funcionario_dependentes where dt_demissao IS NULL OR dt_demissao >= '"&data_hoje&"' order by idade desc"
	'strsql = "SELECT * FROM View_funcionario_dependentes where dt_demissao IS NULL OR dt_demissao = '"&month(data_hoje)-1&"/"&ultimodiames(month(data_hoje)-1,year(data_hoje))&"/"&year(data_hoje)&"'"' order by idade desc"
	strsql = "SELECT * FROM View_funcionario_dependentes where dt_demissao IS NULL OR dt_demissao >= '"&data_comparativa_demissao&"' order by idade desc"
	Set rs = dbconn.execute(strsql)
	
	cabeca = 0
	num_linha = 1
	conta_linha = 0
	while not rs.EOF
		nm_dependente = rs("nm_nome")
		nm_idade = rs("idade")
			nm_idade_ano = nm_idade/12
			nm_idade_mes = nm_idade mod 12
			nm_idade_atual = zero(int(nm_idade_ano))&"a "&zero(nm_idade_mes)&"m"
			faixa_idade = nm_idade
			
		dt_nascimento = rs("dt_nascimento")
		nm_sexo = rs("cd_sexo")
			if nm_sexo = "1" then
				nm_sexo = "Masc"
			elseif nm_sexo = "2" then
				nm_sexo = "Fem"
			end if
		nm_parentesco = rs("nm_parentesco")
		cd_funcionario = rs("cd_funcionario")
		nm_funcionario = rs("nm_funcionario")
		
		if conta_linha = 0 then
			cor_linha = "#ffffff"
		elseif conta_linha = 1 then
			cor_linha = "#d8d8d8"
		end if
		dt_demissao = rs("dt_demissao")
			if dt_demissao <> "" Then
				cor_demissao = "#ff0000"
			end if
	
	if cabeca = 0 then%>
		<tr style="background-color:gray; color:#ffffff;">
			<td>&nbsp;</td>
			<td>Dependente</td>
			<td>Idade</td>
			<td>Nasc.</td>
			<td>Sexo</td>
			<td>Parentesco</td>
			<td>Funcionario</td>
			<td>Demissão</td>
		</tr>		
	<%end if%>
	
	<%if faixa_idade <= 83 then
		faixa_idade_atual = 83
			if ultima_faixa_idade <> 83 then%>
				<tr><td colspan="8">&nbsp;</td></tr>
				<tr style="background-color:silver;color:white;"><td>&nbsp;</td><td colspan="7" >Até 7 anos <%'=" - "&faixa_idade_atual&":"&ultima_faixa_idade%></td></tr>
			<%end if
			ultima_idade_atual = 83%>
			
	<%elseif faixa_idade <= 167 then
		faixa_idade_atual = 167
			if ultima_faixa_idade <> 167 then%>
				<tr><td colspan="8">&nbsp;</td></tr>
				<tr style="background-color:silver;color:white;"><td>&nbsp;</td><td colspan="7" >Até 14 anos</td></tr>
			<%end if
			ultima_idade_atual = 167%>
			
	<%elseif faixa_idade <= 287 then
		faixa_idade_atual = 287
			if ultima_faixa_idade <> 287 then%>
				<tr><td colspan="8">&nbsp;</td></tr><br>
				<tr style="background-color:silver;color:white;"><td>&nbsp;</td><td colspan="7" >Até 24 anos</td></tr>
			<%end if
			ultima_idade_atual = 287%>
			
	<%elseif faixa_idade > 287 then
		faixa_idade_atual = 288
			if ultima_faixa_idade <> 1200 then%>
				<tr style="background-color:silver;color:white;"><td>&nbsp;</td><td colspan="7" >Maior que 24 anos</td></tr>
			<%end if
			ultima_idade_atual = 288%>
	<%end if%>
	<tr style="background-color:<%=cor_linha%>;">
		<td><%=zero(num_linha)%></td>
		<td><%=nm_dependente%></td>
		<td><%=nm_idade_atual%></td>
		<!--td><%'=nm_idade%></td-->
		<td><%=zero(day(dt_nascimento))&"/"&zero(month(dt_nascimento))&"/"&year(dt_nascimento)%></td>
		<td><%=nm_sexo%></td>
		<td><%=nm_parentesco%></td>
		<td onclick="javascript:window.open('empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=cd_funcionario%>&busca=<%=strbusca%>');"><a href="javascript:void(0);"return false;"><%=nm_funcionario%></a></td>
		<td><i style="color:<%=cor_demissao%>"><%=dt_demissao%></i></td>
	</tr>
	<%ultima_faixa_idade = faixa_idade_atual
	cabeca = cabeca + 1
	num_linha = num_linha + 1
	conta_linha = conta_linha + 1
		if conta_linha = 2 then
			conta_linha = 0
		end if
	rs.movenext
	wend%>
	<tr>
		<td><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
	</tr>
</table>


</body>
</html>
