<!--DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"-->

<script language="javascript">
<!-- //1� Passo. Mec�nica da busca
function ordem(a,ordem_total,obj){
	a=a;
	objeto=obj;
	ordem_total=ordem_total;
		if (ordem_total != ''){
			virg =  ',';
			}
		else{
		virg= ''
		}		
	ordem_total = ordem_total + virg
		if (a != ""){
			ordem_total = ordem_total + a;}
		
	document.form.ordem_res.value=a;
	document.form.ordem_total.value=(ordem_total.replace(",,", ","));		
	
	var el=document.getElementById(objeto);
	el.style.display=(el.style.display!='none'?'none':'');
	
	document.getElementsByName(objeto)[1].value= document.form.ordem_inicial.value+'�';
	document.form.ordem_inicial.value = document.form.ordem_inicial.value*1+1;
}
function limpa_ordem(obj){
	document.form.ordem_total.value='';
	
	document.form.ordem_1.value='';
	document.form.ordem_2.value='';
	document.form.ordem_3.value='';
	document.form.ordem_4.value='';
	document.form.ordem_5.value='';
	document.form.ordem_6.value='';
	document.form.ordem_7.value='';
	document.form.ordem_8.value='';
	document.form.ordem_9.value='';
	document.form.ordem_10.value='';
	document.form.ordem_11.value='';
	document.form.ordem_12.value='';
	document.form.ordem_13.value='';
	document.form.ordem_14.value='';
	document.form.ordem_15.value='';
	document.form.ordem_16.value='';
	document.form.ordem_17.value='';
	document.form.ordem_18.value='';
	document.form.ordem_19.value='';
	document.form.ordem_20.value='';
	document.form.ordem_21.value='';
	document.form.ordem_22.value='';
	document.form.ordem_23.value='';
	
	document.form.ordem_inicial.value='1';
	
	document.ordem_1.style.display='inline';
	document.ordem_2.style.display='inline';
	document.ordem_3.style.display='inline';
	document.ordem_4.style.display='inline';
	document.ordem_5.style.display='inline';
	document.ordem_6.style.display='inline';
	document.ordem_7.style.display='inline';
	document.ordem_8.style.display='inline';
	document.ordem_9.style.display='inline';
	document.ordem_10.style.display='inline';
	document.ordem_11.style.display='inline';
	document.ordem_12.style.display='inline';
	document.ordem_13.style.display='inline';
	document.ordem_14.style.display='inline';
	document.ordem_15.style.display='inline';
	document.ordem_16.style.display='inline';
	document.ordem_17.style.display='inline';
	document.ordem_18.style.display='inline';
	document.ordem_19.style.display='inline';
	document.ordem_20.style.display='inline';
	document.ordem_21.style.display='inline';
	document.ordem_22.style.display='inline';
	document.ordem_23.style.display='inline';
}
// -->	
</script>
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
	campo_pis = Instr(1,campos,"pis",0)
	campo_tit_eleitor = Instr(1,campos,"tit_eleitor",0)
	campo_natural = Instr(1,campos,"natural",0)
	campo_conta_sal = Instr(1,campos,"conta_sal",0)
	campo_filia = Instr(1,campos,"filia",0)
	campo_civil = Instr(1,campos,"civil",0)
	campo_conjuge = Instr(1,campos,"conjuge",0)
	campo_contatos = Instr(1,campos,"contatos",0)
	campo_nascimento = Instr(1,campos,"nascimento",0)
	campo_dependentes = Instr(1,campos,"dependentes",0)
	campo_sptrans = Instr(1,campos,"sptrans",0)

'***** Detalhamento da busca ************************************
nome = request("nome")
	if nome <> "" Then
		busca_nome = "AND nm_nome like '"&nome&"%'"
	end if 
sexo = request("sexo")
	if sexo > "0" Then
		busca_sexo = "AND cd_sexo = '"&sexo&"'"
	end if 

func_ativos = request("func_ativos")


'***** 2� Passo. Ordem dos registros da Busca *******************
ordem_total = request("ordem_total")
	
	
	ver_string_ordem = instr(ordem_total,",")
		if ver_string_ordem = 1 then
			ordem_total = mid(ordem_total,2,len(ordem_toral))
		end if
	
			if ordem_total <> "" Then
				ordem_funcionarios = ordem_total
			else
				ordem_funcionarios = "cd_contrato,nm_nome,nm_sigla"
				'ordem_funcionarios = "cd_contrato,nm_nome"
			end if
		
		ordem_inicial = request("ordem_inicial")
		ordem_1 = request("ordem_1")
		ordem_2 = request("ordem_2")
		ordem_3 = request("ordem_3")
		ordem_4 = request("ordem_4")
		ordem_5 = request("ordem_5")
		ordem_6 = request("ordem_6")
		ordem_7 = request("ordem_7")
		ordem_8 = request("ordem_8")
		ordem_9 = request("ordem_9")
		ordem_10 = request("ordem_10")
		ordem_11 = request("ordem_11")
		ordem_12 = request("ordem_12")
		ordem_13 = request("ordem_13")
		ordem_14 = request("ordem_14")
		ordem_15 = request("ordem_15")
		ordem_16 = request("ordem_16")
		ordem_17 = request("ordem_17")
		ordem_18 = request("ordem_18")
		ordem_19 = request("ordem_19")
		ordem_20 = request("ordem_20")
		ordem_21 = request("ordem_21")
		ordem_22 = request("ordem_22")
		ordem_23 = request("ordem_23")
%>
<head>
	<title>Listagem de funcion�rios</title>
</head>

<table align="center" border="1" id="no_print">
	<tr><td colspan="100%" align="center" style="background-color:#000099; color:white; font-size=15; font-weight:bold;">LISTAGEM</td></tr>
	<form action="../../empresa.asp" method="post" name="form">
	<input type="hidden" name="tipo" value="lista">
	
	<!-- *** 3� Passo *** Campos de referencia da Ordem ***-->
	<input type="hidden" name="ordem_res">
	<input type="hidden" name="ordem_total" value="<%=ordem_total%>">
	<input type="hidden" name="ordem_inicial" value="<%if ordem_inicial = "" Then%>1<%else response.write(ordem_inicial) end if%>">
	<tr style="background-color:#0033ff; color:white; font-weight:bold;">
		<td>Data</td>
		<td>Mostrar</td>
		<td>Detalhes</td>
		<td><a href="javascript:void();" onClick="limpa_ordem()" title="Limpa a ordem">Ordem</a></td>
	</tr>
	<tr id="no_print" style="background-color:#ccffff; color:black; font-weight:bold;">
		<td valign="top" align="center">DE:
			<input type="text" name="dia_sel" size="2" maxlength="2" value="<%=dia_sel%>" style="background-color:white;">
			<input type="text" name="mes_sel" size="2" maxlength="2" value="<%=mes_sel%>" style="background-color:white;">
			<input type="text" name="ano_sel" size="4" maxlength="4" value="<%=ano_sel%>" style="background-color:white;">			
			<%'=dia_sel&"/"&mes_sel&"/"&ano_sel&"<br>("&ano_sel&mes_sel&")"%>		
		</td>
		<td>
			&nbsp; Nome do funcion�rio
		</td>
		<td>
			<input type="text" name="nome" size="20" maxlength="100" value="<%=nome%>" style="background-color:white;">
		</td>
		<td><%ordem_var = ordem_1
		ordem_num="ordem_1"
		ordem_campo="nm_nome"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="background-color:#ccffff; border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="sexo" <%if campo_sex <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Sexo</td>
			<%if sexo = 1 Then%><%ck_sx1 = "checked"%><%Elseif sexo = 2 then%><%ck_sx2 = "checked"%><%else%><%ck_sx = "checked"%><%End if%>
		<td><input type="radio" name="sexo" value="1" <%=ck_sx1%>>M&nbsp;<input type="radio" name="sexo" value="2" <%=ck_sx2%>>F&nbsp;<input type="radio" name="sexo" value="0" <%=ck_sx%>>Ambos</td>
		<td><%ordem_var = ordem_2
		ordem_num="ordem_2"
		ordem_campo="cd_sexo"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;Status</td>
		<td><input type="checkbox" name="campos" value="endereco" <%if campo_endereco <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Endere�o</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_3
		ordem_num="ordem_3"
		ordem_campo="nm_endereco"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>	
	</tr>
	<tr>
		<%if func_ativos = "" OR func_ativos = 1 then
			func_ck1 = "checked"
		elseif func_ativos = 2 then
			func_ck2 = "checked"
		end if%>
		<td>&nbsp;<input type="radio" name="func_ativos" value="1" <%=func_ck1%>> Ativo</td>
		<td><input type="checkbox" name="campos" value="rg" <%if campo_rg <> 0 Then response.write("CHECKED") end if%> style=border:0px;> RG</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_4
		ordem_num="ordem_4"
		ordem_campo="nm_rg"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;<input type="radio" name="func_ativos" value="2" <%=func_ck2%>>Inativo</td>
		<td><input type="checkbox" name="campos" value="cpf" <%if campo_cpf <> 0 Then response.write("CHECKED") end if%> style=border:0px;> CPF</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_5
		ordem_num="ordem_5"
		ordem_campo="nm_cpf"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="rgp" <%if campo_rgp <> 0 Then response.write("CHECKED") end if%> style=border:0px;> COREN<!--Registro Prof.--></td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_6
		ordem_num="ordem_6"
		ordem_campo="cd_codigo"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="cargo" <%if campo_cargo <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Cargo</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_7
		ordem_num="ordem_7"
		ordem_campo="nm_cargo"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="unidade" <%if campo_unidade <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Unidade</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_8
		ordem_num="ordem_8"
		ordem_campo="nm_sigla"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="situacao" <%if campo_situacao <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Situacao</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_9
		ordem_num="ordem_9"
		ordem_campo="nm_situacao"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="ctps" <%if campo_ctps <> 0 Then response.write("CHECKED") end if%> style=border:0px;> CTPS</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_10
		ordem_num="ordem_10"
		ordem_campo="cd_ctps, cd_ctps_serie"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="tempo_casa" <%if campo_tempo_casa <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Tempo de casa</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_11
		ordem_num="ordem_11"
		ordem_campo="tempo_casa"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="contratos" <%if campo_contratos <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Contratos</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_12
		ordem_num="ordem_12"
		ordem_campo="cd_contrato"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="pis" <%if campo_pis <> 0 Then response.write("CHECKED") end if%> style=border:0px;> PIS</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_13
		ordem_num="ordem_13"
		ordem_campo="cd_pis"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="tit_eleitor" <%if campo_tit_eleitor <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Tit. Eleitor</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_14
		ordem_num="ordem_14"
		ordem_campo="nm_tit_eleitor"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="natural" <%if campo_natural <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Naturalidade</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_15
		ordem_num="ordem_15"
		ordem_campo="nm_naturalidade"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="conta_sal" <%if campo_conta_sal <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Conta Sal�rio</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_16
		ordem_num="ordem_16"
		ordem_campo="nm_conta_sal"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="filia" <%if campo_filia <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Filia��o</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_17
		ordem_num="ordem_17"
		ordem_campo="nm_mae,nm_pai"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="civil" <%if campo_civil <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Est. Civil</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_18
		ordem_num="ordem_18"
		ordem_campo="nm_estado_civil"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="conjuge" <%if campo_conjuge <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Conjuge</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_19
		ordem_num="ordem_19"
		ordem_campo="nm_conjuge"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="contatos" <%if campo_contatos <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Contatos</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_20
		ordem_num="ordem_20"
		ordem_campo="nm_contato"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="nascimento" <%if campo_nascimento <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Nascimento</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_21
		ordem_num="ordem_21"
		ordem_campo="dt_nascimento"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="dependentes" <%if campo_dependentes <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Dependentes</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_22
		ordem_num="ordem_22"
		ordem_campo="dt_nascimento"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" name="campos" value="transporte" <%if campo_transporte <> 0 Then response.write("CHECKED") end if%> style=border:0px;> Transporte</td>
		<td>&nbsp;</td>
		<td><%ordem_var = ordem_23
		ordem_num="ordem_23"
		ordem_campo="cd_sptrans, cd_cmt_bom"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	
		
	<tr id="no_print">
		<td><input type="submit" name="OK" value="Mostrar" style="background-color:gray; color:white;"></td>
		<td colspan="2">&nbsp;<a href="empresa.asp?tipo=novo"><img src="../../imagens/ic_novo.gif" alt="" width="10" height="12" border="0">Novo colaborador</a></td>
		<td><a href="javascript:void();" onClick="limpa_ordem()" title="Limpa a ordem">limpa ordem</a></td>
	</tr>
	</form>
</table>
<table align="center">
	<tr>
					<%cor_linha = "#FFFFFF"
					cor = 1
					cabecalho = 0
					
					if func_ativos = "" OR func_ativos = 1 Then
						'xsql = "up_funcionario_contrato_lista @dt_data='"&ano_sel&mes_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
						'xsql = "up_funcionario_contrato_lista @dt_data='"&ano_sel&mes_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_sel  &"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
						xsql = "up_funcionario_contrato_lista @dt_data='"&ano_sel&mes_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
						Set rs_func = dbconn.execute(xsql)
					'	response.write(func_ativos)
							
					elseif func_ativos = "2" Then
				xsql = "up_funcionario_contrato_lista_inativos @dt_data='"&ano_sel&mes_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
						Set rs_func = dbconn.execute(xsql)
					'	response.write(func_ativos)
					end if
						
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
								cd_pis = rs_func("cd_pis")
								nm_tit_eleitor = rs_func("nm_tit_eleitor")
								nr_tit_eleitor_zona = rs_func("nr_tit_eleitor_zona")
								nr_tit_eleitor_seccao = rs_func("nr_tit_eleitor_seccao")
								
								nm_rg = rs_func("nm_rg")
								nm_cpf = rs_func("nm_cpf")
								
								dt_admissao = zero(day(rs_func("dt_admissao")))&"/"&zero(month(rs_func("dt_admissao")))&"/"&year(rs_func("dt_admissao"))
								dt_demissao = zero(day(rs_func("dt_demissao")))&"/"&zero(month(rs_func("dt_demissao")))&"/"&year(rs_func("dt_demissao"))
								
								admissao = rs_func("admissao")
								demissao = rs_func("demissao")
								
								cd_contrato = rs_func("cd_contrato")
								
								tempo_casa = rs_func("tempo_casa")
								cd_unidade = rs_func("cd_unidade")
								nm_sigla = rs_func("nm_sigla")
								
								nm_naturalidade = rs_func("nm_naturalidade")
								nm_mae = rs_func("nm_mae")
								nm_pai = rs_func("nm_pai")
								
								nm_banco = rs_func("nm_banco")
								cd_banco_ag = rs_func("cd_banco_ag")
								cd_banco_conta = rs_func("cd_banco_conta")
								nm_estado_civil = rs_func("nm_estado_civil")
								nm_conjuge = rs_func("nm_conjuge")
								dt_nascimento = rs_func("dt_nascimento")
								
								xsql_regprof = "Select * From TBL_funcionario Where cd_codigo = "&cd_funcionario&""
								Set rs_regprof = dbconn.execute(xsql_regprof)
									if not rs_regprof.EOF then
										cd_numreg = rs_regprof("cd_numreg")
									end if
								
								'*** CARGO ***
								xsql_cargo = "Select * From View_funcionario_cargo Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 AND dt_atualizacao <= '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' ORDER BY dt_atualizacao desc"
								Set rs_cargo = dbconn.execute(xsql_cargo)
									if not rs_cargo.EOF then
										nm_cargo = rs_cargo("nm_cargo")
									end if
								
								'*** HOR�RIO ***
								xsql = "Select * From View_funcionario_horario Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 order by dt_atualizacao DESC"
								Set rs_hora = dbconn.execute(xsql)
									if not rs_hora.EOF Then
										hr_entrada = zero(hour(rs_hora("hr_entrada")))&":"&zero(minute(rs_hora("hr_entrada")))
										hr_saida = zero(hour(rs_hora("hr_saida")))&":"&zero(minute(rs_hora("hr_saida")))
										nm_intervalo = rs_hora("nm_intervalo")
									end if
								
								'*** UNIDADE ***
								''xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 order by dt_atualizacao DESC"
								''xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 AND dt_atualizacao between '"&mes_sel&"/1/"&ano_sel&"' AND '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"'order by dt_atualizacao DESC"
								xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 AND dt_atualizacao <= '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' order by dt_atualizacao DESC"
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
								
								
								
								'*** Mostra o cabe�alho ****
									'if cabeca_ativo <>  cd_contrato then
									if cabecalho = "0" then%>
										<tr id="no_print">
											<td colspan="100%">&nbsp;</td>										
										</tr>
										<tr bgcolor=#c0c0c0>
											<td>&nbsp;</td>
											<td><b>Mat.</b></td>
											<td><b>Funcion�rio</b></td>
											<%if campo_sex <> 0 Then%><td><b>Sexo</b></td> <%end if%>
											<%if campo_endereco <> 0 Then%><td><b>Endere�o</b></td><%end if%>
											<%if campo_rg <> 0 Then%><td><b>RG</b></td><%end if%>
											<%if campo_cpf <> 0 Then%><td><b>CPF</b></td><%end if%>
											<td><b>Contrato</b></td>
											<%if campo_ctps <> 0 Then%><td><b>CTPS</b></td><%end if%>
											<%if campo_pis <> 0 Then%><td><b>PIS</b></td><%end if%>
											<%if campo_tit_eleitor <> 0 Then%><td><b>Titulo de Eleitor</b></td><%end if%>											
											<%if campo_rgp <> 0 Then%><td><b>COREN<!--Reg. Profissional--></b></td><%end if%>
											<%if campo_cargo <> 0 Then%><td><b>Cargo</b></td><%end if%>
											<%if campo_hora <> 0 Then%><td><b>Hor�rio</b></td><%end if%>
											<%if campo_unidade <> 0 Then%><td><b>Unidade</b></td><%end if%>
											<%if campo_situacao <> 0 Then%><td><b>Situa��o</b></td><%end if%>
											<%if campo_natural <> 0 Then%><td><b>Naturalidade</b></td><%end if%>
											<%if campo_conta_sal <> 0 Then%><td colspan="3"><b>Conta Salario</b></td><%end if%>
											<%if campo_filia <> 0 Then%><td><b>M�e</b></td><td><b>Pai</b></td><%end if%>
											<%if campo_civil <> 0 Then%><td><b>Est. Civil</b></td><%end if%>											
											<%if campo_conjuge <> 0 Then%><td><b>Conjuge</b></td><%end if%>
											<%if campo_contatos <> 0 Then%><td><b>Contatos</b></td><%end if%>
											<%if campo_nascimento <> 0 Then%><td><b>Nascimento</b></td><%end if%>											
											<%if campo_dependentes <> 0 Then%><td><b>Dependentes</b></td><%end if%>											
											
											<%if cd_situacao = 2 then%><td><b>Empresa</b></td><%end if%>
											<!--td><b>Admiss�o</b></td>
											<%'if cd_situacao = 2 then%>
											<td><b>demiss�o</b></td-->
											<%'end if%>
											<td><b>Admiss�o</b></td>
											<td><b>Recis�o</b></td>
											<%if campo_tempo_casa <> 0 then%><td><b>Tempo de casa</b></td><%end if%>
											<%if campo_contratos <> 0 Then%><td colspan="2" align="center"><b>Todos os contratos</b></td><%end if%>
											
										</tr>
									<%cabecalho = 1
									conta_func = 0
									end if%>
								
								<%'*** Lista os funcionarios ***%>
								<!--tr-->
								<%if int(periodo_sel)=int(demissao) then
									cor_registro = "gray"
								else
									cor_registro = "black"
								end if%>
								<!--tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px; color:<%=cor_registro%>;" bgcolor="<%=cor_linha%>"-->
								<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px; color:<%=cor_registro%>;" bgcolor="<%=cor_linha%>">
									<td align="center" valign="top" style="color:gray;">
										<%if int(periodo_sel) = int(demissao) then
										else
											conta_func = conta_func + 1
											response.write(zero(conta_func))
										end if%>
										<%'=%></td>
									<td align="right" valign="top"><%=cd_matricula%></td>
									<!--td valign="top" onclick="javascript:window.open('empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');"><%=nm_funcionario%></td-->
									<td valign="top" onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');"><%=nm_funcionario%></td>
									<%if campo_sex <> 0 Then%><td valign="top"><%=nm_sexo%></td><%end if%>
									<%if campo_endereco <> 0 Then%><td valign="top"><%=nm_endereco%>, <%=nr_numero%> &nbsp; <br id="mok_print"><%=nm_bairro%> - <%=nm_cidade%>-<%=nm_estado%>&nbsp;</td><%end if%>
									<%if campo_rg <> 0 Then%><td valign="top"><%=nm_rg%></td><%end if%>
									<%if campo_cpf <> 0 Then%><td valign="top"><%=nm_cpf%></td><%end if%>
									<td valign="top"><%=nm_regime_trab%></td>
									<%if campo_ctps <> 0 Then%><td valign="top"><%=cd_ctps&"-"&cd_ctps_serie%></td><%end if%>
									<%if campo_pis <> 0 Then%><td valign="top"><%=cd_pis%></td><%end if%>
									<%if campo_tit_eleitor <> 0 Then%><td valign="top"><b>Tit.: </b><%=nm_tit_eleitor%><br><b>Zona: </b><%=nr_tit_eleitor_zona%><br><b>Sec��o: </b><%=nr_tit_eleitor_seccao%></td><%end if%>									
									<%if campo_rgp <> 0 Then%><td valign="top"><%=cd_numreg%></td><%end if%>
									<%if campo_cargo <> 0 Then%><td valign="top"><%=nm_cargo%></td><%end if%>
									<%if campo_hora <> 0 Then%><td valign="top"><%=hr_entrada%> - <%=hr_saida%></td><%end if%>
									<%if campo_unidade <> 0 Then%><td valign="top"><%'=cd_unidade%><%=nm_sigla%><%=teste%></td><%end if%>
									<%if campo_situacao <> 0 Then%><td valign="top"><%=nm_status%></td><%end if%>
									<%if campo_natural <> 0 Then%><td valign="top"><%=nm_naturalidade%></td><%end if%>
									<%if campo_conta_sal <> 0 Then%><td valign="top"><%=nm_banco%>
									<td valign="top"><%=cd_banco_ag%>
									<td valign="top"><%=cd_banco_conta%></td><%end if%>
									<%if campo_filia <> 0 Then%><td valign="top"><%=nm_mae%></td>
									<td valign="top"><%=nm_pai%></td><%end if%>
									<%if campo_civil <> 0 Then%><td valign="top"><%=nm_estado_civil%></td><%end if%>
									<%if campo_conjuge <> 0 Then%><td valign="top"><%=nm_conjuge%></td><%end if%>
									<%if campo_nascimento <> 0 Then%><td valign="top"><%=dt_nascimento%></td><%end if%>
									<%if campo_dependentes <> 0 Then%><td valign="top">
										<%'*** DEPENDENTES ***
										xsql = "Select * From VIEW_dependentes Where cd_funcionario = "&cd_funcionario&" order by dt_nascimento"
										Set rs = dbconn.execute(xsql)
											while not rs.EOF
												nm_nome = rs("nm_nome")
												nm_parentesco = rs("nm_parentesco")
												nasc_depend = rs("dt_nascimento")
												'nm_numero = rs("nm_numero")
												'nm_obs = rs("nm_obs")%>
											<%=nm_nome&" ("&nm_parentesco&")"&"-"&mesdoano(month(nasc_depend))&"/"&right(nasc_depend,2)%><br>
											<%rs.movenext
											wend%></td><%end if%>
									
									<%if campo_contatos <> 0 Then%><td valign="top">
										<%'*** CONTATOS (COMUNICA��O) ***
										xsql = "Select * From VIEW_contatos Where id_origem = "&cd_funcionario&""
										Set rs_contato = dbconn.execute(xsql)
											while not rs_contato.EOF
												cd_categoria =  rs_contato("cd_categoria")
												nm_nome = rs_contato("nm_nome")
												grau_relac = rs_contato("nm_cargo")
												cd_ddd = rs_contato("cd_ddd")
												nm_numero = rs_contato("nm_numero")
												nm_obs = rs_contato("nm_obs")%>
											<%if cd_categoria = 5 then%><b style="color:red;"><%=nm_nome&" ("&grau_relac&")"%></b><%end if%>
											<%=cd_ddd%><%if cd_categoria <> 3 then%>-<%end if%> 
											<%=nm_numero%><%if nm_obs <> "" then%>-<%=nm_obs%><%end if%><br>
											<%rs_contato.movenext
											wend%></td><%end if%>
									
									
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
								<td  id="no_print"><img src="../../imagens/ic_editar.gif" alt="Editar" width="13" height="14" border="0" onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');"></td>
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
		
	<%if func_ativos <> 2 Then
		
		'*** conta os funcion�rios de cada contrato ***
				'*** seleciona o m�s anterior ***
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
			<%'*** conta os funcion�rios de cada contrato ***
				'*** seleciona o m�s anterior ***
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
			<tr><td colspan="3">Saldo inicio M�s: <b><%=int(conta)%></b> <%'="("&data_sel&")"%></td></tr>
			<tr><td colspan="3">Admiss�es: <b style="color:green;"><%=total_admissao%></b></td></tr>
			<tr><td colspan="3">Recis�es: <b style="color:red;"><%=total_recisao%></b></td></tr>
			<tr><td colspan="3">TOTAL: <b><%total_geral=(conta+total_admissao)-total_recisao%> <%=total_geral%></b></td></tr>			
			<!--tr><td><%=data_sel%></td></tr-->
		<%total_funcionarios = total_funcionarios + total_geral
		
		conta = 0
		rs_contr.movenext
		wend%>
		<tr><td colspan="3"><hr></td></tr>
		<tr><td colspan="3"><b>HEAD COUNT: <%=total_funcionarios%></b> <%if total_funcionarios > 1 then%>colaboradores<%else%>Colaborador<%end if%> </td></tr>
		<%end if%>
		<tr>
			<td colspan="100%">
				SELECT * FROM VIEW_funcionario_contrato_lista <br>
				WHERE dt_admissao <= <%=ano_sel&mes_sel%> AND dt_demissao >= <%=ano_sel&mes_sel%>  OR dt_admissao <= <%=ano_sel&mes_sel%>  AND dt_demissao IS NULL <br>
				ORDER BY <%=ordem_funcionarios%><br>
				<%=dia_final&"/"&mes_sel&"/"&ano_sel%>
			</td>
		</tr>
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
