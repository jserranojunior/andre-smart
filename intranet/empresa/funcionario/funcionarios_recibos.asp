<!--DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"-->


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
modelo = request("modelo")
titulo = request("titulo")
texto = request("texto")
	texto = replace(texto,vbcrlf,"<br>")

cd_situacao = request("cd_situacao")
ordem_funcionarios = request("ordem_funcionarios")
	if ordem_funcionarios = "" Then
		ordem_funcionarios = 1
	end if

campos = request("campos")
	campo_sex = Instr(1,campos,"_sexo",0)
	campo_endereco = Instr(1,campos,"_endereco",0)
	campo_rg = Instr(1,campos,"_rg",0)
	campo_cpf = Instr(1,campos,"_cpf",0)
	campo_rp = Instr(1,campos,"_rp",0)
	campo_cargo = Instr(1,campos,"_cargo",0)
	campo_hora = Instr(1,campos,"_hora",0)
	campo_unidade = Instr(1,campos,"_unidade",0)
	campo_situacao = Instr(1,campos,"_situacao",0)
	campo_ctps = Instr(1,campos,"_ctps",0)
	campo_contratos = Instr(1,campos,"_contratos",0)
	campo_tempo_casa = Instr(1,campos,"_tempo_casa",0)
	campo_pis = Instr(1,campos,"_pis",0)
	campo_tit_eleitor = Instr(1,campos,"_tit_eleitor",0)
	campo_natural = Instr(1,campos,"_natural",0)
	campo_conta_sal = Instr(1,campos,"_conta_sal",0)
	campo_filia = Instr(1,campos,"_filia",0)
	campo_civil = Instr(1,campos,"_civil",0)
	campo_conjuge = Instr(1,campos,"_conjuge",0)
	campo_contatos = Instr(1,campos,"_contatos",0)
	campo_nascimento = Instr(1,campos,"_nascimento,_idade_func",0)
	campo_dependentes = Instr(1,campos,"_dependentes",0)
	campo_transporte = Instr(1,campos,"_transporte",0)
	campo_alimentacao = Instr(1,campos,"_alimentacao",0)
	campo_medica = Instr(1,campos,"_medica",0)
	'campo_odonto = Instr(1,campos,"odonto",0)
	
	
'***** Detalhamento da busca ************************************
nome = request("nome")
	if nome <> "" Then
		busca_nome = " AND nm_nome like ''%"&busca_inteligente(nome)&"%'' "
	end if
sexo = request("sexo")
	if sexo <> "0" Then
		busca_sexo = " AND cd_sexo = "&sexo&""
	end if
endereco = request("endereco")
	if endereco <> "" Then
		busca_endereco = " AND nm_endereco like ''%"&busca_inteligente(endereco)&"%'' "
	end if
numero = request("numero")
	if numero <> "" Then
		busca_numero = " AND nr_numero = "&numero&""
	end if
complemento = request("complemento")
	if complemento <> "" Then
		busca_complemento = " AND nm_complemento like ''%"&busca_inteligente(complemento)&"%'' "
	end if
cep = request("cep")
	if cep <> "" Then
		busca_cep = " AND nm_cep like ''"&cep&"%'' "
	end if
bairro = request("bairro")
	if bairro <> "" Then
		 busca_bairro = " AND nm_bairro like ''%"&busca_inteligente(bairro)&"%'' "
	end if
cidade = request("cidade")
	if cidade <> "" Then
		busca_cidade = " AND nm_cidade like ''%"&busca_inteligente(cidade)&"%'' "
	end if
rg = request("rg")
	if rg <> "" Then
		busca_rg = " AND nm_rg like ''%"&rg&"%'' "
	end if
cpf = request("cpf")
	if cpf <> "" Then
		busca_cpf = " AND nm_cpf like ''%"&cpf&"%'' "
	end if
coren = request("coren")
	if coren <> "" Then
		busca_coren = " AND cd_numreg like ''%"&coren&"%'' "
	end if
cargo = request("cargo")
	if cargo <> "" Then
		busca_cargo = " "
	end if
unidade = request("unidade")
	'if unidade = "0" Then
		'busca_unidade = " "
		'unidade 
	'end if
ctps = request("ctps")
	if ctps <> "" Then
		busca_ctps = " AND cd_ctps like ''%"&ctps&"%'' "
	end if
ctps_serie = request("ctps_serie")
	if ctps_serie <> "" Then
		busca_ctps_serie = " AND cd_ctps_serie like ''%"&ctps_serie&"%'' "
	end if
contratos = request("contratos")
	if contratos <> "" AND contratos <> "0" Then
		'busca_contratos = " AND nm_regime_trab like ''%"&busca_inteligente(contratos)&"%'' "
		busca_contratos = " AND cd_regime_trab = "&contratos
	end if
pis = request("pis")
	if pis <> "" Then
		busca_pis = " AND cd_pis like ''%"&pis&"%'' "
	end if
	
tit_eleitor = request("tit_eleitor")
	if tit_eleitor <> "" Then
		busca_tit_eleitor = " AND nm_tit_eleitor like ''%"&tit_eleitor&"%'' "
	end if
naturalidade = request("naturalidade")
	if naturalidade <> "" Then
		busca_naturalidade = " AND nm_naturalidade like ''%"&busca_inteligente(naturalidade)&"%'' "
	end if
agencia = request("agencia")
	if agencia <> "" Then
		busca_agencia = " AND cd_banco_ag like ''%"&agencia&"%'' "
	end if
conta = request("conta")
	if conta <> "" Then
		busca_conta = " AND cd_banco_conta like ''%"&conta&"%'' "
	end if
filiacao_m = request("filiacao_m")
	if filiacao_m <> "" Then
		busca_filiacao_m = " AND nm_mae like ''%"&busca_inteligente(filiacao_m)&"%'' "
	end if
filiacao_p = request("filiacao_p")
	if filiacao_p <> "" Then
		busca_filiacao_p = " AND nm_pai like ''%"&busca_inteligente(filiacao_p)&"%'' "
	end if
est_civil = request("est_civil")
	if est_civil <> "" Then
		busca_est_civil = " AND nm_estado_civil like ''%"&busca_inteligente(est_civil)&"%'' "
	end if
conjuge = request("conjuge")
	if conjuge <> "" Then
		busca_conjuge = " AND nm_conjuge like ''%"&busca_inteligente(conjuge)&"%'' "
	end if
nascimento_i = request("nascimento_i")
nascimento_f = request("nascimento_f")
nascimento_tipo = request("nascimento_tipo")
	'if nascimento_i <> "" AND nascimento_tipo <> "" OR  nascimento_f <> "" AND nascimento_tipo <> "" Then
	if nascimento_f <> "" AND nascimento_tipo <> "" Then
		if nascimento_tipo = "1" then
			busca_nascimento = " AND dt_nascimento > ''"&zero(month(nascimento_f))&"/"&zero(day(nascimento_f))&"/"&year(nascimento_f)&"'' "
		elseif nascimento_tipo = "2" then
			busca_nascimento = " AND dt_nascimento < ''"&zero(month(nascimento_f))&"/"&zero(day(nascimento_f))&"/"&year(nascimento_f)&"'' "
		elseif nascimento_tipo = "3" then
			busca_nascimento = " AND dt_nascimento = ''"&zero(month(nascimento_f))&"/"&zero(day(nascimento_f))&"/"&year(nascimento_f)&"'' "
		elseif nascimento_tipo = "4" then
			busca_nascimento = " AND dt_nascimento Between ''"&month(nascimento_i)&"/"&day(nascimento_i)&"/"&year(nascimento_i)&"'' AND ''"&zero(month(nascimento_f))&"/"&zero(day(nascimento_f))&"/"&year(nascimento_f)&"'' "
		end if
	end if
transp = request("transp")
	if transp <> "" Then
		busca_transp = " AND cd_sptrans like ''%"&transp&"%'' "
	end if
refei = request("refei")
	if refei <> "" Then
		busca_refei = " AND cd_vr like ''%"&refei&"%'' "
	end if
alim = request("alim")
	if alim <> "" Then
		busca_alim = " AND cd_vale_alimentacao like ''%"&alim&"%'' "
	end if
medic = request("medic")
	if medic <> "" Then
		busca_medic = " AND cd_assistencia_medica_matricula like ''%"&medic&"%'' "
	end if
odonto = request("odonto")
	if odonto <> "" Then
		busca_odonto = " AND cd_assistencia_odonto_matricula like ''%"&odonto&"%'' "
		
	end if
'----------------------------------------------------------------
outras_variaveis = busca_nome&busca_busca_sexo&busca_endereco&busca_numero&busca_complemento&busca_cep&busca_bairro&busca_cidade&busca_rg&busca_cpf&busca_coren&busca_ctps
outras_variaveis = outras_variaveis&busca_contratos&busca_pis&busca_tit_eleitor&busca_naturalidade&busca_agencia&busca_conta&busca_filiacao_m&busca_filiacao_p&busca_est_civil
outras_variaveis = outras_variaveis&busca_conjuge&busca_nascimento&busca_transp&busca_refei&busca_alim&busca_medic&busca_odonto

var_cargo = busca_cargo
var_unidade = busca_unidade
'****************************************************************
func_ativos = request("func_ativos")

'lista_escolhas = 

escolha_documentos = request("escolha_documentos")
	if escolha_documentos = "" Then 
		escolha_documentos = "none"
	end if
escolha_d_prof = request("escolha_d_prof")
	if escolha_d_prof = "" Then 
		escolha_d_prof = "none"
	end if
escolha_d_pess = request("escolha_d_pess")
	if escolha_d_pess = "" Then 
		escolha_d_pess = "none"
	end if
escolha_beneficios = request("escolha_beneficios")
	if escolha_beneficios = "" Then 
		escolha_beneficios = "none"
	end if

'***** 2º Passo. Ordem dos registros da Busca *******************
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
		ordem_24 = request("ordem_24")
		ordem_25 = request("ordem_25")
%>
<head>
	<title>Listagem de funcionários</title>
</head>

<table align="center" border="1" class="no_print">
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="3" align="center"><span style="font-size=8px;color:red;">funcionarios_lista.asp</span></td></tr>
	<%end if%>
	<tr><td colspan="3" align="center" style="background-color:#000099; color:white; font-size=15; font-weight:bold;">RECIBOS</td></tr>
	<form action="../../empresa.asp" method="post" name="form">
	<input type="hidden" name="tipo" value="recibos">
	
	<!-- *** 3º Passo *** Campos de referencia da Ordem ***-->
	<input type="hidden" name="ordem_res">
	<input type="hidden" name="ordem_total" value="<%=ordem_total%>">
	<input type="hidden" name="ordem_inicial" value="<%if ordem_inicial = "" Then%>1<%else response.write(ordem_inicial) end if%>">
	<tr style="background-color:#0033ff; color:white; font-weight:bold;">
		<td>Data</td>
		<td>Mostrar</td>
		<td>Detalhes</td>
	</tr>
	<tr class="no_print" style="background-color:#ccffff; color:black; font-weight:bold;">
		<td valign="top" align="center">DE:
			<input type="text" name="dia_sel" size="2" maxlength="2" value="<%=dia_sel%>" style="background-color:white;">
			<input type="text" name="mes_sel" size="2" maxlength="2" value="<%=mes_sel%>" style="background-color:white;">
			<input type="text" name="ano_sel" size="4" maxlength="4" value="<%=ano_sel%>" style="background-color:white;">			
			<%'=dia_sel&"/"&mes_sel&"/"&ano_sel&"<br>("&ano_sel&mes_sel&")"%>		
		</td>
		<td>&nbsp;Funcionário</td>
		<td>
			<input type="text" name="nome" size="20" maxlength="100" value="<%=nome%>" style="background-color:white;">
		</td>
	</tr>
	
	
	<tr class="<%'=str_seccao%>" style="display:<%'=escolha_d_prof%>;">
		<td>&nbsp;</td>
		<td>&nbsp;Unidade</td>
		<td><%strsql ="TBL_unidades where cd_hospital > 0 and cd_status = 1"
				Set rs_uni = dbconn.execute(strsql)%>
					<select name="unidade" class="inputs"> 
						<option value="0">Selecione</option>
						<%Do While Not rs_uni.eof
						cod_unidade = rs_uni("cd_codigo")
						nome_unidade = rs_uni("nm_unidade")
						%>
						<option value="<%=cod_unidade%>" <%if abs(unidade) = int(cod_unidade) then response.write("SELECTED")%>><%=nome_unidade%></option>
						<%rs_uni.movenext
						loop
						rs_uni.close
						Set rs_uni = nothing%>
					</select></td>
	</tr>
	
	<tr class="<%'=str_seccao%>" style="display:<%'=escolha_d_prof%>;">
		<td>&nbsp;</td>
		<td>&nbsp;Empresa</td>
		<td><%strsql ="TBL_tipo_contrato"
				Set rs_contr = dbconn.execute(strsql)%>
					<select name="contratos" class="inputs"> 
						<option value="0">Selecione</option>
						<%Do While Not rs_contr.eof
						cod_contrato = rs_contr("cd_codigo")
						nm_contrato = rs_contr("nm_regime_trab")%>
						<option value="<%=cod_contrato%>" <%if abs(contratos) = int(cod_contrato) then response.write("SELECTED")%>><%=nm_contrato%></option>
						<%rs_contr.movenext
						loop
						rs_contr.close
						Set rs_contr = nothing%>
					</select></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;Tipo</td>
		<td><input type="radio" name="modelo" value="1" <%if modelo = "1" Then response.write("CHECKED")%>>Unidade<img src="../../imagens/px.gif" alt="" width="24" height="1" border="0">
			<input type="radio" name="modelo" value="2" <%if modelo = "2" Then response.write("CHECKED")%>>Individual</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;Título</td>
		<td><input type="text" name="titulo" value="<%=titulo%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td valign="top">&nbsp;Texto</td>
		<td><textarea cols="27" rows="5" name="texto"><%=replace(texto,"<br>",vbcrlf)%></textarea></td>
	</tr>
	
		
	<tr class="no_print">
		<td><input type="submit" name="OK" value="Mostrar" style="background-color:gray; color:white;"></td>
		<td colspan="2">&nbsp;</td>
		
	</tr>
	</form>
</table>
<table align="center">
	
	
	<tr>
					<%cor_linha = "#FFFFFF"
					cor = 1
					cabecalho = 0
					conta_linha = 1
					
					if func_ativos = "" OR func_ativos = 1 Then
						'xsql = "up_funcionario_contrato_lista @dt_data='"&ano_sel&mes_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
						'xsql = "up_funcionario_contrato_lista @dt_data='"&ano_sel&mes_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_sel  &"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
						xsql = "up_funcionario_contrato_lista_teste @dt_data='"&ano_sel&mes_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"', @outras_variaveis='"&outras_variaveis&"'"
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
								nm_complemento = rs_func("nm_complemento")
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
								nm_unidade = rs_func("nm_unidade")
								nm_sigla = rs_func("nm_sigla")
									if nm_sigla = "" Then
										nm_sigla = "adm"
									end if
								
								nm_naturalidade = rs_func("nm_naturalidade")
								nm_mae = rs_func("nm_mae")
								nm_pai = rs_func("nm_pai")
								
								nm_banco = rs_func("nm_banco")
								cd_banco_ag = rs_func("cd_banco_ag")
								cd_banco_conta = rs_func("cd_banco_conta")
								nm_estado_civil = rs_func("nm_estado_civil")
								nm_conjuge = rs_func("nm_conjuge")
								dt_nascimento = rs_func("dt_nascimento")
								idade_func = rs_func("idade_func")
									idade_func_anos = int(idade_func/12)
									idade_func_mes = idade_func mod 12
										if idade_func_mes <> 0 then
											idade_func_meses = zero(idade_func_mes)&"m"
										else
											idade_func_meses = ""
										end if
									idade_func = idade_func_anos&"a "&idade_func_meses
								
								cd_sptrans = rs_func("cd_sptrans")
								cd_cmt_bom = rs_func("cd_cmt_bom")
								
								cd_vr = rs_func("cd_vr")
								cd_vale_alimentacao = rs_func("cd_vale_alimentacao")
								cd_assistencia_medica_matricula = rs_func("cd_assistencia_medica_matricula")
								cd_assistencia_odonto_matricula = rs_func("cd_assistencia_odonto_matricula")
								
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
								
								'*** HORÁRIO ***
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
								'xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 AND dt_atualizacao <= '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' order by dt_atualizacao DESC"
								'Set rs_unid = dbconn.execute(xsql)
								'	if not rs_unid.EOF Then
								'		nm_unidade = rs_unid("nm_unidade")
								'		nm_sigla = rs_unid("nm_sigla")
								'	end if
									
								'*** STATUS ***
								xsql = "Select * From View_funcionario_status Where cd_funcionario = "&cd_funcionario&" AND cd_suspenso = 0 order by dt_atualizacao DESC"
								Set rs_status = dbconn.execute(xsql)
									if not rs_status.EOF then
										nm_status = rs_status("nm_status")
									end if
								
								'*** CONTRATO (EMPRESA) ***
								'xsql = "Select * From TBL_tipo_contrato Where cd_codigo = "&cd_contrato&""
								'Set rs_contrato = dbconn.execute(xsql)
								'	if not rs_contrato.EOF then
								'		nm_regime_trab = rs_contrato("nm_regime_trab")
								'	end if
								
								'*** EMPRESA ***
								xsql = "Select * From TBL_empresa Where cd_tipo_contrato = "&cd_contrato&""
								Set rs_empresa = dbconn.execute(xsql)
									if not rs_empresa.EOF then
										nm_empresa = rs_empresa("nm_empresa")
										nm_empresa_endereco = rs_empresa("nm_endereco")
										nm_empresa_numero = rs_empresa("nm_numero")
										nm_empresa_complemento = rs_empresa("nm_complemento")
										nm_empresa_bairro = rs_empresa("nm_bairro")
										nm_empresa_cidade = rs_empresa("nm_cidade")
										nm_empresa_estado = rs_empresa("nm_estado")
										nm_empresa_cep = rs_empresa("nm_cep")
										
									end if
								
								
								mostrar_linha = 1%>
								
								<%'*** Lista os funcionarios ***%>
								<!--tr-->
								<%if int(periodo_sel)=int(demissao) then
									cor_registro = "gray"
								else
									cor_registro = "black"
								end if%>
								<!--tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px; color:<%=cor_registro%>;" bgcolor="<%=cor_linha%>"-->
								<%busca_cargo = instr(1,nm_cargo,cargo,1)
								busca_unidade = instr(1,nm_sigla,unidade,1)
								
								'IF busca_cargo <> 0 and busca_unidade <> 0 then
								'IF busca_cargo <> 0 then
								if unidade <> 0 Then 
									if int(cd_unidade) <> int(unidade) then mostrar_linha = 0 else mostrar_linha = 1
								end if
								
								IF mostrar_linha = 1 Then 'inibe a linha desejada
								
								If modelo = 1 then
									'mostra apenas uma vez
									cabecalho = cabecalho + 1
								elseif modelo > 1 then
									'repete a página para cada nome de funcionário
									cabecalho = 1
									conta_linha = 1
								end if
								
								
									if cabecalho = 1 then%>
										<tr><td>&nbsp;<img src="../../imagens/px.gif" alt="" width="1" height="30" border="0"></td></tr>
										<tr style="font-size:45px;">
											<td align="center"><b><%'=cd_contrato%><%=nm_empresa%></b></td>
										</tr>
										<tr><td>&nbsp;<img src="../../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
										<tr style="font-size:25px;">
											<td align="center"><b><%=titulo%></b></td>
										</tr>
										<tr><td>&nbsp;<img src="../../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
										<tr style="font-size:18px;">
											<%strsql ="TBL_unidades where cd_codigo = "&unidade&""
												Set rs_uni = dbconn.execute(strsql)%>													
														<%if not rs_uni.eof Then
														cod_unidade = rs_uni("cd_codigo")
														nm_unidade = rs_uni("nm_unidade")
														end if
														rs_uni.close
														Set rs_uni = nothing%>
													</select><td><b><%=nm_unidade%></b></td>
										</tr>
										<tr><td>&nbsp;<img src="../../imagens/px.gif" alt="" width="1" height="30" border="0"></td></tr>
									<%end if
									
									'conta_linha = conta_linha + 1%>	
									<tr style="font-size:18px;">
										<td valign="top" onclick="javascript:window.open('empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>','ficha_<%=rs_func("cd_funcionario")%>','width=700,height=700,scrollbars=1');">
										<%=nm_funcionario%></td>
									</tr>
								
								<%if modelo = 2 then%>
									<tr><td><img src="../../imagens/px.gif" alt="" width="1" height="100" border="0"></td></tr>
									<tr><td align="left" style="font-size:25px;"><%=texto%></td></tr>
									<tr><td><img src="../../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
									<tr><td align="left" style="font-size:15px;">Data: <%=dia_sel%>/<%=mes_sel%>/<%=ano_sel%></td></tr>
									<tr><td><img src="../../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
									<tr>
										<td align="center" style="font-size:12px;">_____________________________________________<br>
										Assinatura : <%=nm_funcionario%></td>
									</tr>
									<tr><td><img src="../../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr>
									<tr style="font-size:12px;">
										<td align="center"><hr noshade>
										<%=nm_empresa_endereco%>, <%=nm_empresa_complemento%><br>
										<%=nm_empresa_bairro%>,  <%=nm_empresa_cidade%> - <%=nm_empresa_estado%><br>
										CEP:<%=nm_empresa_cep%></td>
									</tr>
									<tr>
										<td colspan="100%" style="page-break-after:always;" align="right"><%'=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%><%'=linha&"-"&qtd_linhas_pg%><%'=linhas_dif%></td>
									</tr>
								<%end if%>
								
								<%if cabeca_ativo <> cd_contrato then%>
								<tr>
									<td></td>
								</tr>						
								<%end if%>
								<%'if cd_situacao = 2 and pulo_linha_inativos = cd_funcionario then%>
								
								<%cabeca_ativo = cd_contrato
								
							espaco_data = ""
							conta_func_total = conta_func_total + 1	
							cor_admiss = "black"
							cor_demiss = "black"
							
							'modelo = modelo + 1
							END IF '*** Inibe a linha desejada ***
							
							rs_func.movenext
							wend%>	
								<%if modelo = 1 then%>
									<%while conta_linha < 10%>
										<tr><td>&nbsp;<img src="../../imagens/blackdot.gif" alt="" width="1" height="30" border="0"></td></tr>
									<%conta_linha = conta_linha + 1
									wend%>
									
									<tr><td><img src="../../imagens/px.gif" alt="" width="1" height="100" border="0"></td></tr>
									<tr><td align="left" style="font-size:25px;"><%=texto%></td></tr>
									<tr><td><img src="../../imagens/px.gif" alt="" width="1" height="50" border="0"></td></tr><tr><td>&nbsp;</td></tr>
									<tr><td align="left">São Paulo, <%=dia_sel%>/<%=mes_sel%>/<%=ano_sel%><br></td></tr>
									<tr>
										<td align="center">____________________________________________________________________<br>
										Assinatura lider ou responsável<br>
										<br>
										<%=nm_empresa_endereco%>, 2001, sala 11<br>
										Vila Maria - São Paulo, Data<br>
										CEP:02113-017</td>
									</tr>
									<tr><td colspan="100"><hr></td></tr>
								<%end if%>
					<%'cabeca = ""
					
					lista_empresa = ""
'*******************************************************%>
		
		
	
		
	<tr>	
		<td><img src="../../imagens/px.gif" alt="" width="600" height="1" border="0"></td>
	</tr>
</table>


</body>
</html>
