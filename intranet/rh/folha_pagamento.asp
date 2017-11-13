<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<script language="javascript">
<!-- //1º Passo. Mecânica da busca
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
	
	document.getElementsByName(objeto)[1].value= document.form.ordem_inicial.value+'º';
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


'***** 2º Passo. Ordem dos registros da Busca *******************
ordem_total = request("ordem_total")
	
	
	ver_string_ordem = instr(ordem_total,",")
		if ver_string_ordem = 1 then
			ordem_total = mid(ordem_total,2,len(ordem_toral))
		end if
	
			if ordem_total <> "" Then
				ordem_funcionarios = ordem_total
			else
				'ordem_funcionarios = "cd_contrato,nm_nome,nm_sigla"
				ordem_funcionarios = "cd_contrato,nm_nome"
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
		
	'xsql = "Select * From TBL_inss Where dt_atualizacao between '"&mes_sel&"/1/"&ano_sel&"' AND '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
	xsql = "Select * From TBL_inss Where dt_atualizacao <= '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
		Set rs_inss = dbconn.execute(xsql)
		if not rs_inss.EOF then
			faixa_1 = rs_inss("pri_faixa")
			faixa_2 = rs_inss("seg_faixa")
			faixa_3 = rs_inss("ter_faixa")
		end if
%>
<head>
	<title>Listagem de funcionários</title>
</head>

<body><br>

<table align="center" border="1" class="no_print">
	<tr><td colspan="100%" align="center" style="background-color:#000099; color:white; font-size=15; font-weight:bold;">Folha de pagamento</td></tr>
	<form action="../../rh.asp" method="post" name="form">
	<input type="hidden" name="tipo" value="folha">
	<input type="hidden" name="empresa" value="cmi">
	
	<!-- *** 3º Passo *** Campos de referencia da Ordem ***-->
	<input type="hidden" name="ordem_res">
	<input type="hidden" name="ordem_total" value="<%=ordem_total%>">
	<input type="hidden" name="ordem_inicial" value="<%if ordem_inicial = "" Then%>1<%else response.write(ordem_inicial) end if%>">
	<tr style="background-color:#0033ff; color:white; font-weight:bold;">
		<td>Data</td>
		<td>Mostrar</td>
		<td>Detalhes</td>
		<td><a href="javascript:void();" onClick="limpa_ordem()" title="Limpa a ordem">Ordem</a></td>
	</tr>
	<tr class="no_print" style="background-color:#ccffff; color:black; font-weight:bold;">
		<td valign="top" align="center">DE:
			<input type="text" name="dia_sel" size="2" maxlength="2" value="<%=dia_sel%>" style="background-color:white;">
			<input type="text" name="mes_sel" size="2" maxlength="2" value="<%=mes_sel%>" style="background-color:white;">
			<input type="text" name="ano_sel" size="4" maxlength="4" value="<%=ano_sel%>" style="background-color:white;">			
			<%'=dia_sel&"/"&mes_sel&"/"&ano_sel&"<br>("&ano_sel&mes_sel&")"%>		
		</td>
		<td>
			&nbsp; Nome do funcionário
		</td>
		<td>
			<input type="text" name="nome" size="20" maxlength="100" value="<%=nome%>" style="background-color:white;">
		</td>
		<td><%ordem_var = ordem_1
		ordem_num="ordem_1"
		ordem_campo="nm_nome"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
	</tr>
	
	<tr class="no_print">
		<td><input type="submit" name="OK" value="Mostrar" style="background-color:gray; color:white;"></td>
		<td colspan="2">&nbsp;<a href="empresa.asp?tipo=novo"><img src="../../imagens/ic_novo.gif" alt="" width="10" height="12" border="0">Novo colaborador</a></td>
		<td><a href="javascript:void();" onClick="limpa_ordem()" title="Limpa a ordem">limpa ordem</a></td>
	</tr>
	</form>
</table>
<%if mes_sel <> "" AND ano_sel <> "" Then%>
<table align="center">
	<tr>
					<%	'*** Aj. Custo ***
						xsql = "Select top 1 * From VIEW_funcionario_valores_padroes WHERE cd_tipo=5 and dt_atualizacao<='"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
						Set rs = dbconn.execute(xsql)
							if not rs.EOF then
								valor_ajuda_custo_padrao = rs("nr_valor")
								ajuda_custo_atualizacao = rs("dt_atualizacao")
								ajuda_custo_padrao = (year(ajuda_custo_atualizacao)&zero(month(ajuda_custo_atualizacao))&zero(day(ajuda_custo_atualizacao)))
							end if
						
						'*** Insalubridade ***
						xsql = "Select top 1 * From VIEW_funcionario_valores_padroes WHERE cd_tipo=7 and dt_atualizacao<='"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
						Set rs = dbconn.execute(xsql)
							if not rs.EOF then
								valor_insalubr_padrao = rs("nr_valor")
								insalubr_atualizacao = rs("dt_atualizacao")
								insalubr_padrao = (year(insalubr_atualizacao)&zero(month(insalubr_atualizacao))&zero(day(insalubr_atualizacao)))
							end if
						
						'*** Auxilio creche ***
						xsql = "Select top 1 * From VIEW_funcionario_valores_padroes WHERE cd_tipo=6 and dt_atualizacao<='"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
						Set rs = dbconn.execute(xsql)
							if not rs.EOF then
								valor_auxilio_creche_padrao = rs("nr_valor")
								auxilio_creche_atualizacao = rs("dt_atualizacao")
								auxilio_creche_padrao = (year(auxilio_creche_atualizacao)&zero(month(auxilio_creche_atualizacao))&zero(day(auxilio_creche_atualizacao)))
							end if
					
					cor_linha = "#FFFFFF"
					cor = 1
					
					if func_ativos = "" OR func_ativos = 1 Then
						xsql = "up_funcionario_rh_lista @dt_data='"&ano_sel&mes_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
							Set rs_func = dbconn.execute(xsql)
							response.write(func_ativos)
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
								
								dt_nascimento = rs_func("dt_nascimento")
									idade_func = datediff("m",dt_nascimento,data_hoje)
								
								cd_plano_saude_conj = rs_func("cd_plano_saude_conj")
									if cd_plano_saude_conj = 1 then
										dt_conjuge_nasc = rs_func("dt_conjuge_nasc")
											idade_conj = datediff("m",dt_conjuge_nasc,data_hoje)
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
								
								'tempo_casa = rs_func("tempo_casa")
								
								cd_unidade = rs_func("cd_unidade")
								nm_unidade = rs_func("nm_unidade")
								nm_sigla = rs_func("nm_sigla")
								
								nr_salario = rs_func("nr_salario")
								
								idade_depend = rs_func("idade_depend")
								
								nm_naturalidade = rs_func("nm_naturalidade")
								
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
									if x = "" then%>
										<tr class="no_print">
											<td colspan="100%">&nbsp;</td>										
										</tr>
										<tr bgcolor=#c0c0c0>
											<td align="center">Qtd.</td>
											<td><b>Unidade</b></td>
											<td><b>Mat.</b></td>
											<td><b>Funcionário</b></td>
											<td><b>Salário</b></td>
											<td><b>Aj. Custo</b></td>
											<td><b>Insalubridade</b></td>
											<td><b>Ad. noturno</b></td>
											<td><b>Dias noturno</b></td>
											<td><b>Aux. Creche</b></td>
											<td><b>Qtd h extra</b></td>
											<td><b>Horas Extras</b></td>
											<td><b>DSR h extra</b></td>
											<td><b>DSR ad. Noturno</b></td>
											<td><b>V. refeição</b></td>
											<td><b>V. transp</b></td>
											<td><b>Base INSS</b></td>
											<td><b>INSS</b></td>
											<td><b>Conv. Médico</b></td>
											<td><b>C. Básica</b></td>
											<td><b>IRRF</b></td>
											<td><b>Desc. Div</b></td>
											
											
											<td><b>Valor Líquido</b></td>
										</tr>
									<%x = 1
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
									<td valign="top"><%=nm_sigla%></td>
									<td align="right" valign="top"><%=cd_matricula%></td>
									<td valign="top" onclick="javascript:window.open('empresa/funcionario/funcionarios_ficha.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');"><%=nm_funcionario%> <%=cd_funcionario%></td>
									<td><%=nr_salario%></td>									
									<td align="right">
									<%'**********************************************************************
									'***                        Ajuda de Custo                            ***
									'************************************************************************
									xsql = "Select * From VIEW_funcionario_valores Where cd_funcionario='"&cd_funcionario&"' AND cd_tipo=5 AND dt_atualizacao<='"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
										Set rs_valor = dbconn.execute(xsql)
										if not rs_valor.EOF then
											dt_aj_custo_indv = rs_valor("dt_atualizacao")
											aj_custo_indv = int(year(dt_aj_custo_indv)&zero(month(dt_aj_custo_indv))&zero(day(dt_aj_custo_indv)))
											ajuda_custo_valor = rs_valor("nr_valor")
										end if
										
										if int(aj_custo_indv) >  int(ajuda_custo_padrao) Then
											ajuda_custo = ajuda_custo_valor
										else
											ajuda_custo = valor_ajuda_custo_padrao
										end if%>									
									<%=ajuda_custo%></td>
									<td align="right">
									<%'**********************************************************************
									'***                         Insalubridade                            ***
									'************************************************************************
									xsql = "Select * From VIEW_funcionario_valores Where cd_funcionario='"&cd_funcionario&"' AND cd_tipo=7 AND dt_atualizacao<='"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
										Set rs_valor = dbconn.execute(xsql)
										if not rs_valor.EOF then
											dt_insalubr_indv = rs_valor("dt_atualizacao")
											insalubr_indv = int(year(dt_insalubr_indv)&zero(month(dt_insalubr_indv))&zero(day(dt_insalubr_indv)))
											insalubr_valor = rs_valor("nr_valor")
										end if
										
										if int(insalubr_indv) >  int(insalubr_padrao) Then
											insalubridade = insalubr_valor
										else
											insalubridade = valor_insalubr_padrao
										end if%>									
									<%=insalubridade%></td>
									<td align="right">
									<%'**********************************************************************
									'***                         Adicional Noturno                        ***
									'************************************************************************
									xsql = "Select * From VIEW_funcionario_valores Where cd_funcionario='"&cd_funcionario&"' AND cd_tipo=8 AND dt_atualizacao between '"&mes_sel&"/1/"&ano_sel&"' AND '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
										Set rs_valor = dbconn.execute(xsql)
										if not rs_valor.EOF then
											dt_ad_noturno_indv = rs_valor("dt_atualizacao")
											ad_noturno_indv = int(year(dt_ad_noturno_indv)&zero(month(dt_ad_noturno_indv))&zero(day(dt_ad_noturno_indv)))
											qtd_dia_noturno = rs_valor("nr_valor")
										end if
										ad_noturno = (((nr_salario*0.40)/180)*7)*qtd_dia_noturno
											if not ad_noturno = 0 then
												response.write(formatnumber(ad_noturno,2))
												ad_noturno_dsr = formatnumber((ad_noturno/25)*6.25,2)
											else
												ad_noturno = 0
											end if%></td>
									<td align="right">
										<%if not qtd_dia_noturno = 0 then
											qtd_dia_noturno = formatnumber(qtd_dia_noturno,0)
										else
											qtd_dia_noturno = "0"
										end if%>
										<%=qtd_dia_noturno%></td>
									<td align="right">
									<%if idade_depend > 0 Then%>
										<%=formatnumber(idade_depend*valor_auxilio_creche_padrao,2)%>
									<%else%>0<%end if%></td>
									<td align="right">
									<%'************************************************************************
									'***                         Horas Extras		                        ***
									'**************************************************************************
									xsql = "Select * From VIEW_funcionario_valores Where cd_funcionario='"&cd_funcionario&"' AND cd_tipo=11 AND dt_atualizacao between '"&mes_sel&"/1/"&ano_sel&"' AND '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
										Set rs_valor = dbconn.execute(xsql)
										if not rs_valor.EOF then
											dt_horas_indv = rs_valor("dt_atualizacao")
											horas_indv = int(year(dt_ad_noturno_indv)&zero(month(dt_ad_noturno_indv))&zero(day(dt_ad_noturno_indv)))
											qtd_horas = rs_valor("nr_valor")
										else
											qtd_horas = 0
										end if%>
									<%=formatnumber(qtd_horas,0)%></td>
									<td align="right">
										<%if qtd_horas > 0 AND nr_salario > 0 then
											horas_extras = ((nr_salario/180)*1.9)*qtd_horas
												 response.write(formatnumber(horas_extras,2))
											horas_extras_dsr = (horas_extras/25)*6.25
												 horas_extras_dsr = (formatnumber(horas_extras_dsr,2))
										else
											horas_extras = 0
										end if%>
									</td>
									<td align="right"><%=horas_extras_dsr%></td>
									<td align="right"><%=ad_noturno_dsr%></td>
									<td align="right"><%xsql = "Select * From VIEW_funcionario_horario Where cd_funcionario='"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/1/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
										Set rs_hora = dbconn.execute(xsql)
											if not rs_hora.EOF then
												expediente = rs_hora("expediente")
													if abs(expediente) > 360 then
														xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=2 AND dt_atualizacao <= '"&mes_sel&"/1/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
															Set rs_refeic = dbconn.execute(xsql)
															if not rs_refeic.EOF then
																valor_refeicao = rs_refeic("nr_valor")
															end if
													end if
												response.write(formatnumber(valor_refeicao))
											end if%>
									</td>
									<td align="right"><%if nr_salario <> "" Then
											v_trasnsp = nr_salario*0.06
											response.write(formatnumber(v_trasnsp))
										else
											response.write("0")
										end if%></td>
									<td align="right"><%base_inss = abs(nr_salario) + abs(insalubridade) + abs(ad_noturno) +abs( horas_extras) + abs(horas_extras_dsr) + abs(ad_noturno_dsr)
									if not base_inss = 0 then response.write(formatnumber(base_inss,2))%></td>
									<td align="right"><%if base_inss <= abs(faixa_1) then
															inss = base_inss * 0.08
															response.write(formatnumber(inss,2))
														elseif base_inss <= abs(faixa_2) then
															inss = base_inss * 0.09
															response.write(formatnumber(inss,2))
														elseif base_inss <= abs(faixa_3) then
															inss = base_inss * 0.11
															response.write(formatnumber(inss,2))
														end if%>														
										</td>
									<td align="right">
										<%'*** Idade do titular do convênio médico ***
											idade_func = int(idade_func/12)
												xsql = "Select top 1 * From VIEW_planosaude_valores Where nr_idade >= '"&idade_func&"' AND cd_categoria=1 AND dt_atualizacao <= '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' ORDER BY nr_idade, dt_atualizacao desc"
													Set rs_saude = dbconn.execute(xsql)
													if not rs_saude.EOF then
														valor_saude = rs_saude("nr_valor")
														'response.write(formatnumber(cesta_b,2))
														'response.write(idade_func&" : "&(formatnumber(valor_saude*.20)))
													end if
											'response.write(idade_func)
										'*** Verifica se o conjuge optou pelo convênio médico e a idade ***
											if cd_plano_saude_conj = 1 then
											idade_conj = int(idade_conj/12)
												xsql = "Select top 1 * From VIEW_planosaude_valores Where nr_idade >= '"&idade_conj&"' AND cd_categoria=1 AND dt_atualizacao <= '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' ORDER BY nr_idade, dt_atualizacao desc"
													Set rs_saude = dbconn.execute(xsql)
													if not rs_saude.EOF then
														valor_saude_conj = rs_saude("nr_valor")
														'response.write(formatnumber(cesta_b,2))
														'response.write(idade_conj&" : "&(formatnumber(valor_saude_conj*.20)))
													end if
													'response.write(idade_conj)
											end if
										'*** Verifica se há algum dependente, sexo  e idade ***
										'xsql = "Select * From VIEW_funcionario_dependentes Where cd_funcionario='"&cd_funcionario&"' AND cd_parentesco=1 AND cd_sexo=1 AND nr_idade >= '"&idade_conj&"' AND cd_categoria=1 ORDER BY nr_idade Desc "
										'xsql = "SELECT * FROM View_funcionario_dependentes WHERE cd_funcionario='"&cd_funcionario&"' AND cd_parentesco=1 AND idade<=128"
										xsql = "SELECT COUNT(nm_nome) AS n_dependentes FROM View_funcionario_dependentes WHERE cd_funcionario='"&cd_funcionario&"' AND cd_parentesco=1 AND idade<=128 AND cd_sexo = 1"
											Set rs_depend = dbconn.execute(xsql)
												while not rs_depend.EOF
													response.write(rs_depend("n_dependentes"))
												rs_depend.movenext
												wend
										%>
									
									</td>
									<td align="right">
										<%xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=10 AND dt_atualizacao <= '"&mes_sel&"/1/"&ano_sel&"' ORDER BY dt_atualizacao DESC"
											Set rs_refeic = dbconn.execute(xsql)
											if not rs_refeic.EOF then
												cesta_b = rs_refeic("nr_valor")
												response.write(formatnumber(cesta_b,2))
											end if%>
									</td>
									<td align="right"></td>
									<td align="right"></td>
									
									<%t_func_receita = abs(nr_salario) + abs(ajuda_custo) + abs(insalubridade) + abs(ad_noturno) + abs(horas_extras)+ abs(horas_extras_dsr) + abs(ad_noturno_dsr)
									t_func_despesa = abs(v_refeicao) + abs(v_trasnsp) + abs(inss) + abs(conv_med) + abs(cesta_b) + abs(sindicato) + abs(desc_div)
										total_func_liq = t_func_receita - t_func_despesa%>
									<td align="right"><%if not total_func_liq = "" then%><%=formatnumber(abs(total_func_liq),2)%><%end if%></td>
									
									
								<td><img src="../../imagens/ic_editar.gif" alt="Editar" width="13" height="14" border="0" onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');"></td>
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
								
							nr_salario = ""
								ajuda_custo = ""
								aj_custo_indv = 0
									insalubridade = ""
									insalubr_indv = 0
										ad_noturno = ""
										qtd_dia_noturno = 0
										ad_noturno_dsr = 0
											valor_refeicao = 0
												qtd_horas = ""
												horas_extras = ""
												horas_indv = 0
												horas_extras_dsr = 0
							
							total_func_liq = ""
							
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
		
		'*** conta os funcionários de cada contrato ***
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
		<%end if%>
		<!--tr>
			<td colspan="100%">
				SELECT * FROM VIEW_funcionario_contrato_lista <br>
				WHERE dt_admissao <= <%=ano_sel&mes_sel%> AND dt_demissao >= <%=ano_sel&mes_sel%>  OR dt_admissao <= <%=ano_sel&mes_sel%>  AND dt_demissao IS NULL <br>
				ORDER BY <%=ordem_funcionarios%>	
			</td>
		</tr-->
		<tr>
			<td><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
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
<%end if%>

</body>
</html>
