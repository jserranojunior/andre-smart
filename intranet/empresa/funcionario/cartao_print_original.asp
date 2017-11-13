<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
	<title>Untitled</title>
<%cd_unidade = request("cd_unidade")
qtd_paginas = request("qtd_paginas")
qtd_etiquetas = request("qtd_etiquetas")
mensagem = request("mensagem")
tipo_print = request("tipo_print")
espacamento = 1
mostra_mes_inicial = request("mostra_mes_inicial")
posicao_topo_cartao = request("posicao_topo_cartao")

dt_inicio_sys = "01/01/2008"

func_sel = request("func_sel")
	func_ver = func_sel
	if func_sel = 0 then
		func_sel = "todos"
	end if

strcd_unidade = request("cd_unidade")
	if strcd_unidade = "" Then	strcd_unidade = 0


dia_sel = zero(request("dia_sel"))
mes_sel = zero(request("mes_sel"))
ano_sel = request("ano_sel")
data_atual = dia_sel&"/"&mes_sel&"/"&ano_sel

mes_print = request("mes_print")
ano_print = request("ano_print")
ultmo_dia_mes_print = ultimodiames(mes_print,ano_print)

if mes_sel = "" AND ano_sel = "" Then
	dia_sel = ultimodiames(month(now),year(now))
end if


dia_prod = zero(day(now))
mes_prod = zero(month(now))
ano_prod = year(now)
data_prod = dia_prod&"/"&mes_prod&"/"&ano_prod

if year(dt_inicio_sys) = int(ano_prod) then
	ano_sel = year(dt_inicio_sys)
end if
%>
<style type="text/css">
<!--
.relatorio_titulo{color: #000000;font-family: verdana;font-size:12px;font-weight:bold}
.relatorio_campos{color: #000000;font-family: verdana;font-size:7px;}
.relatorio_unidade{color: #000000;font-family: verdana;font-size:9px;font-weight:bold}
.relatorio_protocolo{color: #000000;font-family: verdana;font-size:13px;font-weight:bold}
-->
</style>
<!--onClick="window.print()-->
<SCRIPT LANGUAGE="javascript">
function Jsconfirm_print()
{
  if (confirm ("Tem certeza que deseja imprimir?"))
	  {
		//document.location.href('acoes/patrimonio_acao.asp?cd_apaga='+cod1+'&cd_tipo=patrimonio');
		
	  }
}

function limpa_dia()
	{document.Ano.dia_sel.value=''
	}
</SCRIPT>

</head>

<body><br class="no_print">

<%'Passo 1
'*********************************************************
'*				      1ª Parte	 					  	 *
'*********************************************************
'* 1 - Seleção da data *
'***********************%>

<%'if mes_sel = "" AND ano_sel = "" AND cd_codigo = "" Then
if mes_sel = "" AND ano_sel = "" Then%>
				<table border="0" align="center" style="border:1 solid black;" class="no_print">
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
					    <td colspan="100%" align="center">Se necessário, selecione<br>um período para a listagem:</td>
					</tr>
					<!--tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr-->
					<form action="<%=nome_arquivo%>" name="Ano" method="get">
					<input type="hidden" name="tipo" value="<%=tipo%>">
					<!--input type="hidden" name="cd_codigo" value="<%=cd_codigo%>"-->
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">					
					<tr>
						<!--td align="center">DIA</td-->
						<td align="center" colspan="2" style="color:gray;"><input type="text" name="dia_sel" size="2" maxlength="2" value="<%=zero(dia_sel)%>" onMouseup="limpa_dia()">
						<!--Obs.: O dia deve ser referente à listagem desejada.<td>
					</tr>
					<tr>
						<td align="center">MÊS</td>
						<td align="left"-->
						<%ano_inicio = year(dt_inicio_sys)
						  mes_inicio = month(dt_inicio_sys)
								if int(ano_sel) > int(ano_inicio) then
									mes_inicio = "1"
									mes_alt = mes_prod + 2
								else
									mes_alt = 13
								end if%>
								
						<select name="mes_sel">
								<%for i = 1 to 12
								'do while not mes_ano = 13
								'response.write(mes_inicio&"-")
								'response.write(mes_prod&".")%>
								<option value="<%=i%>" <%if int(month(now)) = i then%><%response.write("selected")%><%end if%>><%=mes_selecionado(i)%></option>
								<%next%>			
						</select>						
						<!--/td>
					</tr>
					<tr>
						<td align="center">ANO</td>
						<td align="left"-->
						<select name="ano_sel">
								<%ano_inicio = year(dt_inicio_sys)
								'response.write(ano_inicio)
								do while not ano_inicio = ano_prod + 1%>
								<option value="<%=ano_inicio%>" <%if int(year(now))=ano_prod then%><%response.write("selected")%><%end if%>><%=ano_inicio%></option>
								<%ano_inicio = ano_inicio + 1
								loop%>			
						</select>
						</td>
					</tr>
					<tr><td colspan="2"><hr noshade width="100%"></td></tr>
					<tr bgcolor="silver">
						<td align="center" colspan="2" style="font-size:14px;"><b>IMPRIMIR CARTÕES DE:</b></td>
					</tr>
					<tr><td>&nbsp;</td></tr></tr>
						<td align="center" colspan="2"><select name="mes_print">
								<%for i = 1 to 12
								'do while not mes_ano = 13
								'response.write(mes_inicio&"-")
								'response.write(mes_prod&".")
									mes_ant = i
									mes_post = i + 1
										if mes_post > 12 then
											mes_post = 1
										end if%>
								<option value="<%=i%>" <%if int(month(now)) = i then%><%response.write("selected")%><%end if%>><%=mes_selecionado(mes_ant)&"/"&mes_selecionado(mes_post)%></option>
								<%next%>			
						</select>
						<select name="ano_print">
								<%ano_inicio = year(dt_inicio_sys)
								'response.write(ano_inicio)
								do while not ano_inicio = ano_prod + 1%>
								<option value="<%=ano_inicio%>" <%if int(year(now))=ano_prod then%><%response.write("selected")%><%end if%>><%=ano_inicio%></option>
								<%ano_inicio = ano_inicio + 1
								loop%>			
						</select></td>
					</tr>					
					<tr>
						<td align="right"><br><br>						
						<input type="submit" value="OK">
						</td>
						<td align="center"></td>
					</tr></form>					
				</table>
<%'*******************************
'* 1.3 - Seleção de Funcionários *
'*********************************
Else
'elseif mes_sel <> "" AND ano_sel <> "" AND func_sel = "" Then%>
				<table border="1" align="center" class="no_print">
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
						<td colspan="100%" align="center" style="background-color:gray; color:white;">&nbsp;<br>IMPRESSÃO DE CARTÕES DE PONTO<br>&nbsp;</td>
					</tr>
					<form action="<%=nome_arquivo%>" name="mes" method="get">
					<input type="hidden" name="tipo" value="<%=tipo%>">
					<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">
					
					<input type="hidden" name="ano_sel" value="<%=ano_sel%>">
					<input type="hidden" name="mes_sel" value="<%=int(mes_sel)%>">
					<input type="hidden" name="dia_sel" value="<%=int(dia_sel)%>">
					
					<input type="hidden" name="mes_print" value="<%=int(mes_print)%>">
					<input type="hidden" name="ano_print" value="<%=int(ano_print)%>">
					<tr>
					    <td colspan="100%" align="center">Selecione o colaborador abaixo:</td>
					</tr>
					<tr>
						<td align="center" colspan="100%">
						<%hoje = month(now)&"/"&day(now)&"/"&year(now)
						
						xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao > '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND cd_contrato = 1 OR dt_admissao <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao is NULL AND cd_contrato = 1 ORDER BY nm_nome"
						Set rs = dbconn.execute(xsql)%>
						
						
						<select name="func_sel">
							<option value="0">Todos ativos</option>
							<%while not rs.EOF
								cd_funcionario = rs("cd_funcionario")
								nome = rs("nm_nome")' &" "&rs("nm_sobrenome")
							%>							
							<option value="<%=cd_funcionario%>" <%if int(func_ver)=int(cd_funcionario) then response.write("selected") end if%>" onmousedown="submit()"><%=nome%></option>
							<%rs.movenext
							wend%>									
						</select><br>
						<br>
						<%strsql ="TBL_unidades where cd_hospital > 0 and cd_status = 1"
								  	Set rs_uni = dbconn.execute(strsql)%>
									<select name="cd_unidade" class="inputs"> 
										<option value="">Selecione a unidade</option>
										<%Do While Not rs_uni.eof%>
										<option value="<%=rs_uni("cd_codigo")%>" <%if int(strcd_unidade) = rs_uni("cd_codigo") then response.write("SELECTED")%>><%=rs_uni("nm_Unidade")%></option>
										<%rs_uni.movenext
										loop
										rs_uni.close
										Set rs_uni = nothing%>
									</select><br>				
						
						
						
						<input type="checkbox" name="mostra_mes_inicial" value="1" <%if mostra_mes_inicial = 1 then%>Checked<%end if%>>Mostra apenas o mês inicial no cartão<br><br>
											
						<input type="radio" name="posicao_topo_cartao" value="0" <%if posicao_topo_cartao = "" OR posicao_topo_cartao = "" then%>Checked<%end if%>>Posição normal<br>
						<input type="radio" name="posicao_topo_cartao" value="1" <%if posicao_topo_cartao = 1 then%>Checked<%end if%>>Mais para cima<br>
						<input type="radio" name="posicao_topo_cartao" value="2" <%if posicao_topo_cartao = 2 then%>Checked<%end if%>>Mais para baixo
						<br>
						<br>						
						<input type="submit" value="OK">	&nbsp; <input type="button" value="Limpa" onclick="location.replace('empresa.asp?tipo=cartao_ponto');">
						
						&nbsp;&nbsp;&nbsp;&nbsp;
						<img src="../../imagens/ic_print.gif" alt="" width="24" height="24" border="0" onClick="window.print();">&nbsp;&nbsp;&nbsp;<img src="../../imagens/ic_print_view.gif" alt="" width="24" height="26" border="0" onclick="visualizarImpressao();"></td>
					</tr>
					</form>
					
					<tr>
					    <td colspan="100%" align="center">Período selecionado, referente a:<br>&nbsp;<%'=zero(dia_sel)%>
						<%'=ucase(mes_selecionado(int(mes_sel)))
						'If mostra_mes_inicial = "" Then
						'	if mes_sel < 12 then
						'		mostra_meses = ucase(mes_selecionado(int(mes_sel)))&"/"& ucase(mes_selecionado(int(mes_sel+1)))
						'	elseif mes_sel = 12 then
						'		mostra_meses = ucase(mes_selecionado(int(mes_sel)))&"/"& ucase(mes_selecionado(int(1)))
						'	end if
						'else
						'	mostra_meses = ucase(mes_selecionado(int(mes_sel)))
						'end if
						mes_print_ant = mes_print + 1 
							if mes_print_ant > 12 then
								mes_print_ant = 1
							end if
						mostra_meses = mes_selecionado(mes_print)&"/"&mes_selecionado(mes_print_ant)
						%> 
						
						<b><%=mostra_meses%>
						
						/ <%=ano_sel%></b><br>&nbsp;</td>
					</tr>					
				</table>
<%'***********************
'* 1.4 - Corpo da página *
'*************************
'elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo = "" Then%>
<!--tr><td><%'=data_atual%></td></tr-->
<%'else%>

<%quantidade = 0
num = 1
conta_funcionarios = 1%>


	 
	
		<!--table style="border:1px solid red; width:<%=larg_tabs%>px;" id="ok__print_">
		
			<tr>				
				<td><img src="../../imagens/blackdot.gif" alt="" width="210" height="1" border="0"></td>
				<td style="border:1px solid blue;"-->				
				<%
				'response.write(func_sel)
				
				If func_sel = "todos" Then
					'xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao > '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND cd_contrato = 1 OR dt_admissao <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao is NULL AND cd_contrato = 1 ORDER BY nm_nome"
					'xsql = "up_funcionario_contrato_lista @dt_data='"&mes_sel&"/"&dia_sel&"/"&ano_sel&"', @dt_atualizacao='"&mes_print&"/"&ultmo_dia_mes_print&"/"&ano_print&"', @ordem_funcionarios=nm_nome"
					xsql = "up_funcionario_contrato_lista_teste @dt_data='"&ano_sel&zero(mes_sel)&"', @dt_atualizacao='"&mes_print&"/"&ultmo_dia_mes_print&"/"&ano_print&"', @ordem_funcionarios='nm_nome', @outras_variaveis=''"
				else
					xsql = "Select top 1 * From View_funcionario_contrato_lista Where cd_funcionario = "&func_sel&""
					unico = 1
				end if
				Set rs = dbconn.execute(xsql)
	
	
	Do While Not rs.eof 
	cd_funcionario = rs("cd_funcionario")
	nm_funcionario = rs("nm_nome")
	cd_matricula = rs("cd_matricula")
	
	'if func_sel <> "0" then
	'	strcd_unidade = 0
	'end if
	
		if strcd_unidade <> 0 then
			cd_unidade = rs("cd_unidade")
		else
			cd_unidade = 0
		end if
	
	'nm_sigla = rs("nm_sigla")
	
	'**********************************************************
	'*** DEFINE A APARENCIA EM LARGURA DE TODO O FORMULÁRIO ***
	'**********************************************************
	larg_tabs = 325 + 210
	'**************
	pad_item = round((larg_tabs/4)*0.75,0)
	pad_espaco = round((larg_tabs/4)*0.25,0)
	altura_linha = "24"
	borda_linhas = "0px"
	
	espaco_top = 7'22
		if posicao_topo_cartao = 1 then
			espaco_top = espaco_top - 5
		elseif posicao_topo_cartao =  2 then
			espaco_top = espaco_top + 5
		end if
	'*********************************************************** // CORPO DO CARTÃO \\ ************************************************%>
				<%if int(cd_unidade) = int(strcd_unidade) OR int(cd_unidade) = "" then%>
					<table style="border:0px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; font-size:11px; text-decoration:bold;" class="ok_print">
						<tr><td colspan="100%" align="center"><img src="../../imagens/px.gif" alt="" width="1" height="<%=espaco_top%>" border="0" align="middle"></td></tr>
						<tr id="oks_print">
							<td rowspan="10"><img src="../../imagens/px.gif" alt="" width="215" height="1" border="0"></td>
							<td style="height:<%=altura_linha+1%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2">&nbsp;<!--Nº--></td>
							<td style="height:<%=altura_linha+1%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="4"> &nbsp; &nbsp; CMI CIRURGICA LTDA</td>		
						</tr>							
						<tr id="oks_print">
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="4"> &nbsp; &nbsp; 05.204.601/0001-30</td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2">&nbsp;<!--Atividade Econ.--></td>		
						</tr>							
						<tr>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="6"> &nbsp; &nbsp; <a href="funcionarios_cadastro.asp?tipo=cadastro&pag=0&cod=<%=cd_funcionario%>&busca=" target="_blank"><%'=cd_funcionario&"."%> <%=nm_funcionario%></a></td>								
						</tr>
						<tr id="oks_print">
						<%strsql = "SELECT * FROM VIEW_funcionario_cargo WHERE cd_funcionario = '"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/"&UltimoDiaMes(mes_sel,ano_sel)&"/"&ano_sel&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
						Set rs_cargo = dbconn.execute(strsql)
							if not rs_cargo.EOF then
								nm_cargo = rs_cargo("nm_cargo")
							end if%>						
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2"> &nbsp; &nbsp; <%=cd_matricula%></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2">&nbsp;<!--Nº CTPS--></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2" align="center">&nbsp; <%=nm_cargo%></td>	
						</tr>
						<tr id="oks_print">
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="5">&nbsp;&nbsp;&nbsp;<%'=cd_unidade%><%'=Ucase(local_trabalho)%></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2">&nbsp;</td>
						</tr>
						<tr id="oks_print">							
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" align="center" colspan="5">&nbsp;<%=mostra_meses%><%'=ucase(mes_selecionado(int(mes_sel)))%></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" align="center" colspan="2">&nbsp;<%=ano_sel%></td>	
						</tr>
						<tr id="oks_print">
						<%strsql = "SELECT * FROM View_funcionario_horario WHERE cd_funcionario = '"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/"&UltimoDiaMes(mes_sel,ano_sel)&"/"&ano_sel&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
						Set rs_horario = dbconn.execute(strsql)
							if not rs_horario.EOF then
								hr_entrada = rs_horario("hr_entrada")
								hr_saida = rs_horario("hr_saida")
								nm_intervalo = rs_horario("nm_intervalo")
							end if
								if hr_entrada = "" Then hr_entrada = "00:00" end if
								if hr_saida = "" Then hr_saida = "00:00" end if%>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom">&nbsp;</td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" align="center"><%=zero(hour(hr_entrada))%>:<%=zero(minute(hr_entrada))%> &nbsp;</td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" align="center">&nbsp;&nbsp;<%=nm_intervalo%></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2"> &nbsp; &nbsp; &nbsp;<%=zero(hour(hr_saida))%>:<%=zero(minute(hr_saida))%></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" align="center">&nbsp;</td>	
						</tr>
						<!--tr id="oks_print"-->
						<tr>
							<td><img src="../../imagens/px.gif" alt="" width="33" height="1" border="0"></td>
							<td><img src="../../imagens/px.gif" alt="" width="52" height="1" border="0"></td>
							<td><img src="../../imagens/px.gif" alt="" width="84" height="1" border="0"></td>								
							<td><img src="../../imagens/px.gif" alt="" width="10" height="1" border="0"></td>
							<td><img src="../../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
							<td><img src="../../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
						</tr>
						<tr id="osk_print"><td>&nbsp;</td></tr>
					</table>
				<%end if%>
<%	'*********************************************************** // FIM DO CORPO DO CARTÂO \\ ************************************************%>
				<%if int(cd_unidade) = int(strcd_unidade) OR int(cd_unidade) = "" then%>
<br class="no_print">
				
					<table style="border:0px solid silver; border-collapse:collapse; width:350px; font-size:11px; text-decoration:bold;" align="center" class="no_print">
						<tr><%borda_linhas = 2%>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid red" rowspan="2" align="center">&nbsp;<%=zero(conta_funcionarios)%><!--input type="checkbox" name="cd_impressao" value="" checked--></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="3" bgcolor="#c0c0c0">&nbsp;<a href="javascript:void(0);" return false;" onClick="window.open('empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro&pag=0&cod=<%=cd_funcionario%>&busca=','Cadastro','width=780,height=600,scrollbars=1')"><%=mid(nm_funcionario,1,30)%></a>&nbsp;<%=nm_sigla%></td>								
						<%strsql = "SELECT * FROM VIEW_funcionario_cargo WHERE cd_funcionario = '"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/"&UltimoDiaMes(mes_sel,ano_sel)&"/"&ano_sel&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
						Set rs_cargo = dbconn.execute(strsql)
							if not rs_cargo.EOF then
								nm_cargo = rs_cargo("nm_cargo")
							end if%>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" align="left" bgcolor="#c0c0c0">&nbsp;<%=nm_cargo%></td>							
						<%'strsql = "SELECT * FROM View_funcionario_unidade WHERE cd_funcionario = '"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/"&UltimoDiaMes(mes_sel,ano_sel)&"/"&ano_sel&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
						'Set rs_cargo = dbconn.execute(strsql)
							'if not rs_cargo.EOF then
								'cd_unidade = rs_cargo("cd_unidade")
									if cd_unidade = 27 then
										local_trabalho = "Escritório"
									elseif cd_unidade <> 27 then
										local_trabalho = "Hospital"
									end if
							'end if%>
							<td><%if unico = 1 then%><img src="../../imagens/ic_print.gif" alt="Imprimir" width="18" height="18" border="0" onClick="window.print()"><%else%>&nbsp;<%end if%></td>						
						</tr>						
						<tr bgcolor="#c0c0c0">
						<%strsql = "SELECT * FROM View_funcionario_horario WHERE cd_funcionario = '"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/"&UltimoDiaMes(mes_sel,ano_sel)&"/"&ano_sel&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
						Set rs_horario = dbconn.execute(strsql)
							if not rs_horario.EOF then
								hr_entrada = rs_horario("hr_entrada")
								hr_saida = rs_horario("hr_saida")
								nm_intervalo = rs_horario("nm_intervalo")
							end if
								if hr_entrada = "" Then hr_entrada = "00:00" end if
								if hr_saida = "" Then hr_saida = "00:00" end if%>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom">&nbsp;<%=zero(hour(hr_entrada))%>:<%=zero(minute(hr_entrada))%></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" align="center">&nbsp;<%=nm_intervalo%></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom">&nbsp;<%=zero(hour(hr_saida))%>:<%=zero(minute(hr_saida))%></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="5">&nbsp;<%'=cd_unidade%><%=Ucase(local_trabalho)%></td>						
						</tr>
						<tr>
							<td><img src="../../imagens/px.gif" alt="" width="10" height="1" border="0"></td>
							<td><img src="../../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
							<td><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
							<td><img src="../../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
							<td><img src="../../imagens/px.gif" alt="" width="150" height="2" border="0"></td>
						</tr>
						<tr class="ok_print"><td>&nbsp;</td></tr>
					</table>
				
					<%if unico <> 1 then%>
						<div style="page-break-after:always;" align="center" class="ok_print"></div>	
					<%end if%>
				<%conta_funcionarios = conta_funcionarios + 1
				end if%>	
	<%
	numero = 1
	protocolo = protocolo + 1
	quantidade = quantidade + 1
	num = num + 1
	espacamento = espacamento + 1
	'conta_funcionarios = conta_funcionarios + 1
	
	nm_cargo = ""
	local_trabalho = ""
	hr_entrada = ""
	hr_saida = ""
	nm_intervalo = ""
	
	rs.movenext
	loop%>
												
				<!--/td>
			</tr>
		</table-->
	


<%end if%>
</body>
</html>
