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


dt_inicio_sys = "01/01/2007"

func_sel = request("func_sel")
	func_ver = func_sel
	if func_sel = 0 then
		func_sel = "todos"
	end if


dia_sel = zero(request("dia_sel"))
mes_sel = zero(request("mes_sel"))
ano_sel = request("ano_sel")
data_atual = dia_sel&"/"&mes_sel&"/"&ano_sel

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
</SCRIPT>

</head>

<body><br>

<%'Passo 1
'*********************************************************
'*				      1ª Parte	 					  	 *
'*********************************************************
'* 1 - Seleção da data *
'************************%>

<%'if mes_sel = "" AND ano_sel = "" AND cd_codigo = "" Then
if mes_sel = "" AND ano_sel = "" Then%>
				<table border="0" align="center" style="border:1 solid black;">
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
					    <td colspan="100%" align="center">Selecione o período:</td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<form action="<%=nome_arquivo%>" name="Ano" method="get">
					<input type="hidden" name="tipo" value="<%=tipo%>">
					<!--input type="hidden" name="cd_codigo" value="<%=cd_codigo%>"-->
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">					
					<tr>
						<td align="right">MÊS</td>
						<td align="left" colspan="100%">
						<%ano_inicio = year(dt_inicio_sys)
						  mes_inicio = month(dt_inicio_sys)
								if int(ano_sel) > int(ano_inicio) then
									mes_inicio = "1"
									mes_alt = mes_prod + 2
								else
									mes_alt = 13
								end if%>
								
						<select name="mes_sel">
								<%do while not mes_inicio = 13
								response.write(mes_inicio&"-")
								response.write(mes_prod&".")%>
								<option value="<%=mes_inicio%>" <%if int(mes_sel) = month(now) then%><%response.write("selected")%><%end if%>><%=mes_selecionado(mes_inicio)%></option>
								<%
								
								
								mes_inicio = mes_inicio + 1
								loop%>			
						</select>						
						</td>
					</tr>
					<tr>
						<td align="right">ANO</td>
						<td align="left" colspan="100%"><br>
						<select name="ano_sel">
								<%ano_inicio = year(dt_inicio_sys)
								response.write(ano_inicio)
								do while not ano_inicio = ano_prod + 1%>
								<option value="<%=ano_inicio%>" <%if ano_inicio=ano_prod then%><%response.write("selected")%><%end if%>><%=ano_inicio%></option>
								<%ano_inicio = ano_inicio + 1
								loop%>			
						</select><br>
						<br>
						</td>
					</tr>					
					<tr>
						<td align="right">
						<input type="submit" value="OK">
						</td>
					</tr></form>
					
					<tr>
					    <td colspan="100%">
							<p>&nbsp;</p>
							<p>&nbsp;</p></td>
					</tr>
					<tr>
				</table>
<%'*******************************
'* 1.3 - Seleção de Funcionários *
'*********************************
Else
'elseif mes_sel <> "" AND ano_sel <> "" AND func_sel = "" Then%>
				<table border="1" align="center" id="no_print">
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
						<td colspan="100%"><br>&nbsp;</td>
					</tr>
					<form action="<%=nome_arquivo%>" name="mes" method="get">
					<input type="hidden" name="tipo" value="<%=tipo%>">
					<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
					<input type="hidden" name="arquivo" value="<%=nome_arquivo%>">
					<input type="hidden" name="ano_sel" value="<%=ano_sel%>">
					<input type="hidden" name="mes_sel" value="<%=int(mes_sel)%>">
					<tr>
					    <td colspan="100%" align="center">Selecione o colaborador abaixo:</td>
					</tr>
					<tr>
						<td align="center" colspan="100%">
						<%hoje = month(now)&"/"&day(now)&"/"&year(now)
						
						xsql = "Select * From View_funcionario Where (cd_regime_trabalho = 1) AND dt_desliga IS NULL OR (cd_regime_trabalho = 1) AND (DATEDIFF(day, '"&hoje&"', dt_desliga) > 0) ORDER BY nm_nome,nm_sobrenome"
						Set rs = dbconn.execute(xsql)%>
						
						
						<select name="func_sel">
							<option value="0">Todos ativos</option>
							<%while not rs.EOF
								cd_funcionario = rs("cd_codigo")
								nome = rs("nm_nome") &" "&rs("nm_sobrenome")
							%>							
							<option value="<%=cd_funcionario%>" <%if int(func_ver)=int(cd_funcionario) then response.write("selected") end if%>><%=nome%></option>
							<%rs.movenext
							wend%>									
						</select><br>						
						<br>
						<input type="submit" value="OK">	&nbsp; <input type="button" value="Limpa" onclick="location.replace('empresa.asp?tipo=cartao_ponto');">	
																										  				
						</td>
					</tr>
					</form>
					
					<tr>
					    <td colspan="100%" align="center"><br>&nbsp;<%=ucase(mes_selecionado(int(mes_sel)))%> / <%=ano_sel%><br>&nbsp;</td>
					</tr>
					<tr>
				</table>
<%'***********************
'* 1.4 - Corpo da página *
'*************************
'elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo = "" Then%>
<!--tr><td><%'=data_atual%></td></tr-->
<%'else%>

<%quantidade = 0
num = 1%>


	 
	
		<!--table style="border:1px solid red; width:<%=larg_tabs%>px;" id="ok__print_">
		
			<tr>				
				<td><img src="../../imagens/blackdot.gif" alt="" width="210" height="1" border="0"></td>
				<td style="border:1px solid blue;"-->				
				<%
				'response.write(func_sel)
				
				If func_sel = "todos" Then
					'xsql = "Select * From View_funcionario Where (cd_regime_trabalho = 1) AND dt_desliga IS NULL OR (cd_regime_trabalho = 1) AND (DATEDIFF(day, '"&hoje&"', dt_desliga) > 0) ORDER BY nm_nome,nm_sobrenome"
					'xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao > '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' OR dt_admissao <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao is NULL"' ORDER BY "&ordem_funcionarios&""
					xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&mes_sel&"/1/"&ano_sel&"' AND dt_demissao > '"&mes_sel&"/"&ultimodiames(mes_sel,ano_sel)&"/"&ano_sel&"' OR dt_admissao <= '"&mes_sel&"/01/"&ano_sel&"' AND dt_demissao is NULL ORDER BY nm_nome"
					
					'xsql = ""
				else
					xsql = "Select * From View_funcionario Where cd_codigo = "&func_sel&""
					
				end if
				Set rs = dbconn.execute(xsql)
	
	
	Do While Not rs.eof 
	cd_funcionario = rs("cd_codigo")
	nm_funcionario = rs("nm_nome")
	'nm_sobrenome = rs("nm_sobrenome")
			
	
	'**********************************************************
	'*** DEFINE A APARENCIA EM LARGURA DE TODO O FORMULÁRIO ***
	'**********************************************************
	larg_tabs = 325 + 210
	'**************
	pad_item = round((larg_tabs/4)*0.75,0)
	pad_espaco = round((larg_tabs/4)*0.25,0)
	altura_linha = "24"
	borda_linhas = "0px"
	espaco_top = "13"
	
	'*********************************************************** // CORPO DO CARTÂO \\ ************************************************%>
					<table style="border:0px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; font-size:11px; text-decoration:bold;" class="relatorio_campos" id="nok_print">
						<tr><td colspan="100%" align="center"><img src="../../imagens/px.gif" alt="" width="1" height="<%=espaco_top%>" border="0" align="middle"></td></tr>
						<tr id="oks_print">
							<td rowspan="10"><img src="../../imagens/px.gif" alt="" width="210" height="1" border="0"></td>
							<td style="height:<%=altura_linha+1%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2">&nbsp;<!--Nº--></td>
							<td style="height:<%=altura_linha+1%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="4"> &nbsp; &nbsp; CMI CIRURGICA LTDA</td>		
						</tr>							
						<tr id="oks_print">
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="4"> &nbsp; &nbsp; 00.255640/0001-42</td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2">&nbsp;<!--Atividade Econ.--></td>		
						</tr>							
						<tr>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="6"> &nbsp; &nbsp; <a href="empresa.asp?tipo=cadastro&pag=0&cod=<%=cd_funcionario%>&busca=" target="_blank"><%=nm_funcionario%>.<%=cd_funcionario%></a></td>								
						</tr>
						<tr id="oks_print">
						<%'strsql = "SELECT * FROM VIEW_funcionario_cargo WHERE cd_funcionario = '"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/"&UltimoDiaMes(mes_sel,ano_sel)&"/"&ano_sel&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
						strsql = "SELECT * FROM VIEW_funcionario_cargo WHERE cd_funcionario = '"&cd_funcionario&"'  ORDER BY dt_atualizacao DESC"
						Set rs_cargo = dbconn.execute(strsql)
							if not rs_cargo.EOF then
								nm_cargo = rs_cargo("nm_cargo")
							end if%>						
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2">&nbsp;<!--Registro--></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2">&nbsp;<!--Nº CTPS--></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2" align="center">&nbsp; <%=nm_cargo%></td>	
						</tr>
						<tr id="oks_print">
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="5">&nbsp;&nbsp;&nbsp;<%'=cd_unidade%><%'=Ucase(local_trabalho)%></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="2">&nbsp;</td>
						</tr>
						<tr id="oks_print">							
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" align="center" colspan="5">&nbsp;<%=ucase(mes_selecionado(int(mes_sel)))%></td>
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
<%	'*********************************************************** // FIM DO CORPO DO CARTÂO \\ ************************************************%>

					<table style="border:0px solid silver; border-collapse:collapse; width:350px; font-size:11px; text-decoration:bold;" align="center" class="relatorio_campos" id="no_print">
						<tr><%borda_linhas = 2%>
							
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid red" rowspan="2">&nbsp;<!--input type="checkbox" name="cd_impressao" value="" checked--></td>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="3" bgcolor="#c0c0c0">&nbsp;<a href="empresa.asp?tipo=cadastro&pag=0&cod=<%=cd_funcionario%>&busca=" target="_blank"><%=mid(nm_funcionario,1,30)%></a>&nbsp;<%'=nm_sobrenome%></td>								
						<%strsql = "SELECT * FROM VIEW_funcionario_cargo WHERE cd_funcionario = '"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/"&UltimoDiaMes(mes_sel,ano_sel)&"/"&ano_sel&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
						Set rs_cargo = dbconn.execute(strsql)
							if not rs_cargo.EOF then
								nm_cargo = rs_cargo("nm_cargo")
							end if%>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" align="left" bgcolor="#c0c0c0">&nbsp;<%=nm_cargo%></td>							
						<%strsql = "SELECT * FROM View_funcionario_unidade WHERE cd_funcionario = '"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/"&UltimoDiaMes(mes_sel,ano_sel)&"/"&ano_sel&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
						Set rs_cargo = dbconn.execute(strsql)
							if not rs_cargo.EOF then
								cd_unidade = rs_cargo("cd_unidade")
									if cd_unidade = 27 then
										local_trabalho = "Escritório"
									elseif cd_unidade <> 27 then
										local_trabalho = "Hospital"
									end if
							end if%>							
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
						<tr id="ok_print"><td>&nbsp;</td></tr>
					</table>
					<div style="page-break-after:always;" align="center" id="ok_print"></div>	
	
	<%
	numero = 1
	protocolo = protocolo + 1
	quantidade = quantidade + 1
	num = num + 1
	espacamento = espacamento + 1
	
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
