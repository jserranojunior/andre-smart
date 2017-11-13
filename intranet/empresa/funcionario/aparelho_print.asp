<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
	<title>Untitled</title>
<%cd_user = session("cd_codigo")
pasta_loc = "intranet\empresa\"
arquivo_loc = "aparelho_print.asp"

cd_unidade = request("cd_unidade")
qtd_paginas = request("qtd_paginas")
qtd_etiquetas = request("qtd_etiquetas")
mensagem = request("mensagem")
tipo_print = request("tipo_print")
espacamento = 1
mostra_mes_inicial = request("mostra_mes_inicial")
posicao_topo_cartao = request("posicao_topo_cartao")
str_texto = request("str_texto")

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
dia_final = ultimodiames(mes_sel,ano_sel)
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
<!--#include file="../../includes/arquivo_loc.asp"-->

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
						<td align="center" colspan="2" style="font-size:14px;"><b>IMPRIMIR: Termo de responsabilidade<br>Celular e acessórios</b></td>
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
				<table border="1" align="center"  class="no_print">
					<tr bgcolor="#cococo">
						<td align=center colspan="1"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
						<td colspan="1" align="center" style="background-color:gray; color:white;">&nbsp;<br>IMPRESSÃO<br> TERMO DE RESPONSABILIDADE DE <br>APARELHO CELULAR E ACESSÓRIOS&nbsp;</td>
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
					    <td colspan="1" align="center">Selecione o colaborador abaixo:</td>
					</tr>
					<tr>
						<td align="left" colspan="1">
						<%hoje = month(now)&"/"&day(now)&"/"&year(now)
						ordem_funcionarios = "nm_nome "
						
						'xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao > '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND cd_contrato = 1 OR dt_admissao <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao is NULL AND cd_contrato = 1 ORDER BY nm_nome"
						xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao_geral <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao_geral > '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND cd_contrato = 1 OR dt_admissao_geral <= '"&mes_sel&"/"&dia_sel&"/"&ano_sel&"' AND dt_demissao_geral is NULL AND cd_contrato = 1 ORDER BY nm_nome"
						'xsql = "SELECT * FROM View_funcionario_contrato_lista WHERE dt_admissao_geral <= '"&dia_sel&"/"&mes_sel&"/"&ano_sel&"' AND dt_demissao_geral > '"&dia_sel&"/"&mes_sel&"/"&ano_sel&"' AND cd_contrato = 1 OR dt_admissao_geral <= '"&dia_sel&"/"&mes_sel&"/"&ano_sel&"' AND dt_demissao_geral is NULL AND cd_contrato = 1 ORDER BY nm_nome"
						'xsql = "up_funcionario_contrato_lista3 @dt_data_i='"&mes_sel&"/1/"&ano_sel&"', @dt_data_f='"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"', @outras_variaveis='"&outras_variaveis&"'"
						
						Set rs = dbconn.execute(xsql)%>
						
						
						<select name="func_sel" onChange="javascript:submit();">
							<option value="0">Todos ativos</option>
							<%while not rs.EOF
								cd_funcionario = rs("cd_funcionario")
								nome = rs("nm_nome")' &" "&rs("nm_sobrenome")
							%>							
							<option value="<%=cd_funcionario%>" <%if int(cd_funcionario)=int(func_ver) then response.write("selected") end if%> ><%=nome%></option>
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
						
						
						
						<!--input type="checkbox" name="mostra_mes_inicial" value="1" <%if mostra_mes_inicial = 1 then%>Checked<%end if%>>Mostra apenas o mês inicial no cartão<br><br>
											
						<input type="radio" name="posicao_topo_cartao" value="0" <%if posicao_topo_cartao = "" OR posicao_topo_cartao = 0 then%>Checked<%end if%> onChange="javascript:submit();">Posição normal<br>
						<input type="radio" name="posicao_topo_cartao" value="1" <%if posicao_topo_cartao = 1 then%>Checked<%end if%> onChange="javascript:submit();">Mais para cima<br>
						<input type="radio" name="posicao_topo_cartao" value="2" <%if posicao_topo_cartao = 2 then%>Checked<%end if%> onChange="javascript:submit();">Mais para baixo-->
						<br><br>
						<b>Termo:</b> <br>
						<input type="radio" name="str_texto" value="1" <%if str_texto = "" OR str_texto = 1 then%>Checked<%end if%> onChange="javascript:submit();"> - Individual<br>
						<input type="radio" name="str_texto" value="2" <%if str_texto = 2 then%>Checked<%end if%> onChange="javascript:submit();"> - Comunitário
						
						<br><%'=posicao_topo_cartao%>
						<br>						
						<input type="submit" value="OK">	&nbsp; <input type="button" value="Limpa" onclick="location.replace('empresa.asp?tipo=cartao_ponto');">
						
						&nbsp;&nbsp;&nbsp;&nbsp;
						<img src="../../imagens/ic_print.gif" alt="" width="24" height="24" border="0" onClick="window.print();">&nbsp;&nbsp;&nbsp;<img src="../../imagens/ic_print_view.gif" alt="" width="24" height="26" border="0" onclick="visualizarImpressao();"></td>
					</tr>
					</form>
					
					<tr>
					    <td colspan="1" align="center">Período selecionado, referente a:<br>&nbsp;<%'=zero(dia_sel)%>
						<%
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
					'xsql = "up_funcionario_contrato_lista_teste @dt_data='"&ano_sel&zero(mes_sel)&"', @dt_atualizacao='"&mes_print&"/"&ultmo_dia_mes_print&"/"&ano_print&"', @ordem_funcionarios='nm_nome', @outras_variaveis=''"
					'xsql = "up_funcionario_contrato_lista3 @dt_data_i='"&mes_sel&"/1/"&ano_sel&"', @dt_data_f='"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @dt_atualizacao = '"&mes_sel&"/"&dia_final&"/"&ano_sel&"', @ordem_funcionarios='"&ordem_funcionarios&"', @outras_variaveis='"&outras_variaveis&"'"
					ano_competencia_i = ano_sel'ano_competencia
					mes_competencia_i = mes_sel'mes_competencia
					dia_competencia_i = 25
							
					ano_competencia_f = ano_sel'ano_competencia
					mes_competencia_f = mes_sel'mes_competencia
					final_competencia_f = dia_final'1'final_competencia'30
					
					xsql = "up_funcionario_contrato_lista4 @dt_data_i='"&mes_competencia_i&"/"&dia_competencia_i&"/"&ano_competencia_i&"', @dt_data_f='"&mes_competencia_f&"/"&final_competencia_f&"/"&ano_competencia_f&"', @dt_atualizacao = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @ordem_funcionarios='"&ordem_funcionarios&"', @outras_variaveis='"&outras_variaveis&"'"
							
				else
					xsql = "Select top 1 * From View_funcionario_contrato_lista Where cd_funcionario = "&func_sel&""
					unico = 1
				end if
				Set rs = dbconn.execute(xsql)
	
	Do While Not rs.eof 
	'cd_funcionario = rs("cd_funcionario")
	cd_funcionario = rs("cd_codigo")
	nm_funcionario = rs("nm_nome")
	cd_matricula = rs("cd_matricula")
	nm_rg = rs("nm_rg")
	
	'if func_sel <> "0" then
	'	strcd_unidade = 0
	'end if
	
		if strcd_unidade <> 0 then
			cd_unidade = rs("cd_unidade")
		else
			cd_unidade = 0
		end if
	
	'nm_sigla = rs("nm_sigla")
	
			xsql_1 = "Select * From View_funcionario_aparelhos Where cd_funcionario = "&cd_funcionario&" and cd_suspenso <> 1 ORDER BY dt_atualizacao desc"
				Set rs_1 = dbconn.execute(xsql_1)
					if not rs_1.EOF then
						nm_modelo = rs_1("nm_modelo")
						nm_apelido = rs_1("nm_apelido")
						nr_numero = rs_1("nr_numero")
							nr_numero_1 = mid(nr_numero,1,1)&"-"
							nr_numero_2 = mid(nr_numero,2,4)&"-"
							nr_numero = nr_numero_1&nr_numero_2&mid(nr_numero,6,4)
					end if
	
	'**********************************************************
	'*** DEFINE A APARENCIA EM LARGURA DE TODO O FORMULÁRIO ***
	'**********************************************************
	larg_tabs = 50'325 + 210
	'**************
	pad_item = round((larg_tabs/4)*0.75,0)
	pad_espaco = round((larg_tabs/4)*0.25,0)
	altura_linha = "24"
	borda_linhas = "0px"
	
	espaco_top = 15 '7 1
		if posicao_topo_cartao = 1 then
			espaco_top = espaco_top - 5
		elseif posicao_topo_cartao =  2 then
			espaco_top = espaco_top + 5
		end if
	'*********************************************************** // CORPO DO TERMO DE RESPOSABILIDADE \\ ************************************************%>
				<%if int(cd_unidade) = int(strcd_unidade) AND nr_numero <> "" OR int(cd_unidade) = "" AND nr_numero <> "" then%>
					<%'strsql = "SELECT * FROM VIEW_funcionario_cargo WHERE cd_funcionario = '"&cd_funcionario&"' AND dt_atualizacao <= '"&mes_sel&"/"&UltimoDiaMes(mes_sel,ano_sel)&"/"&ano_sel&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
					'	Set rs_cargo = dbconn.execute(strsql)
					'		if not rs_cargo.EOF then
					'			nm_cargo = rs_cargo("nm_cargo")
					'		end if%>
					<%if unico <> 1 AND conta_funcionarios > 1 then%>
						<div style="page-break-after:always;" align="center" class="ok_print">&nbsp;</div>	
					<%end if%>
					<table align="center" style="border:0px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; font-family:Times New Roman, Arial; font-size:17px; text-decoration:bold; text-align:justify;" class="ok_print">
						<tr><td colspan="1" align="center"><img src="../../imagens/px.gif" alt="" width="1" height="<%=espaco_top%>" border="0" align="middle"></td></tr>
						<tr>
							<td style="height:<%=altura_linha+1%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="4" align="center">
								&nbsp;<br>
								&nbsp;<br>
								&nbsp;<b style="font-size:50px;">CMI CIRURGICA LTDA</b><br>
								&nbsp;<br>
								&nbsp;<br>
								&nbsp;<br>
								&nbsp;
							</td>		
						</tr>
						<tr>
							<td style="height:<%=altura_linha+1%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="4" align="center">
								<b>TERMO DE RESPONSABILIDADE DE APARELHO CELULAR e <BR>ACESSÓRIOS</b><br>
								&nbsp;<br>
								&nbsp;<br>
								&nbsp;
							</td>		
						</tr>
						<%if str_texto = 1 then%>				
						<tr>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="4"> 
								A empresa CMI CIRÚRGICA LTDA, pessoa jurídica de direito privado devidamente inscrita no CNPJ sob o nº 05.204.601/0001-30, entrega neste ato, o aparelho celular modelo: <b><%=nm_modelo%></b>, numero: <b><%=nr_numero%></b>, bem como seus acessórios também descritos a seguir: Carregador USB <b>em mãos</b> do funcionário <b><%=nm_funcionario%></b>, <%=nm_cargo%>, portador do RG sob o nº <b><%=nm_rg%></b> doravante denominado simplesmente "USUÁRIO" sob as seguintes condições:<br>
								1. O equipamento deverá ser utilizado ÚNICA e EXCLUSIVAMENTE durante a jornada de trabalho e para necessária comunicação entre colegas de trabalho durante o desenvolvimento do respectivo ofício do USUÁRIO;<br>
								2. Ficará o USUÁRIO responsável pelo uso e conservação do equipamento;<br>
								3. O USUÁRIO tem somente a DETENÇÃO do respectivo bem móvel, tendo em vista o uso exclusivo para prestação de serviços profissionais e NÃO a PROPRIEDADE do equipamento, sendo terminantemente proibido o empréstimo, aluguel ou cessão deste a terceiros;<br>
								4. Ao término da prestação de serviço ou do contrato individual de trabalho, independentemente de sua modalidade de extinção contratual, o USUÁRIO <b style="text-decoration:underline;">compromete-se</b> a devolver o equipamento em perfeito estado no mesmo dia em que for comunicado ou se comunique seu desligamento, considerando-se sempre, o desgaste natural pelo uso normal do equipamento.<br>
								&nbsp;<br>
								&nbsp;<br>
								&nbsp;<br>
								&nbsp;
							</td>
						</tr>
						<%elseif str_texto = 2 then%>
						<tr>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="4"> 
								A empresa CMI CIRÚRGICA LTDA, pessoa jurídica de direito privado devidamente inscrita no CNPJ sob o nº 05.204.601/0001-30, entrega neste ato, o aparelho celular modelo: <b><%=nm_modelo%></b>, numero: <b><%=nr_numero%></b>, bem como seus acessórios também descritos a seguir: Carregador USB para utilização comunitária entre seus colaboradores, funcionario: <b><%=nm_funcionario%></b>, <%=nm_cargo%>, portador do RG sob o nº <b><%=nm_rg%></b> doravante denominado simplesmente "USUÁRIOS" sob as seguintes condições:<br>
								1. O equipamento deverá ser utilizado ÚNICA e EXCLUSIVAMENTE durante a jornada de trabalho e para necessária comunicação entre colegas de trabalho durante o desenvolvimento do respectivo ofício do USUÁRIO;<br>
								2. Ficarão os USUÁRIOS responsáveis pelo uso e conservação do equipamento;<br>
								3. OS USUÁRIOS têm somente a DETENÇÃO do respectivo bem móvel, tendo em vista o uso exclusivo para prestação de serviços profissionais e NÃO a PROPRIEDADE do equipamento, sendo terminantemente proibido o empréstimo, aluguel ou cessão deste a terceiros;<br>
								4. Ao término da prestação de serviço ou do contrato individual de trabalho, independentemente de sua modalidade de extinção contratual, os USUÁRIOS <b style="text-decoration:underline;">comprometem-se</b> a devolver o equipamento em perfeito estado no mesmo dia em que for comunicado ou se comunique seu desligamento, considerando-se sempre, o desgaste natural pelo uso normal do equipamento.<br>
								&nbsp;<br>
								&nbsp;<br>
								&nbsp;<br>
								&nbsp;
							</td>
						</tr>	
						<%end if%>						
						<tr>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="6">São Paulo,................ de ..............................................  de 2014.<br>
							&nbsp;<br>
							&nbsp;<br>
							&nbsp;<br>
							&nbsp;<br>
							_______________________________________________<br>
							Assinatura do funcionário

							</td>								
						</tr>
						
						
						
						<!--tr class="ok_print"-->
						<tr>
							<td><img src="../../imagens/px.gif" alt="" width="600" height="1" border="0"></td>
						</tr>
						<tr><td>&nbsp;</td></tr>
					</table>
					<%'if unico <> 1 AND not rs_1.EOF then%>
						<!--div style="page-break-after:always;" align="center" class="ok_print">y</div-->	
					<%'end if%>
				<%end if%>
<%	'*********************************************************** // FIM DO CORPO \\ ************************************************%>
				<%if int(cd_unidade) = int(strcd_unidade) AND nr_numero <> ""  OR int(cd_unidade) = "" AND nr_numero <> "" then
				'if int(cd_unidade) = int(strcd_unidade)  OR int(cd_unidade) = "" then%>
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
								
								xsql_1 = "Select * From View_funcionario_aparelhos Where cd_funcionario = "&cd_funcionario&" and cd_suspenso <> 1 ORDER BY dt_atualizacao desc"
								Set rs_1 = dbconn.execute(xsql_1)
									if not rs_1.EOF then
										nm_modelo = rs_1("nm_modelo")
										nm_apelido = rs_1("nm_apelido")
										nr_numero = rs_1("nr_numero")
											nr_numero_1 = mid(nr_numero,1,1)&"-"
											nr_numero_2 = mid(nr_numero,2,4)&"-"
											nr_numero = nr_numero_1&nr_numero_2&mid(nr_numero,6,4)
									end if
							end if
								if hr_entrada = "" Then hr_entrada = "00:00" end if
								if hr_saida = "" Then hr_saida = "00:00" end if%>
							<td style="height:<%=altura_linha%>px; border:<%=borda_linhas%> solid silver;" valign="bottom" colspan="3">&nbsp;<%=nr_numero%></td>
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
						<!--div style="page-break-before:always;" align="center" class="ok_print">y</div-->	
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
	nr_numero = ""
	
	rs.movenext
	loop%>
												
				<!--/td>
			</tr>
		</table-->
	


<%end if%>
</body>
</html>
