<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%tela = request("tela")

if tela = "" Then
	div_height = "200px"
'Else
	'div_height = "auto"
end if
%>
<style type="text/css">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:12px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #000000; }  
.print {color: #000000;font-family: verdana;font-size:9px;text-decoration:none;}
.divrolagem {
/* define barra de rolagem automatica quando o 
conteudo ultrapassar o limite em x ou y */
    overflow: auto; 
/* define o limite maximo da autura do div */
    height: <%=div_height%>;
/* define o limite maximo da largura do div */
    width: auto;}
[code]

-->
</style>

<%
dt_inicio_sys = "01/03/2008"

mes_prod = month(now)
ano_prod = year(now)

dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")
data_atual = ano_sel&mes_sel

cd_codigo = request("cd_codigo")
acao = request("acao")

'Verifica o ultimo mes encerrado
%>

<html>
<head>
	<TITLE> Relatório do Banco de Horas </TITLE>
	<META NAME="Generator" CONTENT="EditPlus">
	<META NAME="Author" CONTENT="">
	<META NAME="Keywords" CONTENT="">
	<META NAME="Description" CONTENT="">
</head>
<LINK href="css/estilo.css " type=text/css rel=stylesheet >

<body><br>
<table width="600" cellspacing="0" cellpadding="0" border="0" class="textoPadrao" align="center">
	<tr>
		<td align="center">
			<table width="100%" cellspacing="1" cellpadding="0" border="0" class="txt" bordercolordark="#FFFFFF">
					<!--tr bgcolor=#cococo>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=200 height=0></td>
						<td><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=150 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=200 height=0></td>
						<td><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td><img src="imagens/blackdot.gif" width=80 height=0></td>
						<td><img src="imagens/blackdot.gif" width=75 height=0></td>
					</tr-->
					<tr>
						<td colspan="100%">&nbsp;<b>FUNCIONÁRIOS - <font color="red"><b>Relatório</b>: Banco de horas</font></b></td>
					</tr>
					<tr>
						<td align=center colspan="100%">&nbsp;<!--img src="../../imagens/px.gif" height="10"--></td>
					</tr>
					
				<%'*******************************
				'*         Seleção do Ano        *
				'*********************************
				'if ano_sel = "" then 'AND mes_sel = "" Then
					
					if year(dt_inicio_sys) = int(ano_prod) then
							ano_sel = year(dt_inicio_sys)
						end if
					if ano_sel = "" Then%>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr><td align=center colspan="100%">&nbsp;</td></tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
					    <td colspan="10" align="center">Selecione o ano abaixo:</td>
					</tr>
					<form action="empresa_banco_horas.asp" name="Ano" method="get">
					<tr>
						<td align="center" colspan="100%"><br>
						<select name="ano_sel">
								<%ano_inicio = year(dt_inicio_sys)
								response.write(ano_inicio)
								do while not ano_inicio = ano_prod + 1%>
								<option value="<%=ano_inicio%>" <%if ano_inicio=ano_prod then%><%response.write("selected")%><%end if%>><%=ano_inicio%></option>
								<%ano_inicio = ano_inicio + 1
								loop%>			
						</select><br>
						<br>
						<input type="submit" value="OK">						
						</td>
					</tr>
					</form>
					<tr>
					    <td colspan="10">
							<p>&nbsp;</p>
							<p>&nbsp;</p></td>
					</tr>
				<%
					'************************
					'* 1.2 - Seleção do MÊS *
					'************************
					
					Elseif ano_sel <> "" AND mes_sel = "" then%>
					
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td colspan="5">&nbsp;</td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					<form action="empresa_banco_horas.asp" name="mes" method="get">
					<input type="hidden" name="ano_sel" value="<%=ano_sel%>">
					<tr>
					    <td colspan="10" align="center">Selecione o mês abaixo:</td>
					</tr>
					<tr>
						<td align="center" colspan="100%"><br>
						
						<%ano_inicio = year(dt_inicio_sys)
						  mes_inicio = month(dt_inicio_sys)
								if int(ano_sel) > int(ano_inicio) then
									mes_inicio = "1"
								end if%>
								
						<select name="mes_sel">
								<%do while not mes_inicio = mes_prod + 1
								response.write(mes_inicio&"-")
								response.write(mes_prod&".")%>
								<option value="<%=mes_inicio%>" <%if mes_inicio=mes_prod Then'- 1 then%><%response.write("selected")%><%end if%>><%=mes_selecionado(mes_inicio)%></option>
								<%
								mes_inicio = mes_inicio + 1
								loop%>			
						</select><br>
						<br>						
						<input type="submit" value="OK">						
						</td>
					</tr>
					</form>
					
					<tr>
					    <td colspan="10">
							<p>&nbsp;</p>
							<p>&nbsp;</p></td>
					</tr>
					
				
				<%Elseif ano_sel <> "" AND mes_sel <> "" Then%>
					<tr>
						<td><b>Situação em <%=mes_selecionado(mes_sel-1)%></b></td>
						<td colspan="4" align="center">4 últimos meses</td>
						<td align="center">Total</td>
					</tr>
					<tr>
						<td><img src="../../imagens/blackdot.gif" width="300" height=0></td>
						<td><img src="../../imagens/blackdot.gif" width="80" height=0></td>
						<td><img src="../../imagens/blackdot.gif" width="80" height=0></td>
						<td><img src="../../imagens/blackdot.gif" width="80" height=0></td>
						<td><img src="../../imagens/blackdot.gif" width="80" height=0></td>
						<td><img src="../../imagens/blackdot.gif" width="80" height=0></td>
					</tr>
					<tr>
						<td align=center colspan="100%">&nbsp;<!--img src="../../imagens/px.gif" height="10"--></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="../../imagens/blackdot.gif" width="100%" height=0></td>
					</tr>
					<%nr_mes = int(mes_sel - 4)
					if nr_mes < 1 then
						nr_mes = 12
					end if%>
					<tr>
						<td>&nbsp;<b>Nome&nbsp;</b></td>
						<!--td align="center" colspan="3"><b>Parciais: 4 últimos meses&nbsp;</b></td-->
						<td align="center"><b>&nbsp;<%=mes_selecionado(nr_mes)%></b></td><%if nr_mes + 1 > 12 then%><%nr_mes = 0%><%end if%>
						<td align="center"><b>&nbsp;<%=mes_selecionado(nr_mes+1)%></b></td><%if nr_mes + 1 > 12 then%><%nr_mes = 0%><%end if%>
						<td align="center"><b>&nbsp;<%=mes_selecionado(nr_mes+2)%></b></td><%if nr_mes + 1 > 12 then%><%nr_mes = 0%><%end if%>
						<td align="center"><b>&nbsp;<%=mes_selecionado(nr_mes+3)%></b></td>
						<td align="right"><b>Saldo&nbsp;</b></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="../../imagens/blackdot.gif" width="100%" height=0></td>
					</tr>
					
						<%
						'xsql = "select * from View_relatorio_banco_horas_total"' where cd_codigo='"&cd_codigo&"'"
						xsql = "up_relatorio_banco_horas_total @dt_mes = '"&mes_sel&"', @dt_ano='"&ano&"'"
							Set rs = dbconn.execute(xsql)
							Do while NOT rs.EOF
						
							cd_total = rs("saldo")
							cd_codigo = rs("cd_codigo")
							nome = rs("nome")
							nr_mes = int(mes_sel - 4)
						%>
					<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#F2F2F2');" bgcolor="#F2F2F2">
						<td>&nbsp;<%'=cd_codigo%> <%=nome%></td>
						<!--td><%=dt_mes%></td-->
						
						<%'*** Mostra últimos meses ***
							numero = 4
						xsql2 = "select * from View_relatorio_banco_horas_meses where cd_codigo='"&cd_codigo&"' AND dt_mes between '"&nr_mes&"' AND '"&mes_sel&"' -1"
						Set rs2 = dbconn.execute(xsql2)
						Do while NOT rs2.EOF
						cd_codigo = rs2("cd_codigo")
						nm_nome = rs2("nome")
						dt_mes = rs2("dt_mes")
						cd_saldo = rs2("cd_horas_excedentes")
							
							'*** Procura saldo negativo e avisa ***
									str_negativo = Instr(1,cd_saldo,"-",1)'Compara se o caracter "-" aparece na string min
									If str_negativo = "1" Then
									cor_saldo = "red"
									Else
									cor_saldo  = "black"
									end if
							'****************************************************************************
								
							do while int(dt_mes) <> int(nr_mes) 'Meses em branco - anteriores
							%>
							<td align="right">&nbsp;</td>
							<%
							nr_mes = nr_mes + "1"
							loop	
							
							'**************************************************
							'* Mostra Quantidade de horas nos ultimos 4 Meses *
							'**************************************************%>
							
							<td align="right"><%'=dt_mes&"-"%><%'=nr_mes&"-"%> <font color=<%=cor_saldo%>><%=total_mes(cd_saldo)%>&nbsp;</font></td>
							<%'do while int(dt_mes) <> int(nr_mes) 
							'nr_mes = nr_mes + "1"
							'loop
							
						nr_mes = nr_mes + "1"
						numero = nr_mes
						rs2.movenext
						loop%>
						
							<%'*** Procura saldo negativo e avisa ***
									str_total_negativo = Instr(1,cd_total,"-",1)'Compara se o caracter "-" aparece na string min
									If str_total_negativo = "1" Then
									cor_total = "red" 
									Else
									cor_total  = "black"
									end if
								'****************************************************************************
							num_mes = nr_mes
							
						'*********************************
						'* Meses em branco - Posteriores *
						'*********************************
						mes_atual = 7'month(now)
						do while numero < int(mes_sel)
							%>
							<td align="right">&nbsp;<%'=dt_mes&"-"%><%'=nr_mes-1&"-"%><%'=mes_sel&"-"%><%'=numero%></td>
							<%
							numero = numero + 1
							loop%>
						<td align="right"><font color=<%=cor_total%>><b><%=total_mes(cd_total)%>&nbsp;</b></font></td>
							<% 
							nr_mes = ""
							cor_saldo = ""
							cor_saldo2 = ""
							'numero = ""
							rs.movenext
							loop%>
						<%'end if%>
					
					<%
					nm_lista = nm_nome
					cd_meses = cd_meses + 1
					cor_saldo = ""
						if cd_meses > 3 then
							cd_meses = 1
						end if
					%>
					</tr>
					<%end if%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
