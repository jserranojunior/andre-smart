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
	<TITLE> Relatório de faltas </TITLE>
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
			<table width="100%" cellspacing="1" cellpadding="0" border="0" class="txt" bordercolordark="#FFFFFF" align="center">
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
						<td colspan="100%">&nbsp;<b>FUNCIONÁRIOS - <font color="red"><b>Relatório</b>: Faltas</font></b></td>
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
					<form action="empresa.asp" name="Ano" method="get">
					<input type="hidden" name="tipo" value="relatorio_faltas">
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
					<form action="empresa.asp" name="Ano" method="get">
					<input type="hidden" name="tipo" value="relatorio_faltas">
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
								<option value="<%=mes_inicio%>" <%if mes_inicio=mes_prod - 1 then%><%response.write("selected")%><%end if%>><%=mes_selecionado(mes_inicio)%></option>
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
						<td colspan="100%"><b> <%=ucase(mes_selecionado(mes_sel))%>/<%=ano_sel%></b></td>
						<!--td colspan="4" align="center">&nbsp;</td>
						<td align="center">&nbsp;</td-->
					</tr>
					<tr>
						<td><img src="../../imagens/px.gif" width="50" height=1></td>
						<td><img src="../../imagens/px.gif" width="300" height=2></td>
						<td><img src="../../imagens/px.gif" width="80" height=3></td>
					</tr>
									
						<%'xsql = "up_relatorio_banco_horas_total @dt_mes = '"&mes_sel&"', @dt_ano='"&ano&"'"
						xsql = "up_funcionarios_ausencias @cd_folga=3, @dt_data_mes='"&mes_sel&"', @dt_data_ano='"&ano_sel&"'"
						Set rs = dbconn.execute(xsql)
						
						Do while NOT rs.EOF
						
							cd_funcionario = rs("cd_funcionario")
							nm_nome = rs("nm_nome")
							nm_sobrenome = rs("nm_sobrenome")
							dt_data = rs("dt_data")
							txt_obs = rs("txt_obs")
				if nm_nome <> nome then%>
					<tr><td>&nbsp;</td></tr>
					<tr bgcolor="#F2F2F2" colspan="100%">
						<td colspan="100%"><font color="#800000">&nbsp;<%=nm_nome%>&nbsp;<%=nm_sobrenome%></font></td>
					</tr>
					<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height=1></td></tr>
					<tr bgcolor="#F2F2F2">
						<td align="center"><b>Dia</b></td>
						<td>&nbsp;<b>Motivo</b></td>
						<td align="center"><b>Qtd. faltas</b></td>
					</tr>
					<tr><td colspan="100%"><img src="../imagens/blackdot.gif" width="100%" height=1></td></tr>

					<%xsql_2 = "Select Count (*) as n_faltas FROM view_funcionarios_ausencias where cd_folga=3 and Month(dt_data)='"&mes_sel&"' AND Year(dt_data)='"&ano_sel&"' AND cd_funcionario='"&cd_funcionario&"'"
						Set rs_2 = dbconn.execute(xsql_2)
						
						n_faltas = rs_2("n_faltas")
				end if%>
					
					<tr bgcolor="#F2F2F2">
						<td align="center"><%=day(dt_data)%></td>
						<td><%=txt_obs%></td>
						<td><%'=qtd_faltas&":"%><%'=n_faltas%></td>
							<%nome = nm_nome
							qtd_faltas = qtd_faltas + 1
					%></tr>
					
					<%if qtd_faltas  = n_faltas then%>
					<tr><td colspan="100%"><img src="../imagens/blackdot.gif" width="100%" height=1></td></tr>
					<tr bgcolor="#F2F2F2" colspan="100%">
						<td>&nbsp;</td>
						<td align="right">&nbsp;Total de faltas:</td>
						<td align="center">&nbsp;<%=n_faltas%></td>
					</tr>
					<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height=2></td></tr>
										
					<%qtd_faltas = 0
					end if
							rs.movenext
							loop%>
										
					
					</tr>
					<%end if%>
					<tr><td><br>
					<br>
					<br>
					</td></tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
