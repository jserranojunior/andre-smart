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
data_inicio = "01/03/2008"

dia_prod = day(now)
mes_prod = month(now)
ano_prod = year(now)


dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")
data_atual = ano_sel&mes_sel



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
					<tr>
						<td colspan="100%">&nbsp;<b>FUNCIONÁRIOS - <font color="red">Banco de horas</font></b>
						&nbsp;&nbsp;&nbsp;<br>
						(Data hoje: <%=dia_prod%>/<b><%=mes_prod%>/<%=ano_prod%></b>)<br>
						(Data sele.: 00<%'=dia_sel%>/<%=mes_sel%>/<%=ano_sel%>)</td>
					</tr>
					<tr>
						<td align=center colspan="100%">&nbsp;<!--img src="../../imagens/px.gif" height="10"--></td>
					</tr>
							
		<%if mes_sel = "" OR ano_sel = "" Then%>
					<tr><td>Selecione o período do relatório</td></tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="../../imagens/blackdot.gif" width="100%" height=0></td>
					</tr>
					<form action="banco_horas_relat.asp" method="post">
					<%%>
					<tr>
						<td align=center colspan="100%">&nbsp;<%=ano_prod%>*<%=year(data_inicio)+1%>*<%=ano_selecionar%></td>
					</tr>
					</form>
					<tr>
						<td align=center colspan="100%">&nbsp;</td>
					</tr>
		<%Else%>
					
					<tr>
						<td>&nbsp;<b>Nome&nbsp;</b></td>
						<td align="right"><b>&nbsp;</b></td>
						<td align="right"><b>&nbsp;</b></td>
						<!--td align="right"><b>&nbsp;</b></td-->
						<!--td align="right"><b>Jun&nbsp;</b></td-->
						<td align="right"><b>Saldo&nbsp;</b></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="../../imagens/blackdot.gif" width="100%" height=0></td>
					</tr>
					<%
					conta_coop = "1"
					condicao = " Where  cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND dt_demissao >= "&data_atual&" OR cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND dt_demissao = null "
					xsql = "up_funcionarios_lista @cd_codigo='', @condicao='"&condicao&"', @ordem= 'nm_nome, nm_sobrenome'"
					Set rs = dbconn.execute(xsql)
					
					Do while NOT rs.EOF
					
					cd_codigo = rs("cd_codigo")
					cd_matricula = rs("cd_matricula")
					nm_nome = rs("nm_nome")
					nm_sobrenome = rs("nm_sobrenome")
					cd_unidades = rs("cd_unidades")
					dt_demissao = rs("dt_demissao")
					cd_status = rs("cd_status")
					%>
					
					<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#F2F2F2');" bgcolor="#F2F2F2">
						<td>&nbsp;<%=nm_nome%> &nbsp;<%=nm_sobrenome%></td>
						<td><%=dt_mes%></td>
						<td align="right"><font color=<%=cor_saldo%>><%'=total_mes(cd_saldo)%>&nbsp;</font></td>
						<td align="right"><font color=<%=cor_saldo%>><b><%'=total_mes(saldo)%>&nbsp;</b></font></td>
							
					<%'rs.movenext
					nm_lista = nm_nome
					cd_meses = cd_meses + 1
					cor_saldo = ""
						if cd_meses > 3 then
							cd_meses = 1
						end if
					'loop%>
					</tr>
					<%rs.movenext
					cor_font = ""
					Loop%>
			</table>
		<%end if%>
		</td>
	</tr>
</table>
</body>
</html>
