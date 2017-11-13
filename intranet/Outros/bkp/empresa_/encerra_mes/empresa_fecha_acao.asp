<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!--include file="../../empresa/horas_trabalhadas/horas_trab_horas.asp"--><!--?mes_sel=3&ano_sel=2008&cd_codigo=97&tela=0&cd_origem=2"-->
<%
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")
cd_origem = request("cd_origem")
'acao = request("acao")

	if mes_sel < 10 Then
	'mes_sel = 0 & mes_sel
	mes_selec = 0 & mes_sel
	End if

data_atual = ano_sel&mes_selec

%>
<%


	'******************************
	'***** Encerra o mês em aberto ****
	'******************************	
	xsql = "up_Fechamento_mes @cd_mes="&mes_sel&", @cd_ano="&ano_sel&", @cd_origem="&cd_origem&", @nm_fechamento=1"
	dbconn.execute(xsql)
	
'	response.redirect("../../emp_fecha_mes.asp")

					

%>
<table border="0" width="100%">
	<tr><td colspan="100%"><img src="../../imagens/blackdot.gif" width=500 height=0></td></tr>
	<tr>
		<td align="right">Data</td>
		<td align="right">Origem</td>
		<td align="right">Cod</td>
		<td>Nome</td>
		<td align="right">saldo</td>
	</tr>
	
				<%
					'***************************************************************************
					'**** Lista os funcionarios um a um para obter o saldo do mes em horas *****
					'***************************************************************************
					'condicao = " Where  cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND dt_demissao >= "&data_atual&" OR cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND dt_demissao = null "
					condicao = " Where  cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND dt_desliga >= "&data_atual&" OR cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND dt_desliga = null "
					xsql_ver = "up_funcionarios_lista @cd_codigo='', @condicao='"&condicao&"', @ordem= 'nm_nome, nm_sobrenome'"
					'xsql_ver = "up_funcionarios_lista @cd_codigo = 97, @condicao='"&condicao&"', @ordem= 'nm_nome, nm_sobrenome'"
					Set rs_ver = dbconn.execute(xsql_ver)
					
					Do while NOT rs_ver.EOF
					
					cd_codigo = rs_ver("cd_codigo")
					nm_nome = rs_ver("nm_nome")
					nm_sobrenome = rs_ver("nm_sobrenome")%>
	<tr id="no_print">
		<td id="no_print" colspan="100%">
			<!--#include file="../../empresa/horas_trabalhadas/horas_trab_cartao.asp"-->
		</td>
	</tr>
	<tr id="no_print><td colspan="100%"><!--#include file="../../empresa/banco_horas/banco_horas_acao_teste.asp"-->	</td></tr>
	<tr id="ok_print">
		<td align="right"><%=mes_sel%>/<%=ano_sel%></td>
		<td align="right"><%=cd_origem%></td>
		<td align="right"><%=cd_codigo%></td>
		<td align="left"><%=nm_nome%>&nbsp;<%=nm_sobrenome%></td>
		<td align="right"><%=total_mes(excedente)%></td>
	</tr>
	

<%
excedente = ""
rs_ver.movenext
loop

response.redirect("../../emp_fecha_mes.asp")%>
</table>