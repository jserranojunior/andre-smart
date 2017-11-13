<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/numero_extenso.asp"-->
<!--#include file="../includes/util.asp"-->
<!--#include file="../css/geral.htm"-->
<html>
<head>
	<title>::Faturamento::..Lista..</title>
</head>
<%cd_user = session("cd_codigo")

modo = request("modo")
dt_dia = request("dt_dia")
dt_mes = request("dt_mes")
dt_ano = request("dt_ano")
	'dt_vencimento = zero(dt_dia)&"/"&zero(dt_mes)&"/"&dt_ano
	dt_vencimento = zero(dt_dia)&"/"&zero(dt_mes)&"/"&dt_ano
	
data_atual = zero(month(now))&"/"&zero(day(now))&"/"&year(now)
	
nm_valor = request("nm_valor")%>
<body>
<table border="1" style="border-collapse:collapse;">
	<tr>
		<td align="center" colspan="3"  style="color:white; background-color:gray;"><b>LISTAGEM</b>: <%=mes_selecionado(dt_mes)&"/"&dt_ano%><%'=cd_user%> </td>
	</tr>
	<tr><td align="center" colspan="3">Valores atrasados</td></tr>
	<tr  style="color:gray; background-color:silver;">
		<td align="center"><b>Unidade</b></td>
		<td align="center"><b>Valores</b></td>
		<td>&nbsp;</td>
	</tr>
	
	<%if modo = "atrasados" then
		strsql = "SELECT * FROM TBL_financeiro_fatura_adia_pgto WHERE dt_vencimento_ori BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"'"
	elseif modo = "receb_atraso" then
		strsql = "SELECT * FROM TBL_financeiro_fatura_adia_pgto WHERE dt_vencimento_novo BETWEEN '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"'"
	end if
		SET rs_lista = dbconn.execute(strsql)
			while not rs_lista.EOF
				cd_atrasadas = rs_lista("cd_codigo")
				nm_valor = rs_lista("nm_valor")
				total = int(total+nm_valor)
				cd_unidade = rs_lista("cd_unidade")%>
	<tr>
		<%xsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&""
		Set rs = dbconn.execute(xsql)
		while not rs.EOF
			strcd_unidade = rs("cd_codigo")
			strnm_unidade = rs("nm_unidade")
			strnm_sigla = rs("nm_sigla")
		rs.movenext
		wend%>
		<td><b><%=strnm_unidade%></b><br><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
		<td align="right"><%=formatnumber(nm_valor,2)%></td>
		<td><img src="../imagens/ic_editar.gif" alt="" width="13" height="14" border="0" onClick="window.open('fatura_recebimento_adiamento.asp?cd_atrasadas=<%=cd_atrasadas%>','adiar','location=0,status=0,width=300,height=200,scrollbars=1resizable')"></td>
	</tr>
	<%rs_lista.movenext
	wend%>
	<tr><td colspan="3"><img src="../imagens/blackdot.gif" alt="" width="250" height="1" border="0"></td></tr>	
	<tr style="color:white; background-color:gray;">
		<td><b>Total</b></td>
		<td align="right"><b><%=formatnumber(total,2)%></b></td>
		<td>&nbsp;</td>
	</tr>
</table>
</body>
</html>
