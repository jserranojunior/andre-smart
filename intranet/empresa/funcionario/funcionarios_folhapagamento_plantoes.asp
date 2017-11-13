<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Untitled</title>
<%cd_user = session("cd_codigo")
pasta_loc = "empresa\funcionario\"
arquivo_loc = "funcionarios_folhapagamento_plantoes.asp"

mens = request("mens")
	mens = replace(mens,"_","&nbsp;")
'dt_dia = request("dt_dia")
dt_mes = request("dt_mes")
dt_ano = request("dt_ano")
	'if dt_dia = "" AND  dt_mes = "" AND dt_ano = "" Then
	if dt_mes = "" AND dt_ano = "" Then
		dt_dia = day(now)
		dt_mes = month(now)
		dt_ano = year(now)
		
			if dt_dia > 10 then
				dt_mes = dt_mes + 1
				dt_ano = dt_ano
					if dt_mes > 12 then
						dt_mes = 1
						dt_ano = dt_ano + 1
					end if
			end if
		
	end if

%>	
	
</head>

<script language="Javascript">
<!-- 
//-------------------------------------------------------------------
function muda_mes() {
	window.rcm.submit();
}

//--------------------------------------------------------------------
-->
</script>
<body>
<!--# include file="../includes/arquivo_loc.asp"-->
<%'*** carrega as informações dos plantões ***

if(dt_mes = 12)then
	dt_mes_gravar = 1
	dt_ano_gravar = dt_ano + 1
	else 
	dt_mes_gravar = dt_mes + 1
	dt_ano_gravar = dt_ano
	end if



	strsql = "SELECT * FROM TBL_funcionario_plantoes WHERE dt_atualizacao='"&dt_mes_gravar&"/1/"&dt_ano_gravar&"'"
		Set rs_plantoes = dbconn.execute(strsql)
		while not rs_plantoes.EOF
			cd_plantoes = rs_plantoes("cd_codigo")
			nr_plantoes_6h = rs_plantoes("nr_plantoes_6h")
			nr_plantoes_8h = rs_plantoes("nr_plantoes_8h")
			nr_plantoes_12h = rs_plantoes("nr_plantoes_12h")
		rs_plantoes.movenext
		wend%>
<table align="center" border="0" style="border-collapse:collapse;">
	<tr>
		<td align="center" colspan="4">Dias trabalhados <%=dt_dia&"/"&dt_mes&"/"&dt_ano%></td>
	</tr>
	<form action="enfermagem.asp?tipo=qtd_plantoes" method="post" name="rcm" id="rcm">
	<input type="hidden" name="acao" value="plantoes">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<tr>
		<td align="center" style="background-color:silver;"><b>Período:</b><br><img src="../../imagens/px.gif" alt="" width="60" height="0" border="0"></td>
		<td align="left" style="background-color:silver;" colspan="3">
			<select name="dt_mes" <%=styleinput%> onchange="muda_mes();">
				<%for i = 1 to 12%>
					<option value="<%=i%>" <%if i = int(dt_mes) then%>SELECTED<%end if%>><%=mesdoano(i)%></option>
				<%next%>
			</select> / <input type="text" name="dt_ano" value="<%=dt_ano%>" size="4" maxlength="4"><input type="submit" name="OK" value="OK"></td>
	</tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	</form>
	<form action="empresa/acoes/rcm_acao.asp" method="post" name="rcm2" id="rcm2">
	<input type="hidden" name="acao" value="plantoes">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="dt_dia" value="1">
	<input type="hidden" name="dt_mes" value="<%=dt_mes_gravar%>">
	<input type="hidden" name="dt_ano" value="<%=dt_ano_gravar%>">
	<tr><td align="center" colspan="4" bgcolor="#c0c0c0"><img src="../../imagens/px.gif" alt="" width="1" height="3" border="0"></td></tr>
	<tr>
		<td align="center" rowspan="5" bgcolor="#c0c0c0"><b>Carga<br>diária</b></td>
		<td align="center"><b>06h</b></td>
		<td><input type="text" name="nr_plantoes_6h" size="2" maxlength="2" value="<%=nr_plantoes_6h%>"></td>
		<td align="center" rowspan="5" bgcolor="#c0c0c0"><input type="submit" value="Grava"></td>
	</tr>
	<tr><td colspan="2" bgcolor="#c0c0c0"><img src="../../imagens/px.gif" alt="" width="1" height="3" border="0"></td></tr>
	<tr>
		<td align="center"><b>08h</b></td>
		<td><input type="text" name="nr_plantoes_8h" size="2" maxlength="2" value="<%=nr_plantoes_8h%>"></td>
	</tr>
	<tr><td colspan="2" bgcolor="#c0c0c0"><img src="../../imagens/px.gif" alt="" width="1" height="3" border="0"></td></tr>
	<tr>
		<td align="center"><b>12h</b></td>
		<td><input type="text" name="nr_plantoes_12h" size="2" maxlength="2" value="<%=nr_plantoes_12h%>"></td>
	</tr>
	<tr><td align="center" colspan="4" bgcolor="#c0c0c0"><img src="../../imagens/px.gif" alt="" width="1" height="3" border="0"></td></tr>
	</form>
</table>

<br>
<p style="text-align:center;"><%=mens%></p>
<table align="center">
<form action="empresa/acoes/rcm_acao.asp" method="post" name="dozehoras" id="dozehoras">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="dt_dia" value="1">
	<input type="hidden" name="dt_mes" value="<%=dt_mes%>">
	<input type="hidden" name="dt_ano" value="<%=dt_ano%>">
	<%if cd_user = "46" then%><tr><td colspan="3" align="center"><i style="color:red;">funcionarios_folhapagamento_plantoes.asp</i></td></tr><%end if %>
    <tr style="background-color:gray; color:white;">
		<td colspan="3" align="center">Acrescenta mais 1 dia de trabalho no mês (12 horas). <%'=cd_user%></td>
	</tr>
	<%acao_1d = "plantoes_1d_insert"
	acao_1d_botao = "incluir"
	
	
	
	
	
	
	xsql = "SELECT * FROM TBL_funcionario_plantoes_1d WHERE dt_data = '"&dt_mes&"/1/"&dt_ano&"'"
	SET rs = dbconn.execute(xsql)
		if not rs.EOF then
			cd_plantoes_1d = rs("cd_codigo")
			selecao_func = rs("nm_funcionarios")
			acao_1d = "plantoes_1d_update"
			acao_1d_botao = "Alterar"
		end if
		
	
	num_linha = 1
	strsql = "SELECT DISTINCT cd_funcionario, dt_atualizacao, nm_nome, cd_contrato, expediente, dt_admissao, dt_demissao FROM VIEW_funcionario_horario WHERE (expediente = 720) AND (dt_admissao <= '"&dt_mes&"/1/"&dt_ano&"') AND (dt_demissao >= '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"') OR (expediente = 720) AND (dt_admissao <= '"&dt_mes&"/1/"&dt_ano&"') AND (dt_demissao IS NULL) ORDER BY nm_nome, dt_atualizacao,cd_contrato, expediente, dt_admissao, dt_demissao DESC, cd_funcionario"
			SET rs = dbconn.execute(strsql)
			
			while not rs.EOF
				cd_funcionario = rs("cd_funcionario")
					ver_func = instr(1,selecao_func,cd_funcionario,1)
				nm_nome = rs("nm_nome")
				expediente = rs("expediente")
				%>
				
	<tr <%if ver_func > 0 then%>bgcolor="#ccffcc"<%end if%>>
		<td align="right"><%=num_linha%></td>
		<td><input type="checkbox" name="selecao_func" value="<%=cd_funcionario%>" <%if ver_func > 0 then%>checked<%end if%>></td>
		<td><%=nm_nome%> <%'=expediente/60%></td>
	</tr>
			<%num_linha = num_linha + 1
			rs.movenext
			wend%>
	<tr><td align="center" colspan="3"><input type="submit" name="aplicar" value="<%=acao_1d_botao%>"></td></tr>
	<input type="hidden" name="cd_plantoes_1d" value="<%=cd_plantoes_1d%>">
	<input type="hidden" name="acao" value="<%=acao_1d%>">
</form>
</table>
</body>
</html>
