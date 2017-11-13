<!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->

<html>
<head>
	<title>Untitled</title>
</head>
<style type="text/css">
<!--
#escala{ border: 1px solid black; border-collapse:collapse;}  
-->
</style>
<%
tipo = request("tipo")
dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")
cd_unidade = request("cd_unidade")



if dia_sel = "" OR mes_sel = "" OR ano_sel = "" Then
	data = day(now)&"/"&month(now)&"/"&year(now)
Else
	data = dia_sel&"/"&mes_sel&"/"&ano_sel
end if

fim_mes = ultimodiames(Month(data),year(data))
%>

<body><br>


<table align="center" style="border-collapse:collapse;  id="escala"">
	<tr><td colspan="100%">Seleção para Escala de plantões</td></tr>
	<form action="empresa.asp" name="Form" id="Form">
		<input type="hidden" name="tipo" value="<%=tipo%>">
		<input type="hidden" name="dia_sel" value="1">
	<tr>
		<td>Mês:</td>
		<td><input type="text" name="mes_sel" size="2" maxlength="2" value="<%=month(data)%>"></td>
		<td>Ano:</td>
		<td><input type="text" name="ano_sel" size="4" maxlength="4" value="<%=year(data)%>"></td>
	</tr>
	<tr>
		<td>Hospital:</td>
		<td colspan="3"><select name="cd_unidade">
			<%xsql = "Select * From TBL_unidades where cd_status = 1 order by nm_unidade"
				Set rs_unid = dbconn.execute(xsql)
					while not rs_unid.EOF
							cod_unidade = rs_unid("cd_codigo")
							nome_unidade = rs_unid("nm_unidade")
							nome_sigla = rs_unid("nm_sigla")%>
							<option value="<%=cod_unidade%>" <%if int(cd_unidade) = int(cod_unidade) then response.write("SELECTED") end if%>><%=nome_unidade%></option>
					<%rs_unid.movenext
					wend%>
			</select>
		</td>
	</tr>
	<tr>
		<td><input type="submit" name="escala" value="Mostra"></td>
	</tr>
	</form>
</table><br>
<%xsql = "Select * From TBL_unidades where cd_codigo = '"&cd_unidade&"'"
	Set rs_unid = dbconn.execute(xsql)
		while not rs_unid.EOF
			'cd_unidade = rs_unid("cd_codigo")
			nm_unidade = rs_unid("nm_unidade")
			nm_sigla = rs_unid("nm_sigla")%>
		<%rs_unid.movenext
		wend

if cd_unidade <> "" Then%>
<table border="0" align="center" style="border-collapse:collapse; border: 0px solid black;">
	<tr>
		<td align="center" colspan="100%">VD LAP - CIRURGICA LTDA<br>
		ESCALA DE PLANTÃO DOS TÉC. DE VÍDEO DO HOSPITAL <%=nm_unidade%><br>
		XXXXXXXXXX de XXXX</td>
	</tr>
	<tr>
		<td></td>
		<td id="escala">Colaboradores</td>
		<td id="escala">Setor</td>
		<td id="escala">Cargo</td>
		<td id="escala">Função</td>
		<td id="escala">Coren</td>
		<td id="escala">Horário</td>
		<%for i=1 to fim_mes%>
			<td id="escala">
				<%=i%>
			</td>
		<%next%>
		<tr><td colspan="7" id="escala">&nbsp;</td>
		<%inicio_mes = 1
		dia_semana = weekday("1/"&month(data)&"/"&year(data))
		 for i=1 to fim_mes%>
			<td id="escala" id="escala">
				<%=ucase(left(weekdayname(dia_semana),1))%>
			<%dia_semana = dia_semana + 1
			if dia_semana > 7 Then
				dia_semana = 1
			end if%>
			</td>
		<%next%>
	</tr>		
		<%'*** Lista e separa todos os funcionarios ativos na data selecionada ***
		xsql = "Select * From View_funcionario_contrato_lista where dt_admissao <= '"&month(data)&"/1/"&year(data)&"' AND dt_demissao >='"&month(data)&"/"&fim_mes&"/"&year(data)&"' OR dt_admissao <= '"&month(data)&"/1/"&year(data)&"' AND dt_demissao is null order by nm_nome"
		Set rs = dbconn.execute(xsql)
			while not rs.EOF
				'cd_funcionario = rs("cd_codigo")
				
				if relacao_funcionarios = "" Then
					relacao_funcionarios =  rs("cd_codigo")
				else
					relacao_funcionarios =  relacao_funcionarios &";"& rs("cd_codigo")
				end if
				'response.write(relacao_funcionarios)
			rs.movenext
			wend
			
			
		'*** Lista e separa os lideres ativos na data selecionada do horario 7:00hs as 16:00hs***
		'xsql = "Select * From View_funcionario_horario where dt_contratado <= '"&month(data)&"/1/"&year(data)&"' AND dt_desliga >='"&month(data)&"/"&fim_mes&"/"&year(data)&"' and {fn HOUR(hr_entrada)} = '7' OR dt_contratado <= '"&month(data)&"/1/"&year(data)&"' AND dt_desliga is null and {fn HOUR(hr_entrada)} = '7'"' order by nm_nome"
		'xsql = "Select * From View_funcionario_horario where dt_contratado <= '"&month(data)&"/1/"&year(data)&"' AND dt_desliga >='"&month(data)&"/"&fim_mes&"/"&year(data)&"' and {fn HOUR(hr_entrada)} = '7' OR dt_contratado <= '"&month(data)&"/1/"&year(data)&"' AND dt_desliga is null and {fn HOUR(hr_entrada)} = '7'"' order by nm_nome"
		xsql = "Select * From View_funcionario_horario "
		Set rs = dbconn.execute(xsql)
			while not rs.EOF
				'cd_funcionario = rs("cd_codigo")
				
				if relacao_funcionarios_lider_1 = "" Then
					relacao_funcionarios_lider_1 =  rs("cd_codigo")
				else
					relacao_funcionarios_lider_1 =  relacao_funcionarios_lider_1 &";"& rs("cd_codigo")
				end if
				'response.write(relacao_funcionarios)
			rs.movenext
			wend
			
		'******************************************************************************************
		'*** Lista e separa os lideres ativos na data selecionada do horario 13:00hs as 22:00hs ***
		'xsql = "Select * From View_funcionario_horario where dt_atualizacao <= '"&month(data)&"/1/"&year(data)&"' and dt_contratado <= '"&month(data)&"/1/"&year(data)&"' AND dt_desliga >='"&month(data)&"/"&fim_mes&"/"&year(data)&"' and {fn HOUR(hr_entrada)} = '13' OR dt_atualizacao <= '"&month(data)&"/1/"&year(data)&"' and dt_contratado <= '1/"&month(data)&"/"&year(data)&"' AND dt_desliga is null and {fn HOUR(hr_entrada)} = '13'"' order by nm_nome"
		xsql = "Select * From View_funcionario_horario where dt_atualizacao <= '"&month(data)&"/1/"&year(data)&"'"
		Set rs = dbconn.execute(xsql)
			while not rs.EOF
				'cd_funcionario = rs("cd_codigo")
				
				if relacao_funcionarios_lider_2 = "" Then
					relacao_funcionarios_lider_2 =  rs("cd_codigo")
				else
					relacao_funcionarios_lider_2 =  relacao_funcionarios_lider_2 &";"& rs("cd_codigo")
				end if
				'response.write(relacao_funcionarios)
			rs.movenext
			wend
		%>

		<tr><td colspan="100%">&nbsp;</td></tr>		
		<tr><td colspan="100%">Todos: <%'=relacao_funcionarios%></td></tr>
		<tr><td colspan="100%">Unidade: <%=relacao_funcionarios_unidade%></td></tr>
		<tr><td colspan="100%">Lider 07h: <%=relacao_funcionarios_lider_1%></td></tr>
		<tr><td><td>&nbsp;</td><td colspan="6">
		<%func_array = split(relacao_funcionarios_lider_1,";")
			For Each func_item In func_array
		
				xsql = "Select * From View_funcionario where cd_codigo = '"&func_item&"'"
				Set rs = dbconn.execute(xsql)
					while not rs.EOF
						nm_nome = rs("nm_nome")&"<br>"
						response.write(nm_nome)
					rs.movenext
					wend
			next%><br></td></tr>
		<tr><td colspan="100%">Lider 13h: <%=relacao_funcionarios_lider_2%></td></tr>
		<tr><td><td>&nbsp;</td><td colspan="6">
		<%func_array = split(relacao_funcionarios_lider_2,";")
			For Each func_item In func_array
		
				xsql = "Select * From View_funcionario where cd_codigo = '"&func_item&"'"
				Set rs = dbconn.execute(xsql)
					while not rs.EOF
						nm_nome = rs("nm_nome")&"<br>"
						response.write(nm_nome)
					rs.movenext
					wend
			next%><br></td></tr>
		<tr><td colspan="100%">Hs: <%=relacao_funcionarios_lider_hr%></td></tr>
		<tr><td colspan="100%">F12h: <%=relacao_funcionarios_horario_12h%></td></tr>
	</tr>
</table>
<%end if%>

</body>
</html>
