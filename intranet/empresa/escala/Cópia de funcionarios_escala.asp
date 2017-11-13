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
							<option value="<%=cod_unidade%>"><%=nome_unidade%></option>
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
		wend%>
<table align="center" style="border-collapse:collapse;  id="escala"">
	<tr>
		<td align="center" colspan="100%">VD LAP - CIRURGICA LTDA<br>
		ESCALA DE PLANTÃO DOS TÉC. DE VÍDEO DO HOSPITAL <%=nm_unidade%><br>
		XXXXXXXXXX de XXXX</td>
	</tr>
	<tr>
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
		<tr><td colspan="6" id="escala">&nbsp;</td>
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
		
		<%xsql = "Select * From View_funcionario_unidade where cd_unidade = '"&cd_unidade&"'"
		Set rs = dbconn.execute(xsql)
		while not rs.EOF
		cd_funcionario = rs("cd_funcionario")%>
		<tr>
			<%xsql_cargo = "Select * From View_funcionario_horario Where cd_funcionario='"&cd_funcionario&"' AND dt_contratado <= '"&data&"' AND dt_desliga = NULL ORDER BY dt_atualizacao desc"
					Set rs_horario = dbconn.execute(xsql_cargo)
						if not rs_horario.EOF then
							nm_nome = rs_horario("nm_nome")
							cd_numreg = rs_horario("cd_numreg")
							hr_entrada = rs_horario("hr_entrada")
							dt_contratado = rs_horario("dt_contratado")
						end if%>
			<td id="escala"><%=nm_nome%></td>
			<td  id="escala" align="center">C.C.</td>
			<%xsql_cargo = "Select * From View_funcionario_cargo Where cd_funcionario='"&cd_funcionario&"' ORDER BY dt_atualizacao desc"
					Set rs_cargo = dbconn.execute(xsql_cargo)
						if not rs_cargo.EOF then
							nm_cargo = rs_cargo("nm_cargo")
							nm_cargo_abrv = rs_cargo("nm_cargo_abrv")
						rs_cargo.movenext
						end if%>
			<td id="escala"><%=nm_cargo_abrv%></td>
			<td id="escala">T.V.L.</td>
			<td id="escala"><%=cd_numreg%></td>
			<td id="escala"><%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%>às<%=hour(hr_saida)&":"&minute(hr_saida)%></td>
			
		</tr>
		<%rs.movenext
		wend%>
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td id="escala">C.C : Centro Cirurgico<br>
			A.E : Auxiliar de Enfermagem<br>
			T.E : Técnico em Enfermagem<br>
			T.V.L : Técnico em Video Laparoscopia<br>
			F.E : Folga da Enfermagem
			</td>
		</tr>
	</tr>
</table>


</body>
</html>
