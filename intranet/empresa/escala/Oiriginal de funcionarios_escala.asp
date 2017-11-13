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
		xsql = "Select * From View_funcionario where dt_contratado <= '"&month(data)&"/1/"&year(data)&"' AND dt_desliga >='"&month(data)&"/"&fim_mes&"/"&year(data)&"' OR dt_contratado <= '"&month(data)&"/1/"&year(data)&"' AND dt_desliga is null order by nm_nome"
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
			
			func_array = split(relacao_funcionarios,";")
				For Each func_item In func_array
					'*** Verifica quais funcionários trabalham no hospital selecionado ***
					xsql_unidade = "Select * From View_funcionario_unidade Where cd_funcionario="&func_item&" and cd_unidade = "&cd_unidade&" AND dt_atualizacao <= '"&data&"' order by dt_atualizacao asc"
					Set rs_unidade = dbconn.execute(xsql_unidade)
						if not rs_unidade.EOF Then
							if relacao_funcionarios_unidade = "" Then
								relacao_funcionarios_unidade = rs_unidade("cd_funcionario")
							Else
								relacao_funcionarios_unidade = relacao_funcionarios_unidade&";"&rs_unidade("cd_funcionario")
							end if
						end if
				next
				
			contagem = 1
			func_array = split(relacao_funcionarios_unidade,";")
				For Each func_item In func_array
'*****************************************************************
'*** Verifica os líderes que trabalham no hospital selecionado ***
'*****************************************************************
					xsql_horario = "Select * From View_funcionario_horario Where cd_funcionario="&func_item&" and expediente = 540 AND dt_atualizacao <= '"&month(data)&"/1/"&year(data)&"'"
					Set rs_horario = dbconn.execute(xsql_horario)
						if not rs_horario.EOF Then
							if relacao_funcionarios_horario_lider = "" Then
								relacao_funcionarios_horario_lider = rs_horario("cd_funcionario")
							Else
								relacao_funcionarios_horario_lider = relacao_funcionarios_horario_lider&";"&rs_horario("cd_funcionario")
							end if
							
							xsql = "Select * From View_funcionario where cd_codigo = "&rs_horario("cd_funcionario")&" order by nm_nome"
							Set rs_funcionario = dbconn.execute(xsql)
								if not rs_funcionario.EOF then
									nm_nome = rs_funcionario("nm_nome")
										'*** Mostra o horário de trabalho do Líder ***
										xsql_horario = "Select * From View_funcionario_horario Where cd_funcionario="&rs_horario("cd_funcionario")&" and expediente = 540 AND dt_atualizacao <= '"&month(data)&"/1/"&year(data)&"'"
										Set rs_horario = dbconn.execute(xsql_horario)
											if not rs_horario.EOF Then
												funcionario_horario_lider = rs_horario("cd_funcionario")
												hr_entrada = rs_horario("hr_entrada")
												hr_saida = rs_horario("hr_saida")
											end if
										
								end if%>							
							<tr>
							<td id="escala"><%=contagem%></td>
							<td id="escala"><%'=cd_funcionario%><%=nm_nome%></td>
							<td  id="escala" align="center">C.C.</td>
							<td id="escala"><%=nm_cargo_abrv%></td>
							<td id="escala">T.V.L.</td>
							<td id="escala"><%=cd_numreg%></td>
							<td id="escala"><%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%>às<%=zero(hour(hr_saida))&":"&zero(minute(hr_saida))%></td>
						</tr>
						<%contagem = contagem + 1
						end if%>
						
				<%next%>
				<tr><td>&nbsp;</td></tr>
				
			<%func_array = split(relacao_funcionarios_unidade,";")
				For Each func_item In func_array
'****************************************************************************
'*** Verifica quais funcionários trabalham no hospital selecionado em 6hs ***
'****************************************************************************
					xsql_horario = "Select * From View_funcionario_horario Where cd_funcionario="&func_item&" and expediente <> 360 AND dt_atualizacao <= '"&month(data)&"/1/"&year(data)&"'"
					Set rs_horario = dbconn.execute(xsql_horario)
						if not rs_horario.EOF Then
							if relacao_funcionarios_horario_6h = "" Then
								relacao_funcionarios_horario_6h = rs_horario("cd_funcionario")
							Else
								relacao_funcionarios_horario_6h = relacao_funcionarios_horario_6h&";"&rs_horario("cd_funcionario")
							end if
						
						xsql = "Select * From View_funcionario where cd_codigo = '"&rs_horario("cd_funcionario")&"' order by nm_nome"
							Set rs_funcionario = dbconn.execute(xsql)
								if not rs_funcionario.EOF then
									nm_nome = rs_funcionario("nm_nome")
										'*** Mostra o horário de trabalho do funcionário de 6hs ***
										xsql_horario = "Select * From View_funcionario_horario Where cd_funcionario="&rs_horario("cd_funcionario")&" and expediente = 360 AND dt_atualizacao <= '"&month(data)&"/1/"&year(data)&"'"
										Set rs_horario = dbconn.execute(xsql_horario)
											if not rs_horario.EOF Then
												funcionario_horario_lider = rs_horario("cd_funcionario")
												hr_entrada = rs_horario("hr_entrada")
												hr_saida = rs_horario("hr_saida")
											end if
								end if%>							
							<tr>
							<td id="escala"><%=contagem%></td>
							<td id="escala"><%'=cd_funcionario%><%=nm_nome%></td>
							<td  id="escala" align="center">C.C.</td>
							<td id="escala"><%=nm_cargo_abrv%></td>
							<td id="escala">T.V.L.</td>
							<td id="escala"><%=cd_numreg%></td>
							<td id="escala"><%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%>às<%=hour(hr_saida)&":"&minute(hr_saida)%></td>
						</tr>						
						<%contagem = contagem + 1
						end if%>
						
				<%next%>
				<tr><td colspan="100%">&nbsp;</td></tr>
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
	
			<%func_array = split(relacao_funcionarios_unidade,";")
				For Each func_item In func_array
'********************************************************************************
'*** Verifica quais funcionários trabalham no hospital selecionado em 12/36hs ***
'********************************************************************************
					xsql_horario = "Select * From View_funcionario_horario Where cd_funcionario="&func_item&" and expediente = 720 AND dt_atualizacao <= '"&month(data)&"/1/"&year(data)&"'"
					Set rs_horario = dbconn.execute(xsql_horario)
						if not rs_horario.EOF Then
							if relacao_funcionarios_horario_12h = "" Then
								relacao_funcionarios_horario_12h = rs_horario("cd_funcionario")
							Else
								relacao_funcionarios_horario_12h = relacao_funcionarios_horario_12h&";"&rs_horario("cd_funcionario")
							end if
						
						xsql = "Select * From View_funcionario where cd_codigo = '"&rs_horario("cd_funcionario")&"' order by nm_nome"
							Set rs_funcionario = dbconn.execute(xsql)
								if not rs_funcionario.EOF then
									nm_nome = rs_funcionario("nm_nome")
										xsql_horario = "Select * From View_funcionario_horario Where cd_funcionario="&rs_horario("cd_funcionario")&" and expediente = 720 AND dt_atualizacao <= '"&month(data)&"/1/"&year(data)&"'"
										Set rs_horario = dbconn.execute(xsql_horario)
											if not rs_horario.EOF Then
												funcionario_horario_lider = rs_horario("cd_funcionario")
												hr_entrada = rs_horario("hr_entrada")
												hr_saida = rs_horario("hr_saida")
											end if
								end if%>							
							<tr>
							<td id="escala"><%=contagem%></td>
							<td id="escala"><%'=cd_funcionario%><%=nm_nome%></td>
							<td  id="escala" align="center">C.C.</td>
							<td id="escala"><%=nm_cargo_abrv%></td>
							<td id="escala">T.V.L.</td>
							<td id="escala"><%=cd_numreg%></td>
							<td id="escala"><%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%>às<%=hour(hr_saida)&":"&minute(hr_saida)%></td>
						</tr>
						<%contagem = contagem + 1
						end if
						
				next
		'nm_nome = rs("nm_nome")
		'cd_numreg = rs("cd_numreg")
		
		
		
		
		'*** verifica se o funcionario trabalhou no hospital no periodo solicitado ***
		'xsql_unidade = "Select * From View_funcionario_unidade Where cd_funcionario="&cd_funcionario&""
		'	Set rs_unidade = dbconn.execute(xsql_unidade)
		'		if not rs_unidade.EOF then
		'			und_funcionario = rs_unidade("cd_funcionario")
		'		end if
					
					'*** mostra o funcionario caso seja do hospital solicitado ***
		'			if und_funcionario = cd_funcionario then
					
					'*** Mostra o Lider (exp: 8hs) ***
					
					
					'xsql = "Select * From View_funcionario where cd_codigo = "&cd_funcionario&" order by nm_nome"
					'Set rs_funcionario = dbconn.execute(xsql)
					'	if not rs_funcionario.EOF then
					'		nm_nome = rs_funcionario("nm_nome")
					'	end if%>
						<!--tr>
							<td id="escala"><%=contagem%></td>
							<td id="escala"><%'=cd_funcionario%><%=nm_nome%></td>
							<td  id="escala" align="center">C.C.</td>
							<td id="escala"><%=nm_cargo_abrv%></td>
							<td id="escala">T.V.L.</td>
							<td id="escala"><%=cd_numreg%></td>
							<td id="escala"><%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%>às<%=hour(hr_saida)&":"&minute(hr_saida)%></td>
						</tr-->
					
					<%contagem = contagem + 1%>
					<%'end if
		
		'rs.movenext
		'wend%>
		<tr><td colspan="100%">&nbsp;</td></tr>
		<tr>
			<td></td>
			<td id="escala">C.C : Centro Cirurgico<br>
			A.E : Auxiliar de Enfermagem<br>
			T.E : Técnico em Enfermagem<br>
			T.V.L : Técnico em Video Laparoscopia<br>
			F.E : Folga da Enfermagem
			</td>
		</tr>
		<tr><td colspan="100%">Todos: <%'=relacao_funcionarios%></td></tr>
		<tr><td colspan="100%">Unidade: <%=relacao_funcionarios_unidade%></td></tr>
		<tr><td colspan="100%">Lider: <%=relacao_funcionarios_horario_lider%></td></tr>
		<tr><td colspan="100%">F06h: <%=relacao_funcionarios_horario_6h%></td></tr>
		<tr><td colspan="100%">F12h: <%=relacao_funcionarios_horario_12h%></td></tr>
	</tr>
</table>
<%end if%>

</body>
</html>
