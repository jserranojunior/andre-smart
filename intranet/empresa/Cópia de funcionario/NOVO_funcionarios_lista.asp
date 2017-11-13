<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<%
dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

if dia_sel = "" or mes_sel = "" or ano_sel = "" then
	dia_sel = day(now)
	mes_sel = month(now)
	ano_sel = year(now)
end if

data_selecionada = dia_sel&"/"&mes_sel&"/"&ano_sel

'data_selecionada = request("data_selecionada")
'data_selecionada = month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)
%>
<head>
	<title>Untitled</title>
</head>

<body><br>

<table align="center" border="1">
	<form action="../../empresa.asp" method="post">
	<input type="hidden" name="tipo" value="lista">
	<tr>
		<td colspan="100%">
		<input type="text" name="dia_sel" size="2" maxlength="2" value="<%=zero(day(data_selecionada))%>">
		<input type="text" name="mes_sel" size="2" maxlength="2" value="<%=zero(month(data_selecionada))%>">
		<input type="text" name="ano_sel" size="4" maxlength="4" value="<%=year(data_selecionada)%>"> &nbsp; <input type="submit" name="OK" value="OK">
			
			
			
		
		</td>
	</tr>
	</form>
	<!--tr>
		<td colspan="100%">Lista todos os funcionarios presentes no banco de dados</td>
	</tr-->
	<tr>
	<%'*** Lista TODOS os funcionarios do banco de dados ***
	xsql = "Select * FROM TBL_funcionario order by nm_nome "
	Set rs = dbconn.execute(xsql)
		while not rs.EOF
			lista_funcionario = lista_funcionario&";"&rs("cd_codigo")'&"<br>"
			rs.movenext
		wend%>
		<td colspan="100%"><%'=lista_funcionario%></td>		
	</tr>	
	<%'*** lista todas as empresas do banco de dados ***
	xsql = "Select * FROM TBL_tipo_contrato order by cd_codigo"
	Set rs = dbconn.execute(xsql)
		while not rs.EOF
			
			'*** Lista cada empresa e separa os respectivos funcionarios pela data solicitada***
			cd_empresa_contratante = rs("cd_codigo")
			nm_regime_trab = rs("nm_regime_trab")%>	
								
			<%'*** Lista os funcionarios ***************************
				conta_func = 0
				
				'lista_funcionario = mid(lista_funcionario,2,len(lista_funcionario))
				funcionario_array = split(lista_funcionario,";")
					For Each funcionario_item In funcionario_array%>
						
						<%''xsql = "SELECT top 1 * FROM View_funcionario_contrato_lista WHERE cd_funcionario = '"&funcionario_item&"' AND dt_admissao <= '"&day(data_selecionada)&"/"&month(data_selecionada)&"/"&year(data_selecionada)&"' AND dt_demissao >= '"&data_selecionada&"' AND cd_contrato = 1 OR cd_funcionario = '"&funcionario_item&"' AND dt_admissao <= '"&data_selecionada&"' AND dt_demissao IS NULL AND cd_contrato = 1 ORDER BY dt_admissao DESC"
						xsql = "SELECT top 1 * FROM View_funcionario_contrato_lista WHERE cd_funcionario = '"&funcionario_item&"' AND dt_admissao <= '"&month(data_selecionada)&"/"&day(data_selecionada)&"/"&year(data_selecionada)&"' AND cd_contrato = "&cd_empresa_contratante&" ORDER BY dt_admissao DESC"
						Set rs_func = dbconn.execute(xsql)
							
							if not rs_func.EOF then
								conta_func = conta_func + 1
								'lista_empresa = zero(conta_func)&"-<b>"&rs_func("cd_funcionario")&"</b> - ("&zero(day(rs_func("dt_admissao")))&"/"&zero(month(rs_func("dt_admissao")))&"/"&year(rs_func("dt_admissao"))&") - "&rs_func("nm_nome")&"<br>"
								
								cd_funcionario = rs_func("cd_funcionario")
								nm_funcionario = rs_func("nm_nome")
								dt_admissao = zero(day(rs_func("dt_admissao")))&"/"&zero(month(rs_func("dt_admissao")))&"/"&year(rs_func("dt_admissao"))
								
								
								xsql_regprof = "Select * From TBL_funcionario Where cd_codigo = "&cd_funcionario&""' ORDER BY dt_atualizacao desc"
								Set rs_regprof = dbconn.execute(xsql_regprof)
									if not rs_regprof.EOF then
										cd_numreg = rs_regprof("cd_numreg")
									end if
								
								
								xsql_cargo = "Select * From View_funcionario_cargo Where cd_funcionario = "&cd_funcionario&" ORDER BY dt_atualizacao desc"
								Set rs_cargo = dbconn.execute(xsql_cargo)
									if not rs_cargo.EOF then
										nm_cargo = rs_cargo("nm_cargo")
									end if
							
								xsql = "Select * From View_funcionario_horario Where cd_funcionario = "&cd_funcionario&" order by dt_atualizacao DESC"
								Set rs_hora = dbconn.execute(xsql)
									if not rs_hora.EOF Then
										hr_entrada = zero(hour(rs_hora("hr_entrada")))&":"&zero(minute(rs_hora("hr_entrada")))
										hr_saida = zero(hour(rs_hora("hr_saida")))&":"&zero(minute(rs_hora("hr_saida")))
										nm_intervalo = rs_hora("nm_intervalo")
									end if
									
								xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&cd_funcionario&" order by dt_atualizacao DESC"
								Set rs_unid = dbconn.execute(xsql)
									if not rs_unid.EOF Then
										nm_unidade = rs_unid("nm_unidade")
										nm_sigla = rs_unid("nm_sigla")
									end if
									
								xsql = "Select * From View_funcionario_status Where cd_funcionario = "&cd_funcionario&" order by dt_atualizacao DESC"
								Set rs_status = dbconn.execute(xsql)
									if not rs_status.EOF then
										nm_status = rs_status("nm_status")
									end if
								
								
								'*** Mostra o cabeçalho
									if cabeca = "" then%>
									<!--tr><td><%=lista_funcionario%></td></tr-->
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan = "100%"><%=cd_empresa_contratante%> - <%=nm_regime_trab%></td>								
									</tr>
									<tr bgcolor=#c0c0c0>
										<td>&nbsp;</td>
										<td>&nbsp;<b>Funcionário</b></td>
										<td>&nbsp;<b>Reg. Profissional</b></td>
										<td>&nbsp;<b>Cargo</b></td>
										<td>&nbsp;<b>Horário</b></td>
										<td>&nbsp;<b>Unidade</b></td>
										<td>&nbsp;<b>Situação</b></td>
										<td>&nbsp;<b>Admissão</b></td>
										<td>&nbsp;<b>demissão</b></td>
									</tr>
									<%end if%>
								
								<%'*** Lista os funcionarios ***%>
								<!--tr-->
								<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs_func("cd_funcionario")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px;" bgcolor="<%=cor_linha%>">
									<td><%=zero(conta_func)%></td>
									<td>&nbsp;<%=nm_funcionario%></td>
									<td>&nbsp;<%=cd_numreg%></td>
									<td>&nbsp;<%=nm_cargo%></td>
									<td>&nbsp;<%=hr_entrada%> - <%=hr_saida%></td>
									<td>&nbsp;<%=nm_sigla%></td>
									<td>&nbsp;<%=nm_status%></td>
									<td>&nbsp;<%=dt_admissao%></td>
									<td>&nbsp;<%=dt_demissao%></td>
								</tr>								
								<%
								'*** retira o funcionários da lista de remanescentes ***
								funcionario_item = ""
								cabeca = 1
							
							end if
					
					if funcionario_item = "" then
						remanescentes = remanescentes
					else
						remanescentes = remanescentes&";"&funcionario_item
					end if%>				
					<%next
					cabeca = ""
					
					lista_empresa = ""
'*******************************************************%>
							<!--tr>
								<td><%=lista_empresa%>...</td>
							</tr-->
							
					
						

		<%'next
		rs.movenext
		wend%>
		<tr>
			<td><%'=lista_funcionario%><%'=remanescentes%></td>
		</tr>
		<td></td>
	</tr>
</table>


</body>
</html>
