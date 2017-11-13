<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<head>
	<title>Mapa de utilização</title>
</head>
<%
dia_anterior = request("dia_anterior")
dia_posterior = request("dia_posterior")

dia_i = request("dia_i")
mes_i = request("mes_i")
ano_i = request("ano_i")
	if dia_anterior = 1 then
		dia_i = dia_i-1
			if dia_i < 1 then
				dia_i = ultimodiames(mes_i-1,ano_i)
				mes_i = mes_i - 1
					if mes_i < 1 then
						mes_i = 12
						ano_i = ano_i - 1
					end if
			end if
	end if
	if dia_posterior = 1 then
		dia_i = dia_i+1
			if dia_i > ultimodiames(mes_i,ano_i) then
				dia_i = 1
				mes_i = mes_i + 1
					if mes_i > 12 then
						mes_i = 1
						ano_i = ano_i + 1
					end if
			end if
	end if


	if dia_i = "" or mes_i = "" or ano_i = "" Then
		dia_i = day(now)
		mes_i = month(now)
		ano_i = year(now)
	end if

'dia_f = request("dia_f")
'mes_f = request("mes_f")
'ano_f = request("ano_f")
'	if dia_f = "" or mes_f = "" or ano_f = "" Then
'		dia_f = ultimodiames(month(now),year(now))
'		mes_f = month(now)
'		ano_f = year(now)
'	end if

cd_unidade = request("cd_unidade")
%>
<body>
	<table align="center" style="" border="1" id="no_print">
	<form action="../protocolo.asp" name="form" id="form">
	<input type="hidden" name="tipo" value="mapa_util_racks">
		<tr>
			<td colspan="3" align="center"><b>Relatório da Gestão de Manutenção</b></td>
		</tr>		
		<%bgc_linha = "FFFFFF"%>
		<tr id="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
			<td><b>Período</b></td>
			<td>
				<select name="dia_i">
					<%numero = 1
					do while NOT numero = 32
					if int(dia_i) = numero Then
					check = "selected"
					end if%>					
					<option value="<%=numero%>" <%=check%>><%=numero%></option>
				<%numero = numero+1
				check = ""
				loop%>
				</select>/
				<select name="mes_i">
					<%numero = 1
					do while NOT numero = 13
					if int(mes_i) = numero Then
					check = "selected"
					end if%>					
					<option value="<%=numero%>" <%=check%>><%=left(mes_selecionado(numero),3)%></option>
				<%numero = numero+1
				check = ""
				loop%>
				</select>/
				<%if ano_i = "" then%><%ano_i=1900%><%end if%>
				<input type="text" name="ano_i" size="4" maxlength="4" value="<%=ano_i%>">
			</td>
		</tr>
		<tr>
			<td><b>Unidade</b></td>
			<td>
			<select name="cd_unidade" class="inputs">
				<option value="0"></option>
					<%strsql = "SELECT * FROM TBL_unidades WHERE cd_codigo <> '6' AND cd_status > '0' Order by 'nm_unidade'"
					Set rs_uni = dbconn.execute(strsql)
					
					while not rs_uni.EOF
					cd_uni = rs_uni("cd_codigo")
					nm_sigla = rs_uni("nm_sigla")
					nome_unidade = rs_uni("nm_unidade")
					
					if int(cd_unidade) = cd_uni Then
					ck_uni = "selected"
					end if%>						
				<option value="<%=cd_uni%>" <%=ck_uni%>><%=nome_unidade%></option>
					<%rs_uni.movenext
					ck_uni = ""
					wend
					%>
			</select>			
		</tr>
		<tr><td><input type="submit" name="OK" value="OK"></td><td align="center"><a href="protocolo.asp?tipo=mapa_util_racks&dia_i=<%=dia_i%>&mes_i=<%=mes_i%>&ano_i=<%=ano_i%>&cd_unidade=<%=cd_unidade%>&dia_anterior=1">dia anterior</a> &nbsp; &nbsp; &nbsp; &nbsp; <a href="protocolo.asp?tipo=mapa_util_racks&dia_i=<%=dia_i%>&mes_i=<%=mes_i%>&ano_i=<%=ano_i%>&cd_unidade=<%=cd_unidade%>&dia_posterior=1">próximo dia</a></td></tr>
	</form>
</table>
<br id="no_print">
<%If cd_unidade <> "" Then

'xsql = "SELECT DISTINCT cd_unidade,nm_sigla_os FROM VIEW_os_lista2 WHERE  (dt_entrada BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"') and fecha between 0 and 0" 
	'Set rs = dbconn.execute(xsql)
		'while not rs.EOF
			'cd_unidade = rs("cd_unidade")
			'nm_sigla_os = rs("nm_sigla_os")
			'lista_und = lista_und &","& cd_unidade%>		
		<%'rs.movenext
		'wend%>

<table border="1" style="border-collapse:collapse;">
	<tr>
		<td colspan="50" align="center">Mapa de utilização dos racks</td>
	</tr>
	<tr>
		<td colspan="50" align="center"><b>Nome do Hospital - <%=cd_unidade%></b></td>
	</tr>
	<tr>
		<td><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
		<td colspan="45" align="center"><%=zero(dia_i)&"/"&zero(mes_i)&"/"&ano_i%></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td align="center"><b>Rack</b></td>		
		<td align="center" colspan="4">06:00</td>
		<td align="center" colspan="4">07:00</td>
		<td align="center" colspan="4">08:00</td>
		<td align="center" colspan="4">09:00</td>
		<td align="center" colspan="4">10:00</td>
		<td align="center" colspan="4">11:00</td>
		<td align="center" colspan="4">12:00</td>
		<td align="center" colspan="4">13:00</td>
		<td align="center" colspan="4">14:00</td>
		<td align="center" colspan="4">15:00</td>
		<td align="center" colspan="4">16:00</td>
		<td align="center" colspan="4">17:00</td>
		<td align="center" colspan="4">18:00</td>
		<td align="center" colspan="4">19:00</td>
		<td align="center" colspan="4">20:00</td>
		<td align="center" colspan="4">21:00</td>
		<td align="center" colspan="4">22:00</td>
		<td align="center" colspan="4">23:00</td>
		<td align="center" colspan="4">00:00</td>
		<td align="center" colspan="4">01:00</td>
		<td align="center" colspan="4">02:00</td>
		<td align="center" colspan="4">03:00</td>
		<td align="center" colspan="4">04:00</td>
		<td align="center" colspan="4">05:00</td>
	</tr>
	<%linha = 1
	strsql = "SELECT * FROM TBL_unidades_racks WHERE (cd_unidade = "&cd_unidade&") ORDER BY cd_rack"
	Set rs = dbconn.execute(strsql)
	while not rs.EOF
		cd_rack = rs("cd_rack")
		nm_rack = rs("nm_rack")%>
	<tr>
		<td align="center"><%=zero(linha)%>&nbsp;</td>
		<td><%=zero(cd_rack)&" - "&nm_rack%>&nbsp;</td>
			
			<%linha = 1
			x=0
			'strsql = "SELECT * FROM TBL_unidades_racks WHERE (cd_unidade = "&cd_unidade&") ORDER BY cd_rack"
			strsql = "SELECT COUNT (a002_numpro) AS conta , a002_datpro,a053_codung, nm_sigla,a002_horage,a002_horini,a002_horfin,tempo_cir,a056_codrac, cd_protocolo,nm_sigla, a056_desrac FROM VIEW_protocolo_lista Where A002_datpro BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND a053_codung = "&cd_unidade&" AND a002_tipage <> 'E'  AND a056_codrac = '"&cd_rack&"' Group by a002_datpro,a053_codung, nm_sigla, a002_horage, a002_horini,a002_horfin,tempo_cir, a056_codrac, cd_protocolo,nm_sigla, a056_desrac ORDER BY a056_codrac,a002_horini"
			Set rs_horarios = dbconn.execute(strsql)
			while not rs_horarios.EOF
				hora_inicio = rs_horarios("a002_horini")
					if minute(hora_inicio) >= 0 AND minute(hora_inicio) < 15 then
						hora_inicio_num = zero(hour(hora_inicio))&"00"'zero(minute(hora_inicio))
					elseif minute(hora_inicio) >= 15 AND minute(hora_inicio) < 30 then
						hora_inicio_num = zero(hour(hora_inicio))&"15"'zero(minute(hora_inicio))
					elseif minute(hora_inicio) >= 30 AND minute(hora_inicio) < 45 then
						hora_inicio_num = zero(hour(hora_inicio))&"30"'zero(minute(hora_inicio))
					elseif minute(hora_inicio) >= 30 AND minute(hora_inicio) < 59 then
						hora_inicio_num = zero(hour(hora_inicio))&"45"'zero(minute(hora_inicio))
					end if
					
					
				hora_fim = rs_horarios("a002_horfin")
					if minute(hora_fim) > 00 AND minute(hora_fim) <= 15 then
						hora_fim_num = zero(hour(hora_fim))&"30"'zero(minute(hora_inicio))
					elseif minute(hora_fim) > 15 AND minute(hora_fim) <= 30 then
						hora_fim_num = zero(hour(hora_fim))&"45"'zero(minute(hora_inicio))
					elseif minute(hora_fim) > 30 AND minute(hora_fim) <= 45 then
						hora_fim_num = zero(hour(hora_fim))&"00"'zero(minute(hora_inicio))
					elseif minute(hora_fim) > 45 AND minute(hora_fim) <= 59 then
						hora_fim_num = zero(hour(hora_fim))&"15"'zero(minute(hora_inicio))
					end if
				
					
				lista_horarios = retira_virgula(lista_horarios&","&hora_inicio_num&","&hora_fim_num&",")
				lista_horarios = REPLACE(lista_horarios,",,",",")
				'lista_horarios = REPLACE(lista_horarios,"0900,0900,","0900,")
				lista_h = retira_virgula(lista_h&","&hora_inicio_num&","&hora_fim_num&",")
				lista_h = REPLACE(lista_h,",,",",")
				'lista_h = REPLACE(lista_h,"0900,0900,","0900,")
				
				'lista_horarios = "600,630,1015,1115"
				
				if linha = 1 then
					pri_cir = hora_inicio
				end if%>
			<%rs_horarios.movenext
			wend%>
					<%img_preench = "hachura.gif"
					str_lrg = "20"
					str_alt = "12"%>
					
					<%hora_cel = "0600"%>
						<td align="center">
							<%if instr(1,lista_horarios,hora_cel,1) = 1 then%>
								<img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0">
									<%if x<>1 then%>
										<%x=1%>
									<%else%>
										<%x=0%>
									<%end if%>
							<%else%>
								<%if x=1 then%>
									<img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0">
								<%else%>
									<img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0">
								<%end if%>
							<%end if%>
						</td>
							
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					
					
					
					
					<!--<%hora_cel = "0600"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					-->
					<%hora_cel = "0615"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0630"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0645"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0700"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0715"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0730"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0745"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0800"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0815"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0830"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0845"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0900"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0915"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0930"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0945"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1000"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1015"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1030"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1045"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1100"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1115"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1130"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1145"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1200"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1215"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1230"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1245"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1300"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1315"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1330"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>		
			
					<%hora_cel = "1345"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>		
			
					<%hora_cel = "1400"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1415"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>		
			
					<%hora_cel = "1430"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1445"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1500"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1515"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1530"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1545"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1600"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1615"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1630"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1645"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1700"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1715"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1730"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1745"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1800"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1815"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1830"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1845"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1900"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1915"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1930"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "1945"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2000"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2015"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2030"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2045"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2100"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2115"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2130"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2145"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2200"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2215"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2230"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2245"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2300"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2315"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2330"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2345"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "2359"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0015"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0030"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0045"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0100"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0115"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0130"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0145"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0200"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0215"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0230"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0245"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0300"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0315"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0330"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0345"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0400"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0415"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0430"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0445"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0500"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0515"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0530"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					<%hora_cel = "0545"%><td align="center"><%if instr(1,lista_horarios,hora_cel,1) = 1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%if x<>1 then%><%x=1%><%else%><%x=0%><%end if%><%else%><%if x=1 then%><img src="../imagens/<%=img_preench%>" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%else%><img src="../imagens/px.gif" alt="" width="<%=str_lrg%>" height="<%=str_alt%>" border="0"><%end if%><%end if%></td>
					<%lista_horarios = retira_virgula(replace(lista_horarios,hora_cel,""))%>
					
					
	</tr>
	
	<%linha = linha + 1
	
	lista_horarios = ""
	rs.movenext
	wend%>
	<tr><td colspan="48"><%=lista_h%></td></tr>
</table>
<%end if%>

</body>
</html>
