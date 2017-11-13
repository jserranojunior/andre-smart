<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../css/geral.htm"-->
<%
dia_i = request("dia_i")
mes_i = request("mes_i")
ano_i = request("ano_i")

	if dia_i = "" or mes_i = "" or ano_i = "" Then
		dia_i = day(now)
		mes_i = month(now)
		ano_i = year(now)-1
	end if
data_i = mes_i&"/"&dia_i&"/"&ano_i
	

dia_f = request("dia_f")
mes_f = request("mes_f")
ano_f = request("ano_f")

	if dia_f = "" or mes_f = "" or ano_f = "" Then
		dia_f = ultimodiames(month(now),year(now))
		mes_f = month(now)
		ano_f = year(now)
	end if
data_f = mes_f&"/"&dia_f&"/"&ano_f
	
jan = request("jan")
	if jan = 1 then
		url = "manutencao_relatorio_gestao.asp"
	else
		url = "../manutencao2.asp"
		jan = 0
	end if
%>

<table align="center" style="" border="1" class="no_print">
	<form action="<%=url%>" name="form" id="form">
	<input type="hidden" name="tipo" value="gestao">
	<input type="hidden" name="jan" value="<%=jan%>">
	<tr>
		<td colspan="3" align="center"><b>Relatório da Gestão de Ordens de Serviço</b></td>
	</tr>
	
	<%bgc_linha = "FFFFFF"%>
	<tr class="no_print" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#<%=bgc_linha%>');" bgcolor="#<%=bgc_linha%>">
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
			até
			<select name="dia_f">
				<%numero = 1
				do while NOT numero = 32
					'if dia_f = "" then
					'teste = "nada"
					'Else
					'teste = "algo"
					'end if
					
					if int(dia_f) = numero Then
					check = "selected"
					'elseif
					end if%>					
				<option value="<%=numero%>" <%=check%>><%=numero%></option>
			<%numero = numero+1
			check = ""
			loop%>
			</select>/
			<select name="mes_f">
				<%numero = 1
				do while NOT numero = 13
					if mes_f = "" Then
						if mes_hoje = numero Then
						'numero = 12
						check = "selected"
						end if
					Elseif int(mes_f) = numero Then
						check = "selected"
					end if%>		
				<option value="<%=numero%>" <%=check%>><%=left(mes_selecionado(numero),3)%></option>
				<%numero = numero+1
			check = ""
			loop%>							
			</select> /
			<%if ano_f = "" then%><%ano_f=ano_atual%><%end if%> 
			<input type="text" name="ano_f" size="4" maxlength="4" value="<%=ano_f%>"> 
		</td>
		<!--td>
		<%ordem_var = ordem_17
		ordem_num="ordem_17"
		ordem_campo="dt_os"%>
		<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none; background-color:#<%=bgc_linha%>;" value="<%=ordem_var%>"></td-->
	</tr>
	<tr><td><input type="submit" name="OK" value="OK"></td></tr>
	</form>
</table>
<br class="no_print">
<table border="1" align="center" style="border-collapse:collapse; border:1px black solid;">
	<tr><td colspan="11" align="center" style="background-color:#808080; color:#ffffff;font-size:15px;">RELATÓRIO DA GESTÃO DE ORDENS DE SERVIÇO</td></tr>
	<tr><td colspan="11" align="center" style="background-color:#ffffff; color:#000000;font-size:12px;">&nbsp;Período: <%=zero(day(data_i))&"/"&zero(month(data_i))&"/"&year(data_i)%> até <%=zero(day(data_f))&"/"&zero(month(data_f))&"/"&year(data_f)%></td></tr>
		<%'strsql2 = "SELECT COUNT(num_qtd) as conta,cd_unidade_os FROM VIEW_os_lista2 WHERE (dt_entrada BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"') AND (fecha BETWEEN 1 AND 1) GROUP BY fecha, cd_unidade_os ORDER BY cd_unidade_os"
		strsql2 = "SELECT COUNT(num_qtd) as conta,cd_unidade FROM VIEW_os_lista2 WHERE (dt_entrada BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"') AND (fecha BETWEEN 1 AND 1) GROUP BY fecha, cd_unidade ORDER BY cd_unidade"
							set rs_fechadas = dbconn.execute(strsql2)
								while not rs_fechadas.EOF 
									conta_fecha = rs_fechadas("conta")
									'cod_unidade = rs_fechadas("cd_unidade_os")
									cod_unidade = rs_fechadas("cd_unidade")
									Lista_fechadas = retira_virgula(lista_fechadas&","&conta_fecha&":"&cod_unidade)
									'response.write(Lista_fechadas&"<br>")
								rs_fechadas.movenext
								wend%>
	
		<%'xsql = "SELECT DISTINCT cd_unidade_os,nm_sigla FROM VIEW_os_lista2 WHERE  (dt_entrada BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"') and fecha between 0 and 1"
		xsql = "SELECT DISTINCT cd_unidade,nm_sigla FROM VIEW_os_lista2 WHERE  (dt_entrada BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"') and fecha between 0 and 1"
		'xsql = "SELECT DISTINCT cd_unidade,nm_sigla FROM VIEW_os_lista2 WHERE  (dt_entrada BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"') and fecha between 0 and 1" 
		Set rs = dbconn.execute(xsql)
			while not rs.EOF
				'cd_unidade = rs("cd_unidade_os")
				cd_unidade = rs("cd_unidade")
				nm_sigla_os = rs("nm_sigla")
				lista_und = lista_und &","&cd_unidade&":"&nm_sigla_os
				'lista_und = lista_und &","&cd_unidade%>		
			<%rs.movenext
			wend%>
	<tr style="border:5px black solid;background-color:#b2b2b2; color:#000000;">
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<%xsql = "SELECT * FROM TBL_Gestao ORDER BY cd_ordem"
		Set rs_gest = dbconn.execute(xsql)
			while not rs_gest.EOF
				cd_gestao = rs_gest("cd_codigo")
				nm_gestao = rs_gest("nm_gestao")
					lista_gest = retira_virgula(lista_gest &","& cd_gestao)
					lista_gest_parc = lista_gest
					qtd_gestao = qtd_gestao + 1
					%>
					<td align="center"><b><%=nm_gestao%> <%=" - "&cd_gestao%>
			<%rs_gest.movenext
			wend%>
		<td align="center"><b>Pendentes</b></td>
		<td align="center"><b>Fechadas</b></td>
		<td align="center"><b>Total</b></td>	
	</tr>
			<%num_linha = 1
			unidades_array = split(retira_virgula(lista_und),",")
				For Each unidade_item In unidades_array
					cd_unidade_item = mid(unidade_item,1,instr(1,unidade_item,":",1)-1)
					nm_sigla_item = mid(unidade_item,instr(1,unidade_item,":",1)+1,len(unidade_item))
					'cd_unidade_item = unidade_item
					if IsNumeric(cd_unidade_item) then%>
						
						<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
							<td style="background-color:#e2e2e2;">&nbsp;<%=zero(num_linha)%></td>
							<td><b><%=nm_sigla_item%></b></td>
							<%'strsql = "SELECT COUNT(num_qtd) as conta,fecha,cd_unidade_os,nm_sigla,cd_gestao,nm_gestao,cd_gestao_ordem FROM VIEW_os_lista2 WHERE (dt_entrada BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"') AND (cd_unidade_os = '"&cd_unidade_item&"') AND (fecha BETWEEN '0' AND '0') GROUP BY sequencia, fecha, cd_unidade_os,nm_sigla, cd_gestao,nm_gestao,cd_gestao_ordem, sequencia ORDER BY cd_gestao_ordem"
							strsql = "SELECT COUNT(num_qtd) as conta,fecha,cd_unidade,nm_sigla,cd_gestao,nm_gestao,cd_gestao_ordem FROM VIEW_os_lista2 WHERE (dt_entrada BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"') AND (cd_unidade = '"&cd_unidade_item&"') AND (fecha BETWEEN '0' AND '0') GROUP BY sequencia, fecha, cd_unidade,nm_sigla, cd_gestao,nm_gestao,cd_gestao_ordem, sequencia ORDER BY cd_gestao_ordem"
							set rs = dbconn.execute(strsql)
								
								while not rs.EOF
									conta = rs("conta")
									nm_sigla = rs("nm_sigla")
									'cd_unidade_os = rs("cd_unidade_os")
									cd_unidade_os = rs("cd_unidade")
									cd_gestao = rs("cd_gestao")
									nm_gestao = rs("nm_gestao")
									conta_pos = conta_pos + 1%>
											<%if instr(1,lista_gest_parc,cd_gestao,1) > 1 then
												do while not instr(1,lista_gest_parc,cd_gestao,1) = 1
													lista_gest_parc = retira_virgula(mid(lista_gest_parc,2,len(lista_gest_parc)))%>
													<td>&nbsp;</td>
												<%conta_pos = conta_pos + 1
												loop
											end if%>
												<%titulo = "Gestão"
												'condicao = "WHERE (dt_os BETWEEN |"&data_i&"| AND |"&data_f&"|) AND (cd_unidade_os = "&cd_unidade_item&") AND (cd_gestao = "&cd_gestao&") AND fecha = 0"
												condicao = "WHERE (dt_os BETWEEN |"&data_i&"| AND |"&data_f&"|) AND (cd_unidade = "&cd_unidade_item&") AND (cd_gestao = "&cd_gestao&") AND fecha = 0"%>
												<td align="center" style="cursor:pointer; cursor:hand;" onclick="window.open('<%if jan <> 1 then%>manutencao_2/<%end if%>manutencao_lista_janela2.asp?condicao=<%=condicao%>&titulo=<%=titulo%>','visualizar','width=520,height=350,scrollbars=yes'); return false;">&nbsp;<span style="color:red;"><%'=cd_gestao&" - "%></span><%'=instr(1,lista_gest_parc,cd_gestao,1)&" - "%><%'=lista_gest_parc&" - "%><%=conta%></td>
											
										<%lista_gest_parc = replace(lista_gest_parc,cd_gestao&",","")
										sub_gestao = retira_virgula(sub_gestao&","&conta&":"&cd_gestao)
										total_und = total_und + conta
										
										'identificador_und = cd_unidade_os
										identificador_und = cd_unidade
								rs.movenext
								wend
									
									do while not lista_gest_parc = ""
										lista_gest_parc = retira_virgula(mid(lista_gest_parc,2,len(lista_gest_parc)))
									loop%>					
								
								<%if total_und = 0 then
									for i_nada = 1 to qtd_gestao
										conta_pos = conta_pos + 1%>
										<td>&nbsp;</td>
									<%next
								end if%>
									<%if conta_pos < qtd_gestao then
										for i = conta_pos to qtd_gestao-1
											conta_pos = conta_pos + 1%>
											<td>&nbsp;</td>
										<%next%>
									<%end if%>
								
								<%titulo = "Total Pendencias"
								'condicao = "WHERE (dt_os BETWEEN |"&data_i&"| AND |"&data_f&"|) AND (cd_unidade_os = "&cd_unidade_item&") AND fecha = 0"%>
									<td align="center" style="color:green;cursor:pointer; cursor:hand;" onclick="window.open('manutencao_2/manutencao_lista_janela2.asp?condicao=<%=condicao%>&titulo=<%=titulo%>','visualizar','width=520,height=350,scrollbars=yes'); return false;"><b><%=total_und%></b></td>
								<td align="center" style="color:red;background-color:e2e2e2;"><%und_fechadas = retira_virgula(instr(1,lista_fechadas,":"&cd_unidade_item,1))
					'*** Fechadas ********************************************************
								'response.write("("&und_fechadas&") - {"&lista_fechadas&"}")
								if lista_fechadas <> "" Then
									pos_und = instrRev(lista_fechadas,":"&cd_unidade_item,-1,1)
									if pos_und > 0 then
										pos_nfechadas = instrRev(lista_fechadas,",",pos_und,1)
									else
										pos_nfechadas = 0
									end if
									len_lista_fechadas = len(lista_fechadas)
										if pos_nfechadas > 0 then
											qtd_und = mid(lista_fechadas,pos_nfechadas+1,(pos_und-pos_nfechadas)-1)
										else
											qtd_und = 0
										end if
									
								end if%>
								<b><%=qtd_und%></b>
								<%if qtd_und = "" Then
									qtd_und = 0
									total_fechadas = 0
								end if
								total_fechadas = int(total_fechadas)+qtd_und
								'lista_fechadas = retira_virgula(replace(lista_fechadas,qtd_und&":"&cd_unidade_item,""))%></td>
								<td align="center" style="background-color:e2e2e2;"><b><%=qtd_und+total_und%></b></td>
						</tr>
					<%total_geral = total_geral + qtd_und + total_und
					conta_pos = 0
					conta_fechadas = 0
					total_und = 0
					lista_gest_parc = lista_gest
					num_linha = num_linha + 1
					end if
				Next%>
	<tr style="background-color:#b2b2b2;">
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
		<%gestao_array = split(lista_gest,",")
		For Each gestao_item In gestao_array%>
			<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<%next%>
		<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
	</tr>
	<tr style="background-color:#e2e2e2;">
		<td>&nbsp;</td>
		<td><b>TOTAL</b></td>
		<%gestao_array = split(lista_gest,",")
		For Each gestao_item In gestao_array
			
			subgestao_array = split(sub_gestao,",")'sub_gestao
			conta_parc_gestao = 0
			For Each subgestao_item In subgestao_array
				conta_sub_item = mid(subgestao_item,1,instr(1,subgestao_item,":",1)-1)
				conta_sub_gest = mid(subgestao_item,instr(1,subgestao_item,":",1)+1,len(subgestao_item))
					if conta_sub_gest = gestao_item then
						conta_parc_gestao = int(conta_parc_gestao) + conta_sub_item
					end if
			next%>
		<td align="center"><b><%=conta_parc_gestao%></b></td>
		<%total_conta_parc_gestao = total_conta_parc_gestao+conta_parc_gestao
		conta_parc_gestao = 0
		conta_sub_item = 0
		next%>
		<td align="center" ><span style="font-weight:bold; color:green;"><%=total_conta_parc_gestao%></span></td>
		<td align="center"><span style="font-weight:bold; color:red;"><%=total_fechadas%></span></td>
		<td align="center"><span style="font-weight:bold; color:black; font-weight:bold;"><%=total_geral%></span></td>		
	</tr>
	<tr id="ok_print"><td align="right" colspan="<%=qtd_gestao+5%>" style="font-size:9px; font-family:verdana; color:gray;">impresso em: <%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)&" "&zero(hour(now))&":"&zero(minute(now))%> &nbsp; </td></tr>
</table>

<!--
'SELECT DISTINCT cd_unidade, nm_unidade
'FROM         VIEW_os_lista
'WHERE     (sequencia = '1') AND (dt_entrada BETWEEN '6/1/2011' AND '6/30/2011')
'GROUP BY cd_unidade, nm_unidade
'ORDER BY nm_unidade
-->