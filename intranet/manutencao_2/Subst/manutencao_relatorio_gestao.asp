<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%
dia_i = request("dia_i")
mes_i = request("mes_i")
ano_i = request("ano_i")
	if dia_i = "" or mes_i = "" or ano_i = "" Then
		dia_i = day(now)
		mes_i = month(now)
		ano_i = year(now)
	end if

dia_f = request("dia_f")
mes_f = request("mes_f")
ano_f = request("ano_f")
	if dia_f = "" or mes_f = "" or ano_f = "" Then
		dia_f = ultimodiames(month(now),year(now))
		mes_f = month(now)
		ano_f = year(now)
	end if

%>

<table align="center" style="" border="1">
	<form action="../manutencao.asp" name="form" id="form">
	<input type="hidden" name="tipo" value="gestao">
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
<br id="no_print">
<table border="1" align="center">
	<tr><td colspan="100" align="center">RELATÓRIO DE GESTÃO</td></tr>
	<tr><td colspan="100" align="center">&nbsp;</td></tr>
	
	<tr><td>&nbsp;</td>
		<%xsql = "SELECT DISTINCT cd_unidade, nm_sigla FROM VIEW_os_lista WHERE (sequencia = '1') AND (dt_entrada BETWEEN '"&mes_i&"/"&dia_i&"/"&ano_i&"' AND '"&mes_f&"/"&dia_f&"/"&ano_f&"')" 
		Set rs = dbconn.execute(xsql)
		
		while not rs.EOF
			cd_unidade = rs("cd_unidade")
			nm_sigla = rs("nm_sigla")
			lista_und = lista_und &","& cd_unidade%>
		<td><%=nm_sigla%></td>
		<%rs.movenext
		wend%>
	</tr>
	
	<tr>
		<td></td>
		<td><%=lista_und%></td>
	</tr>


	
	


</table>

<!--
'SELECT DISTINCT cd_unidade, nm_unidade
'FROM         VIEW_os_lista
'WHERE     (sequencia = '1') AND (dt_entrada BETWEEN '6/1/2011' AND '6/30/2011')
'GROUP BY cd_unidade, nm_unidade
'ORDER BY nm_unidade
-->