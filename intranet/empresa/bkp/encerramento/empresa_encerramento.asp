<!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--include file="includes/inc_area_restrita.asp"-->

<HTML>
<HEAD>
<TITLE> Encerramento </TITLE>

</HEAD>
<BODY><br><br>
<table align="center" border="1" cellspacing="0" cellpadding="0">
<%'Verifica o ultimo mês encerrado
	xsql = "SELECT * FROM TBL_fechamento_mes WHERE (cd_origem >= 1) ORDER BY cd_codigo DESC"
	'dbconn.execute xsql 
	Set rs = dbconn.execute(xsql)
	if not rs.eof then
			nm_fechado = rs("nm_fechado")
			cd_mes = rs("cd_mes")
			cd_ano = rs("cd_ano")
			
		'response.write(nm_fechado)
	end if
	
'********* Data Limite inicial **********
	if cd_mes < 3 AND cd_ano <= 2008 Then
		cd_mes = 2
		cd_ano = 2008
		start_empresa = 0
	end if
'****************************************
	%>
	<tr>
		<td>
			<table width=350 border=0 cellspacing="0" cellpadding="0">
				<tr>
					<td class="txt_cinza"><br>&nbsp;<b>Encerramento &raquo; <font color="red">Empresa</font></b><br><br></td>
				</tr>	
				<tr>
					<td align=center bgcolor=black colspan=3></td>
				</tr>
				<%if month(now) = cd_mes+1 then%>
				<tr>
				<td><%response.write("&nbsp;O mes atual está em andamento<br>&nbsp;não pode ser encerrado")%></td>
				</tr>
				<%Else
				
				'*** String do mês subsequênte ***
					str_mes = cd_mes + 1
					str_ano = cd_ano
						if str_mes > 12 Then
						str_mes = 1
						str_ano = cd_ano + 1
						end if
				'*********************************
				%>
				
				<tr><form action="acoes/empresa_encerramento_acao.asp" name="fim_mes_emp" id="fim_mes_emp" method=".">
					<input type="hidden" name="mes_sel" value="<%=str_mes%>">
					<input type="hidden" name="ano_sel" value="<%=str_ano%>">
					<input type="hidden" name="cd_origem" value="2">
					<td><br>&nbsp;Mês atual: <%=mes_selecionado(month(now))%>/<%=year(now)%>.<br><br>
					<%if start_empresa <> "0" Then%>
					&nbsp;Último mês encerrado: <%=mes_selecionado(cd_mes)%>/<%=cd_ano%>.<br><br>
					<%end if%>
					&nbsp;<font color=red>Deseja encerrar: <%=mes_selecionado(str_mes)%>/<%=str_ano%>?</font><br><br>
					&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="sim" Value="Sim" onClick="document.write(mensagem)">&nbsp;<br><br>
					</form></td>
				</tr>
				<%end if%>
				<tr>
					<td></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</BODY>
</HTML>