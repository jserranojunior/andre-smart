<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->

<script language="JavaScript">
	function periodo_inicial(){
		document.form.periodo_in.value="";
	}
	function periodo_final(){
		document.form.periodo_out.value="";
	}
</script>

<%
ok = request("ok")

origem = zero(1)
dt_gravacao = right(year(now),2)&zero(month(now))&zero(day(now))&zero(hour(now))&zero(minute(now))
usuario =  session("cd_codigo")
nome_arquivo = request("nome_arquivo")
nome_gravacao = origem&"_"&dt_gravacao&"_"&zero(usuario)&"-"&nome_arquivo
grava_txt = request("conteudo")

gravados_bkp = request("gravados_bkp")

strcd_unidade = request("cd_unidade")
periodo_in = request("periodo_in")
periodo_out = request("periodo_out")
	if periodo_in = "" or periodo_out = "" Then
		periodo_in = "01/"&zero(month(now))&"/"&year(now)
		periodo_out = ultimodiames(month(now),year(now))&"/"&zero(month(now))&"/"&year(now)
	end if
%>

	<body>
	<table border="1">
		<form action="patrimonio_plan_serv_abre.asp" name="form">
		<input type="hidden" name="conteudo" value="<%=grava_txt%>">
		<tr><td colspan="2" align="center">ABRIR PLANEJAMENTO DE SERVIÇO</td></tr>
		<tr>
			<td>Unidade</td>
			<td>
			<%strsql = "SELECT * FROM TBL_unidades where cd_status = 1 "'AND dt_gravacao Between '"&&"' and '"&&"'"
			Set rs = dbconn.execute(strsql)%>
				<select name="cd_unidade">
					<option value="">***</option>
					<%while not rs.EOF
						cd_unidade = rs("cd_codigo")
						nm_unidade = rs("nm_unidade")%>
						<option value="<%=cd_unidade%>" <%if int(strcd_unidade)=cd_unidade then%>Selected<%end if%>><%=nm_unidade%></option>
					<%rs.movenext
					wend%>
				</select>
		</td>
		<tr>
			<td>Períoido</td>
			<td><input type="text" name="periodo_in" size="10" maxlength="10" value="<%=periodo_in%>" onClick="periodo_inicial();"> até <input type="text" name="periodo_out" size="10" maxlength="10"  value="<%=periodo_out%>" onClick="periodo_final();"></td>
		</tr>
		</tr>
		<tr><td colspan="2"><input type="submit" name="ok" value="Abrir"></td></tr>
		</form>	
	</table>
	
<%if ok <> "" Then%>
	<br>
	<table border="1">
		<tr><td colspan="4" align="center"><b>LISTA DE PLANOS SALVOS</b></td></tr>
			<%num = 1
			strsql = "SELECT * FROM TBL_grava_bkp where dt_gravacao between '"&month(periodo_in)&"/"&day(periodo_in)&"/"&year(periodo_in)&" 00:00' AND '"&month(periodo_out)&"/"&day(periodo_out)&"/"&year(periodo_out)&" 23:59'"
			Set rs = dbconn.execute(strsql)
			
				if rs.EOF then%>
				<tr><td>Nenhum registro encontrado</td></tr>
				<%else%>
				<tr>
					<td>&nbsp;</td>
					<td>Nome</td>
					<td>Data</td>
					<td>Usuario</td>
				</tr>
					<%while not rs.EOF
						nm_arquivo = rs("nm_arquivo")
						dt_gravacao = rs("dt_gravacao")
						cd_usuario = rs("cd_usuario")%>
						<tr>
							<td><%=num%></td>
							<td><a href="javascript:void(0);" onclick="window.open('patrimonio_plan_serv_bkp.asp?nm_arquivo=<%=nm_arquivo%>','nome','scrollbars = yes','width=20,height=10'); return false;"><%=mid(nm_arquivo,"21",len(nm_arquivo))%></a></td>
							<td><%=zero(day(dt_gravacao))&"/"&zero(month(dt_gravacao))&"/"&year(dt_gravacao)&""%></td>
							<td><%=cd_usuario%></td>
						</tr>				
					<%num = num + 1
					rs.movenext
					wend
				End if%>
	</table>
<%end if%>
	
</body>
