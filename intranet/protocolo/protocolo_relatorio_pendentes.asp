<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
	<title>Untitled</title>
<%strcd_unidade = request("cd_unidade")
dt_inicial = request("data_inicial")
dt_final = request("data_final")
	'ultimo_dia = UltimoDiaMes(month(dt_final),year(dt_final))
		'if ultimo_dia > day(dt_final) Then
		'	dt_final = ultimo_dia&"/"&month(dt_final)&"/"&year(dt_final)
		'end if
		if dt_inicial = "" Then
			dt_final = "1/"&month(now)&"/"&year(now)
			ultimo_dia = UltimoDiaMes(month(now),year(now))
			dt_final = ultimo_dia&"/"&month(now)&"/"&year(now)
		end if
		if dt_final = "" Then
			ultimo_dia = UltimoDiaMes(month(dt_inicial),year(dt_inicial))
			dt_final = ultimo_dia&"/"&month(dt_inicial)&"/"&year(dt_inicial)
		End if
	
qtd_paginas = request("qtd_paginas")
qtd_etiquetas = request("qtd_etiquetas")
mensagem = request("mensagem")
tipo_print = request("tipo_print")
espacamento = 1
seq_inicial = request("seq_inicial")
seq_final = request("seq_final")
%>
<style type="text/css">
<!--
.relatorio_titulo{color: #000000;font-family: verdana;font-size:12px;font-weight:bold}
.relatorio_campos{color: #000000;font-family: verdana;font-size:7px;}
.relatorio_unidade{color: #000000;font-family: verdana;font-size:9px;font-weight:bold}
.relatorio_protocolo{color: #000000;font-family: verdana;font-size:13px;font-weight:bold}
-->
</style>
<!--onClick="window.print()-->
<SCRIPT LANGUAGE="javascript">
function Jsconfirm_print()
{
  if (confirm ("Tem certeza que deseja imprimir?"))
	  {
		//document.location.href('acoes/patrimonio_acao.asp?cd_apaga='+cod1+'&cd_tipo=patrimonio');
		
	  }
}
</SCRIPT>
<%
function dverif(protocolo)
	tam_sequencia = len(protocolo)
	ponto = instr(protocolo,".")
		cod_unidade = mid(protocolo,1,ponto-1)
			if unidade < 10 Then
				cod_unidade = "0"&cod_unidade
			Else
				cod_unidade = cod_unidade
			End if
			
		cod_protocolo = mid(protocolo,ponto+1,(tam_sequencia-ponto))
			len_proto = len(cod_protocolo)
				'for i = len_proto to 6
				'	cod_protocolo = "0"&cod_protocolo
				'next
				if 	len_proto = 5 then
					cod_protocolo = "0"&cod_protocolo
				elseif 	len_proto = 4 then
					cod_protocolo = "00"&cod_protocolo
				elseif 	len_proto = 3 then
					cod_protocolo = "000"&cod_protocolo
				elseif 	len_proto = 2 then
					cod_protocolo = "0000"&cod_protocolo
				elseif 	len_proto = 1 then
					cod_protocolo = "00000"&cod_protocolo
				end if
			
		dverif = cod_unidade&"."&cod_protocolo
end function
%>
</head>

<body><%'Passo 1%>
<table border="1" class="no_print" align="center" style="position:relative; top:10;">
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">protocolo_relatorio_pendentes.asp</span></td></tr>
	<%end if%>
	<form action="../../protocolo.asp" name="pendentes" id="pendentes">
	<input type="hidden" name="tipo" value="pendentes">	
	<tr>
		<td colspan="2" align="center">Protocolos pendentes</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<td>Unidade</td>
		<td><select name="cd_unidade">
				<option value="0"></option>
				<%strsql ="Select * from TBL_unidades where cd_status = 1   order by nm_unidade "
					Set rs = dbconn.execute(strsql)%>
					<%Do while Not rs.EOF%>
					<%cd_unidade = rs("cd_codigo")
					nm_unidade = rs("nm_unidade")%>
						<%if int(strcd_unidade) = int(cd_unidade) Then%><%ck_und="selected"%><%else%><%ck_und=""%><%end if%>
					<option value="<%=cd_unidade%>" <%=ck_und%>><%=nm_unidade%></option><%rs.movenext
					Loop%>
			</select>
		</td>
	</tr>
	<tr>
		<td>Período</td>
		<td><input type="text" name="data_inicial" size="11" maxlength="10" value="<%=dt_inicial%>"> &nbsp; até: &nbsp; <input type="text" name="data_final" size="11" maxlength="10" value="<%=dt_final%>"></td>
	</tr>
	
	<tr>
		<td colspan="2" align="center"><input type="submit" name="Imprimir" value="Ok" width="20">&nbsp;&nbsp;&nbsp;&nbsp;<a href="protocolo.asp?tipo=pendentes"><img src="../../imagens/ic_reprovado.gif" alt="" width="10" height="12" border="0"></a><br>
	</tr>	
	</form>
</table>
<br class="no_print"><br class="no_print"><br class="no_print">
<table align="center" cellpadding="2">
	<%strsql ="Select * from View_protocolo_etiquetas Where cd_unidade='"&strcd_unidade&"' and dt_data between '"&month(dt_inicial)&"/"&day(dt_inicial)&"/"&year(dt_inicial)&"' and '"&month(dt_final)&"/"&day(dt_final)&"/"&year(dt_final)&"' Order by cd_num_fim "
	set rs_etiq = dbconn.execute(strsql)%>
	<tr>
		<td colspan="100%" align="center" bgcolor="#c0c0c0"><b>Relação de impressão de formulários</b></td>
	</tr>
	<tr>
		<td align="center"><b>Unidade</b></td>
		<td align="center"><b>Ínicio</b></td>
		<td align="center"><b>Fim</b></td>
		<td align="center"><b>Qtd.</b></td>
		<td align="center"><b>Data</b></td>
		<td><b>Usuário</b></td>		
	</tr>	
	<%While not rs_etiq.EOF
	
		cd_codigo = rs_etiq("cd_codigo")
		nm_unidade = rs_etiq("nm_unidade")
		cd_num_inicio = proto(rs_etiq("cd_num_inicio"))
		cd_num_fim = rs_etiq("cd_num_fim")
		qtd_etq = (cd_num_fim+1) - cd_num_inicio
		dt_data = rs_etiq("dt_data")
		user_inc = rs_etiq("nm_nome")
		
		
			if cor_num = 1 Then
				cor_linha = "#FFFFFF"
			else 
				cor_linha = "#dfdfdf"
			end if
			
			if int(cd_num_inicio) = int(seq_inicial) and int(cd_num_fim) = int(seq_final)then
				str_stilo = " font-size:12 px;" 
			end if%>
		
		<tr bgcolor="<%=cor_linha%>">
			<td align="center" style="<%=str_stilo%>"><%=nm_unidade%></td>
			<td align="center" style="color:green; <%=str_stilo%>"><%=digitov(zero(strcd_unidade)&"."&proto(int(cd_num_inicio)))%></td>
			<td align="center" style="color:red; <%=str_stilo%>"><%=digitov(zero(strcd_unidade)&"."&proto(int(cd_num_fim)))%></td>
			<td align="center" style="color:gray; <%=str_stilo%>"><%=qtd_etq%></td>
			<td align="center" style="<%=str_stilo%>"><%=dt_data%></td>
			<td align="center" style="<%=str_stilo%>"><%=user_inc%></td>
		</tr>
		<%cor_num = cor_num + 1
		
		if int(cor_mum) > 1 Then
			cor_num = ""
		End if
		str_stilo = ""
		
		qtd_total = qtd_total + qtd_etq
		
		rs_etiq.movenext
	wend	%>
		<tr>
			<td colspan="3" align="right"><b>Total: &nbsp;</b></td>
			<td><b><%=qtd_total%></b></td>
		</tr>
		<tr>
			<td><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="120" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
		</tr>
		<tr><td colspan="6" style="background-color:silver;"><img src="../imagens/px.gif" alt="" width="1" height="1" border="0"></td></tr>
</table>
<%'qtd_total = ""%>
<br><br class="no_print">
<br class="no_print">
<table border="1" align="center" cellpadding="2" style="position:relative; top:10;" width="300">
	<%
	'strsql ="Select * from View_protocolo_etiquetas Where cd_unidade='"&strcd_unidade&"' and dt_data between '"&month(dt_inicial)&"/"&day(dt_inicial)&"/"&year(dt_inicial)&"' and '"&month(dt_final)&"/"&day(dt_final)&"/"&year(dt_final)&"' Order by cd_num_fim "
	strsql = "SELECT ALL MAX(A002_numpro) AS MAX FROM TBL_Protocolo WHERE (A053_codung = '"&strcd_unidade&"')"
	set rs_max = dbconn.execute(strsql)
	if not rs_max.EOF then
		ultimo_digitado = rs_max("max")
	end if
	
	
	strsql ="Select * from View_protocolo_etiquetas Where cd_unidade='"&strcd_unidade&"' and dt_data between '"&month(dt_inicial)&"/"&day(dt_inicial)&"/"&year(dt_inicial)&"' and '"&month(dt_final)&"/"&day(dt_final)&"/"&year(dt_final)&"' Order by cd_num_fim "
	set rs_etiq = dbconn.execute(strsql)%>
	<tr>
		<td colspan="100%" align="center" bgcolor="#c0c0c0"><b>Relação de números de formulários não digitados</b></td>
	</tr>
	<tr style="color:white;" bgcolor="#808080">
		<td align="center"><b>&nbsp;</b></td>
		<td align="center"><b>Protocolo</b></td>
		<td align="center"><b>Data de impressão</b></td>
		<td align="center"><b>Total<br>Impressos</b></td>
	</tr>
	
	<%conta_linha = 1
	rodape = ""
	While not rs_etiq.EOF
		
		
		cd_codigo = rs_etiq("cd_codigo")
		nm_unidade = rs_etiq("nm_unidade")
		cd_num_inicio = proto(rs_etiq("cd_num_inicio"))
		cd_num_fim = rs_etiq("cd_num_fim")
		qtd_etq = (cd_num_fim+1) - cd_num_inicio
		dt_data = rs_etiq("dt_data")
		user_inc = rs_etiq("nm_nome")
			if int(cd_num_fim) >= ultimo_digitado then
				cd_num_fim = ultimo_digitado
				asd = "maior"
			else
			'	cd_num_fim = cd_num_fim
			end if
		
		
			for i = cd_num_inicio to cd_num_fim
					'strsql_pendentes = "SELECT * FROM View_protocolo_lista WHERE (A002_numpro = "&i&") and A053_codung = "&strcd_unidade&" order by a002_numpro" 
					strsql_pendentes = "SELECT * FROM TBL_protocolo WHERE (A002_numpro = "&i&") and A053_codung = "&strcd_unidade&" order by a002_numpro" 
					Set rs_pendentes = dbconn.execute(strsql_pendentes)%>
						<%if not rs_pendentes.EOF Then
							a002_protocolo = rs_pendentes("a002_numpro")
						end if%>
				
				<%if i <> a002_protocolo then%>	
					<tr>
						<td align="center" style="color:gray;"><%=zero(conta_linha)%></td>
						<td align="center"><a href="protocolo.asp?tipo=digitacao&modalidade=cancelar&cd_unidade=<%=strcd_unidade%>&no_protocolo=<%=i%>&no_digito=o"><%=dverif(strcd_unidade&"."&i)%></a></td>
						<td align="center"><%=zero(day(dt_data))&"/"&zero(month(dt_data))&"/"&year(dt_data)%></td>
						<td align="center"><%=qtd_etq%></td>
					</tr>
					<%conta_linha = conta_linha + 1
					mostra_sub = 1
					subtotal = subtotal + 1
				else
					rs_pendentes.movenext
				end if%>				
				
			<%next%>		
		<%if mostra_sub = 1 then%>
		<tr><td colspan="4" bgcolor="silver"> &nbsp; <b>Sub-total:</b> <%=subtotal%> (<%=round(((subtotal)/qtd_etq)*100,2)%>%)</td></tr>
		<%mostra_sub = "0"
		end if%>
		
	<%subtotal = 0
	rs_etiq.movenext
	wend
		
		if (conta_linha-1) > 0 Then%>
		
		<tr><td colspan="100%">&nbsp;</td></tr>
		<tr>
			<td align="right"><b>Total: &nbsp;</b></td>
			<td colspan="2"><b><%=conta_linha-1%> formulários não digitados do total de <%=qtd_total%></b></td>
			<td>(<%=round(((conta_linha-1)/qtd_total)*100,2)%>%)</td>
		</tr>
		<%end if%>
		<tr>
			<td><img src="imagens/blackdot.gif" width="50" height="1"></td>
			<td><img src="imagens/blackdot.gif" width="115" height="1"></td>
			<td><img src="imagens/blackdot.gif" width="115" height="1"></td>
			<td><img src="imagens/blackdot.gif" width="50" height="1"></td>
		</tr>
</table><br>

</body>
</html>
