<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
	<title>Untitled</title>
<%cd_unidade = request("cd_unidade")
qtd_paginas = request("qtd_paginas")
qtd_etiquetas = request("qtd_etiquetas")
mensagem = request("mensagem")
tipo_print = request("tipo_print")
espacamento = 1
seq_inicial = request("seq_inicial")
seq_final = request("seq_final")

function digv(protocolo)

	seq_ponto = instr(protocolo,".")
	seq_len = len(protocolo)
	str_unidade = mid(protocolo,1,seq_ponto-1)
	str_protocolo = proto(mid(protocolo,seq_ponto+1,seq_len))
	response.write(" "&protocolo&" - ")
	'Verificar se o protocolo é proveniente do siga
	'***********************************************
	'	if str_unidade = 2 and str_protocolo < 17222 OR str_unidade = 3 and str_protocolo < 114145 OR str_unidade = 11 and str_protocolo < 1476 OR str_unidade = 14 and str_protocolo < 2955  OR str_unidade = 15 and str_protocolo < 3102 OR str_unidade = 19 and str_protocolo < 2412 OR str_unidade = 20 and str_protocolo < 4138 OR str_unidade = 22 and str_protocolo < 65 then
	'	pula = "1"
	'	end if
	'-----------------------------------------------
	'if pula <> 1 then
	
		'	str_unid_proto = zero(str_unidade)&proto(str_protocolo)
		'		multipl = 2
		'	for i = 1 to len(str_unid_proto)
		'		caracteres = mid(str_unid_proto,i,1)
		'		soma_tudo = int(caracteres*multipl) + int(soma_tudo)
		'		multipl = multipl + 1 
		'	next		
		
		'	div_soma = soma_tudo/11
		'	resultado = round(div_soma,1)
		'		ver_resto = instr(resultado,",")
		'		resto = mid(resultado, ver_resto + 1, 1)*10
		'	resultado2 = round(div_soma,2)
		'		ver_resto2 = instr(resultado2,",")
		'		resto2 = mid(resultado2, ver_resto2 + 1,2)
			
			'Sempre Arredonda para cima
			'if int(resto2) > int(resto) then 
			'resto = int(resto/10+1)
			'Else
			'resto = resto/10
			'end if
			'*************************
			'dv =  11 - resto		
			'		if dv = 10 then
			'		dv = 0
			'		end if
		
		digv = (str_unidade&"."&str_protocolo&"-"&dv)
	'Else'if pula = 1 then
	'	digv = (str_unidade&"."&str_protocolo&"-<u style='color:gray;'>X</u>")
	'end if

end function%>
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

</head>

<body><%'Passo 1%>
<table border="1" id="no_print" align="center" style="position:relative; top:10;">
<%if cd_unidade = "" and mensagem = "" Then '***  ***%>
	<form action="../../protocolo.asp" name="pendentes" id="pendentes">
	<input type="hidden" name="tipo" value="pendentes">
	<input type="text" name="protocolo" value="<%=protocolo%>">
	<tr>
		<td colspan="2" align="center">Protocolos pendentes</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
	<tr>
		<td>Unidade</td>
		<td><select name="cd_unidade">
				<option value="0"></option>
				<%strsql ="Select * from TBL_unidades where cd_status <> 2  order by nm_unidade"
					Set rs_etq = dbconn.execute(strsql)%>
					<%Do while Not rs_etq.EOF%>
					<%cd_conv = rs_etq("cd_codigo")
					nm_unidade = rs_etq("nm_unidade")%>
					<%if cd_etiqueta=cd_conv Then%>
					<%ck_etq="selected"%><%else ck_etq=""%>
					<%end if%>
					<option value="<%=cd_conv%>" <%=ck_etq%>><%=nm_unidade%></option><%rs_etq.movenext
					Loop%>
			</select>
		</td>
	</tr>
<%elseif cd_unidade <> "" AND mensagem = "" then '%>	
	<form action="../../protocolo.asp" name="pendentes" id="pendentes">
	<input type="hidden" name="tipo" value="pendentes">
	<input type="hidden" name="protocolo" value="<%=protocolo%>">
	<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
	<tr>
		<td colspan="2" align="center">Protocolos pendentes</td>
	</tr><%if cd_unidade = 0 then%><%cd_unidade = 999%><%end if%>
	<tr><td colspan="2">&nbsp;</td></tr>
	<%strsql ="Select * from TBL_protocolo_etiquetas where cd_unidade="&cd_unidade&" order by cd_num_fim desc"
			Set rs_etq = dbconn.execute(strsql)
			if not rs_etq.eof then
			cd_num_fim = rs_etq("cd_num_fim")
			else cd_num_fim = 0
			end if%>	
	<%if cd_unidade <> 999 then%>
	<tr>
		<td>Unidade</td>
		<td><%strsql ="Select * from TBL_unidades where cd_codigo = '"&cd_unidade&"'"
					Set rs_etq = dbconn.execute(strsql)
					if not rs_etq.EOF then%>
						<%=rs_etq("nm_unidade")%>						
					<%end if%>
		</td>
	</tr>
			
	<%else%>
	<tr>
		<td>Qtd. já impressas</td>
		<td style="color:red;">&nbsp;<%'=zero(cd_unidade)&"."%><%=zero(int(cd_num_fim))%></td>
	</tr><%cd_num_inicio = cd_num_fim + 1%>
	<!--tr><td>Próxima</td>
		<td style="color:green;">&nbsp;<%=zero(cd_unidade)&"."%><%=proto(cd_num_inicio)%></td>
	</tr-->
	<%end if%>
	
	<tr><td colspan="2">&nbsp;</td></tr>
	
<%Else%>
	
	
<%end if%>
	<tr>
		<td colspan="2" align="center"><input type="submit" name="Imprimir" value="Ok" width="20">&nbsp;&nbsp;&nbsp;&nbsp;<a href="protocolo.asp?tipo=pendentes"><img src="../../imagens/ic_reprovado.gif" alt="" width="10" height="12" border="0"></a><br>
	</tr>	
	</form>
</table>
<br><br><br>
<table align="center" cellpadding="2" id="no_print">
	<%if cd_unidade <> "" Then
		strsql ="Select top 20 * from View_protocolo_etiquetas Where cd_unidade='"&cd_unidade&"' Order by cd_num_fim desc"
		Set rs_etiq = dbconn.execute(strsql)
	Else
		strsql ="Select top 40 * from View_protocolo_etiquetas Order by nm_unidade,cd_num_fim  desc"
		Set rs_etiq = dbconn.execute(strsql)
	End if
	%>
	<tr>
		<td colspan="100%" align="center" bgcolor="#c0c0c0"><b>Histórico de impressão de formulários</b></td>
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
			<td style="<%=str_stilo%>"><%=nm_unidade%></td>
			<td style="color:green; <%=str_stilo%>"><%=digitov(zero(cd_unidade)&"."&proto(int(cd_num_inicio)))%></td>
			<td style="color:red; <%=str_stilo%>"><%=digitov(zero(cd_unidade)&"."&proto(int(cd_num_fim)))%></td>
			<td style="color:gray; <%=str_stilo%>"><a href="protocolo.asp?tipo=pendentes&cd_unidade=<%=cd_unidade%>&seq_inicial=<%=int(cd_num_inicio)%>&seq_final=<%=int(cd_num_fim)%>"><%=qtd_etq%></a></td>
			<td style="<%=str_stilo%>"><%=dt_data%></td>
			<td style="<%=str_stilo%>"><%=user_inc%></td>
		</tr>
		<%cor_num = cor_num + 1
		
		if int(cor_mum) > 1 Then
			cor_num = ""
		End if
		str_stilo = ""
		
		rs_etiq.movenext
	wend%>
</table>

<br><br>
<table align="center" width="400">
	<tr><td colspan="2" align="center">Relatório de protocolos não digitados</td></tr>
	<tr>
		<td colspan="2" align="center"><%=nm_unidade%></td>
	</tr>
	<tr bgcolor="#808080" style="color:white;">
		<td>Nº</td>
		<td>Pendentes</td>				
	</tr>
	<%conta_linha = 1
	
	strsql_pendentes =   "SELECT * FROM View_protocolo_lista WHERE (A002_numpro BETWEEN "&seq_inicial&" AND "&seq_final&") and A053_codung = "&cd_unidade&" order by a002_numpro" 
	Set rs_pendentes = dbconn.execute(strsql_pendentes)%>		
	
	<%for i = seq_inicial to seq_final%>
			<%if not rs_pendentes.EOF Then
				a002_protocolo = rs_pendentes("a002_numpro")
				
				'rs_pendentes.movenext
			end if%>
				
				<%if i <> a002_protocolo then%>
					<tr>
						<td style="color:gray;"><%=conta_linha%></td>
						<td><%=i%></td>									
					</tr>
				<%conta_linha = conta_linha + 1	
				
				else
					rs_pendentes.movenext
				end if%>
			

	<%next
			
	rs_pendentes.close
	set rs_pendentes = nothing%>
</table>
<br>
</body>
</html>
