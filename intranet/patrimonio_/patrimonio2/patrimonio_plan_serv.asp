<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_pat_codigo = request("cd_pat_codigo")
'str_patrimonio = request("cd_patrimonio")
'nm_tipo = request("nm_tipo")
'cd_equipamento =  request("cd_equipamento")
'nm_equipamento = request("nm_equipamento")
'cd_especialidade = request("cd_especialidade")
'	if cd_especialidade = "" Then 
'		cd_especialidade = "0" 
'	end if
'cd_marca =  request("cd_marca")
'nm_marca =  request("nm_marca")
'cd_ns =  request("cd_ns")
'cd_pn =  request("cd_pn")
'cd_unidade =  request("cd_unidade")
'nm_sigla =  request("nm_sigla")'

'dt_data =  request("dt_data")

ano_sel = request("ano_sel")
	if ano_sel = "" Then
		ano_sel = year(now)
	end if

strcd_unidade  =  request("cd_unidade")
ordem = request("ordem")
	if ordem = "" Then
	ordem = "nm_tipo, cd_patrimonio"
	end if

'acao = request("acao")
'	if acao = "" Then
'	acao = "inserir"
'	Else
'	acao = "editar"
'	end if

mensagem = request("mensagem")
%>
<head>
	<title>Untitled</title>
</head>

<body>
<br id="no_print">
<table align="center" id="no_print">
	<form action="patrimonio.asp" method="post">
	<input type="hidden" name="tipo" value="plan_serv">
	<tr>
		<td>&nbsp; Ano</td>
		<td><input type="text" name="ano_sel" value="<%=ano_sel%>"></td>
	</tr>
	<tr>
	    <td>&nbsp; Unidade:</td>
			<%strsql = "SELECT * FROM TBL_unidades where cd_status <> 0"
			Set rs_uni = dbconn.execute(strsql)%>
	    <td><select name="cd_unidade" class="inputs" onFocus="nextfield ='num_hospital';">
						<option value=""></option>
						<%Do While Not rs_uni.eof
						if int(cd_unidade) = rs_uni("cd_codigo") then%><%unidade_check = "selected"%><%end if%>
						<option value="<%=rs_uni("cd_codigo")%>" <%=unidade_check%>><%=rs_uni("nm_unidade")%></option>
						<%rs_uni.movenext
						unidade_check = ""
						loop%>
						</select>
						<a href="javascript:;" onclick="window.open('adm/unidade_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a>
		</td>
	<tr>		
		<td><input type="submit" name="ok" value="ok"></td>		
	</tr>	
		</form>
	</tr>
</table><br id="no_print">
<%if strcd_unidade <> "" Then

for i = 1 to 3%>
<table class="txt" border="0" width="600" align="center">	
				<%condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				'strsql ="up_patrimonio_lista @ordem= '"&ordem&"', @condicao = '"&condicao&"'"
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' AND cD_preventiva = 1 order by cd_rack, nm_equipamento"
			  	Set rs_patrimonio = dbconn.execute(strsql)
				
				Do while not rs_patrimonio.EOF
				cd_pat_codigo = rs_patrimonio("cd_codigo")
				cd_patrimonio = rs_patrimonio("cd_patrimonio")
				cd_equipamento = rs_patrimonio("cd_equipamento")
				nm_equipamento = rs_patrimonio("nm_equipamento")
				cd_marca = rs_patrimonio("cd_marca")
				nm_marca = rs_patrimonio("nm_marca")
				cd_ns = rs_patrimonio("cd_ns")
				cd_pn = rs_patrimonio("cd_pn")
				nm_tipo = rs_patrimonio("nm_tipo")
				cd_especialidade = rs_patrimonio("cd_especialidade")
				cd_unidade = rs_patrimonio("cd_unidade")
					strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" AND cd_status <> 0"
					Set rs_uni = dbconn.execute(strsql)
					while not rs_uni.EOF
						nm_unidade = rs_uni("nm_unidade")
					rs_uni.movenext
					wend
					
					
				nm_sigla = rs_patrimonio("nm_sigla")
				cd_rack = rs_patrimonio("cd_rack")
				num_hospital =  rs_patrimonio("num_hospital")
				cd_rack = rs_patrimonio("cd_rack")
				num_hospital = rs_patrimonio("num_hospital")
				dt_data = rs_patrimonio("dt_data")
				
				cd_preventiva = rs_patrimonio("cd_preventiva")
				dt_periodicidade_prev = rs_patrimonio("dt_periodicidade_prev")
				
				cd_calibracao = rs_patrimonio("cd_calibracao")
				dt_periodicidade_cal = rs_patrimonio("dt_periodicidade_cal")
				
				cd_seg_elet = rs_patrimonio("cd_seg_elet")
				dt_periodicidade_elet = rs_patrimonio("dt_periodicidade_elet")
				
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
					strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"' "
				  	Set rs_espec = dbconn.execute(strsql_espec)
						if not rs_espec.EOF then
							nm_especialidade = rs_espec("nm_especialidade")
						end if
				
				if i = 1 then
					tipo_servico = "preventiva"
				elseif i = 2 then
					tipo_servico = "calibra�ao"
				elseif i = 3 then
					tipo_servico = "seguran�a el�trica"
				end if%>
	<%if cabeca = 0 then%>
	<tr>
		<td><img src="../imagens/logotipo_sao_luiz.gif" alt="" width="100" height="30" border="0"></td>
		<td colspan="13" align="center" valign="middle">PLANEJAMENTO DE SERVI�OS REFERENTE A <%=Ucase(tipo_servico)%> ANO <%=ano_sel%><br>VDLAP</td>
		<td colspan="3" align="right"><img src="../imagens/logotipo_vdlap.gif" alt="" width="80" height="40" border="0"></td>
	</tr>
	<tr bgcolor="#C0C0C0">
		<td align="center" valign="bottom"><b>EQUIPAMENTO</b></td>
		<td align="center" valign="bottom"><b>CONTROLE DE<br>MANUTEN��O</b></td>
		<td align="center" valign="bottom"><b>MARCA</b></td>
		<td align="center" valign="bottom"><b>MODELO</b></td>
		<td align="center" valign="bottom"><b>PERIODICIDADE</b></td>		
		<td align="center" valign="bottom"><b>JAN</b></td>
		<td align="center" valign="bottom"><b>FEV</b></td>
		<td align="center" valign="bottom"><b>MAR</b></td>
		<td align="center" valign="bottom"><b>ABR</b></td>
		<td align="center" valign="bottom"><b>MAI</b></td>
		<td align="center" valign="bottom"><b>JUN</b></td>
		<td align="center" valign="bottom"><b>JUL</b></td>
		<td align="center" valign="bottom"><b>AGO</b></td>
		<td align="center" valign="bottom"><b>SET</b></td>
		<td align="center" valign="bottom"><b>OUT</b></td>
		<td align="center" valign="bottom"><b>NOV</b></td>
		<td align="center" valign="bottom"><b>DEZ</b></td>
	</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');">
		<td valign="top" align="left" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=nm_equipamento%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%="0"&cd_patrimonio%></td>				
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=nm_marca%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=cd_pn%></td>
		<%if i = 1 then%>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=dt_periodicidade_prev%> <%if dt_periodicidade_prev <> "" Then response.write("meses") end if%></td>
		<%elseif i = 2 then%>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=dt_periodicidade_cal%> <%if dt_periodicidade_cal <> "" Then response.write("meses") end if%></td>
		<%elseif i = 3 then%>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=dt_periodicidade_elet%> <%if dt_periodicidade_elet <> "" Then response.write("meses") end if%></td>
		<%end if%>
	</tr>	
	<%
	cabeca = 1
	nm_especialidade = ""
	num_lista = num_lista + 1
	rs_patrimonio.movenext
	loop
	%>
	<tr>
		<td colspan="5">ITA20211.PQ.006 - ANEXO 3 - PLANEJAMENTO DE SERVI�OS - VDLAP<p>OBS. As datas marcadas s�o referentes a <%=tipo_servico%></p></td>
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td align="center" colspan="14"></td>
		<td align="center" colspan="3"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>
	<tr>
		<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
	</tr><tr><td colspan="100%">&nbsp;</td></tr>
	


</table><p style="page-break-after:always;"></p>

<%cabeca = 0
next
	
end if%>

</body>
</html>
