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

<body><br>

<table align="center" id="no_print">
	<form action="patrimonio.asp" method="post">
	<input type="hidden" name="tipo" value="invent_hosp">
	
	
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
</table><br>
<%if strcd_unidade <> "" Then%>
<table class="txt" border="0" width="600" align="center">	
	<!--tr>
		<td colspan="100%"><img src="../imagens/px.gif" alt="" width="100%" height="1" border="0"></td>
	</tr-->
	<%num_lista = 1			
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				'strsql ="up_patrimonio_lista @ordem= '"&ordem&"', @condicao = '"&condicao&"'"
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' order by cd_rack, nm_equipamento"
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
				
				cd_calibracao = rs_patrimonio("cd_calibracao")
				
				cd_seg_elet = rs_patrimonio("cd_seg_elet")
				
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
					strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"' "
				  	Set rs_espec = dbconn.execute(strsql_espec)
						if not rs_espec.EOF then
							nm_especialidade = rs_espec("nm_especialidade")
						end if
				%>
	<%if cabeca = 0 then%>
	<tr>
		<td colspan="2"><img src="../imagens/logotipo_sao_luiz.gif" alt="" width="100" height="30" border="0"></td>
		<td colspan="5" align="center" valign="middle">INVENTÁRIO DE EQUIPAMENTOS - VDLAP</td>
		<td colspan="2" align="right"><img src="../imagens/logotipo_vdlap.gif" alt="" width="80" height="40" border="0"></td>
	</tr>
	<tr bgcolor="#C0C0C0">
		<td align="center" valign="bottom"><b>Item</b></td>
		<td align="center" valign="bottom"><b>Controle de<br>Manutenção</b></td>
		<td align="center" valign="bottom"><b>Descrição</b></td>
		<td align="center" valign="bottom"><b>Marca</b></td>
		<td align="center" valign="bottom"><b>Modelo</b></td>
		<td align="center" valign="bottom"><b>Nº Série</b></td>
		<td align="center" valign="bottom"><b>Preventiva</b></td>
		<td align="center" valign="bottom"><b>Calibração</b></td>
		<td align="center" valign="bottom"><b>Seg. Eletrica</b></td>
	</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');"   onmouseout="mOut(this,'');">
		<td align="center" valign="top" bgcolor="#C0C0C0">&nbsp;<%=num_lista%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%="0"&cd_patrimonio%></td>		
		<td valign="top" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=nm_equipamento%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=nm_marca%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=cd_pn%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%=cd_ns%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%if cd_preventiva = 1 then response.write("Sim") Else response.write("Não Aplicável")end if%></td>		
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%if cd_calibracao = 1 then response.write("Sim") Else response.write("Não Aplicável")end if%></td>
		<td valign="top" align="center" onmouseover="mOvr(this,'#CFC8FF');"  onmouseout="mOut(this,'');">&nbsp;<%if cd_seg_elet = 1 then response.write("Sim") Else response.write("Não Aplicável")end if%></td>
	</tr>	
	<%
	cabeca = 1
	nm_especialidade = ""
	num_lista = num_lista + 1
	rs_patrimonio.movenext
	loop
	%>
	<tr>
		<td colspan="5">ITA20211.PQ.006 - ANEXO 2 - INVENTÁRIO DE EQUIPAMENTO VDLAP</td>
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td>&nbsp;</td>
		<td align="center" colspan="7"><b id="no_print"><%=nm_unidade%></b></td>
		<td align="center"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>
	<tr>
		<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td id="ok_print"><img src="../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
	</tr><tr><td colspan="100%">&nbsp;</td></tr>


</table>
<%end if%>

</body>
</html>
