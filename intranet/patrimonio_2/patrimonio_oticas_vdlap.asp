<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_pat_codigo = request("cd_pat_codigo")

strcd_unidade  =  request("cd_unidade")
cd_ordem = request("cd_ordem")
mensagem = request("mensagem")
cd_mostra_desc = request("cd_mostra_desc")
%>
<head>
	<title>Patrimônio</title>
</head>

<body><br id="no_print">

<table align="center" id="no_print" style="border:1 solid blue;">
	<%if session("cd_codigo") = 46 then%>
		<tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">patrimonio_oticas_vdlap.asp</span></td></tr>
	<%end if%>
	<form action="patrimonio.asp" method="post">
	<input type="hidden" name="tipo" value="oticas_vdlap">
	<tr style="border:1 solid orange; background-color:silver;">
		<td colspan="100%" align="center"><B>INVENTÁRIO DE ÓTICAS VDLAP</B></td>
	</tr>
	<tr>
	    <td>&nbsp; Ordem:</td>			
	    <td><select name="cd_ordem" class="inputs" onFocus="nextfield ='num_hospital';">
						
						<%if cd_ordem = "1" then
							cat_nome = "selected"
							ordem = "nm_equipamento"
						elseif cd_ordem = "2" then
							cat_check = "selected"
							ordem = "no_patrimonio"
						elseif cd_ordem = "3" then
							cat_marca = "selected"
							ordem = "nm_marca"
						elseif cd_ordem = "4" then
							cat_unidade = "selected"
							ordem = "nm_unidade"
						else
							cat_nada = "selected"
							ordem = ""
						end if%>
						<option value=""<%=cat_nada%>></option>
						<option value="1" <%=cat_nome%>>Descrição</option>
						<option value="2" <%=cat_patrimonio%>>Patrimonio</option-->
						<option value="3" <%=cat_marca%>>Marca</option>
						<option value="4" <%=cat_unidade%>>Unidade</option>
						</select>
						<!--a href="javascript:;" onclick="window.open('adm/unidade_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a-->
		</td>
	</tr>
	<tr>
		<td>&nbsp; Mostra descarte?</td>
		<%if cd_mostra_desc = "" OR cd_mostra_desc = "0" Then
			ck_desc0 = "checked"
			mostra_desc = " "
		elseif cd_mostra_desc = "1" Then
			ck_desc1 = "checked"
			mostra_desc = "AND nm_sigla <> 'desc'"
		end if%>
		<td><input type="radio" name="cd_mostra_desc" value="1" <%=ck_desc1%>>Não &nbsp; &nbsp; <input type="radio" name="cd_mostra_desc" value="0" <%=ck_desc0%>>Sim</td>
	</tr>
	
	<tr>		
		<td><input type="submit" name="ok" value="ok"></td>
		<td align="center"><img src="../imagens/ic_print.gif" alt="" width="24" height="24" border="0" onClick="window.print();"> &nbsp; <img src="../imagens/ic_print_view.gif" alt="" width="24" height="26" border="0" onclick="visualizarImpressao();"> &nbsp; &nbsp; &nbsp; &nbsp; </td>		
	</tr>	
		</form>
	</tr>	
</table><br id="no_print">
<%if cd_ordem <> "" Then%>

<table class="txt" border="0" width="600" align="center" style="border-collapse:collapse; border:0px solid white;">	
	<%num_lista = 1
	conta_registro = 1
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				'strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' AND nao_constar <> 1 order by cd_posicao, nm_equipamento"
				'strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE nm_tipo = 'O' AND nao_constar <> 1 "&mostra_desc&" order by "&ordem&""
				strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE nm_tipo = 'O' "&mostra_desc&" order by "&ordem&""
				'strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE nm_tipo = 'E' order by "&ordem&", nm_equipamento"
			  	Set rs_patrimonio = dbconn.execute(strsql)
				
				Do while not rs_patrimonio.EOF
				cd_pat_codigo = rs_patrimonio("cd_codigo")
				no_patrimonio = rs_patrimonio("no_patrimonio")
				cd_equipamento = rs_patrimonio("cd_equipamento")
				nm_equipamento = rs_patrimonio("nm_equipamento_novo")
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
				cd_posicao = rs_patrimonio("cd_posicao")
				dt_data = rs_patrimonio("dt_data")
				
				nao_constar = rs_patrimonio("nao_constar")
				
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
					strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"' "
				  	Set rs_espec = dbconn.execute(strsql_espec)
						if not rs_espec.EOF then
							nm_especialidade = rs_espec("nm_especialidade")
						end if%>
	<%if cabeca = 0 then
	estilo_titulo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:11px; color:black;"
	estilo_corpo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:11px;"%>
	<tr><td colspan="1000">&nbsp;</td></tr>
	<tr><td colspan="1000">&nbsp;</td></tr>
	<tr>
		<td>&nbsp;</td>
		<td>
		&nbsp;</td>
		<td colspan="5" align="center" valign="middle" style="font-family:arial; font-size:12px; color:black;"><b>INVENTÁRIO DE ÓTICAS - VDLAP</b></td>
		<td colspan="2" align="right">
			<img src="../imagens/logotipo_vdlap_bw.gif" alt="" width="88" height="43" border="0" id="ok_print">
			<img src="../imagens/logotipo_vdlap.gif" alt="" width="88" height="43" border="0" id="no_print"></td>
	</tr>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr>
		<td align="center" valign="bottom">&nbsp;</td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Item</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Patrimonio</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Descrição</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Marca</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Modelo</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Nº Série</b></td>
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Unidade</b></td>
	</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="center" valign="bottom" style="border:0px solid white; border-collapse:collapse; background-color:white;">&nbsp;</td>
		<td align="center" valign="top" style="<%=estilo_corpo%>"><%'=conta_registro%>&nbsp;<%=num_lista%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=prefx_pat(no_patrimonio)%></td>		
		<td valign="top" align= "left"  style="<%=estilo_corpo%>">&nbsp;<a href="javascript:viod(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar','Alterações','width=700','height=600'); return false;"><%=nm_equipamento%></a> <%'="&nbsp;- "&cd_posicao%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_marca%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_pn%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_ns%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_sigla%></td>
	</tr>	
	<%
	conta_registro = conta_registro + 1
	
	if int(conta_registro) > 55 then
		cabeca = 0
		conta_registro = 1
		pagina = pagina + 1%>
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td>&nbsp;</td>
		<td align="center" colspan="6">&nbsp;Pág.:<%=zero(pagina)%></td>
		<td align="center"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>
	<tr><td style="page-break-after:always;">&nbsp;</td></tr>
	<%else
		cabeca = 1
	end if
	
	nm_especialidade = ""
	num_lista = num_lista + 1
	rs_patrimonio.movenext
	loop
	
	conta_registro = conta_registro + 1%>	
	<tr><td>&nbsp;</td></tr>
	<tr>
		<td>&nbsp;</td>
		<td colspan="3" style="border:1px solid black; border-collapse:collapse; font-family:arial; font-size:11px;">&nbsp;<b>Total de óticas: <%=num_lista-1%></b></td>
	</tr>
	
	<%while not conta_registro > 55%>
		<tr><td align="center" colspan="100%" style="border:1px solid white; border-collapse:collapse; font-family:arial; font-size:11px;">&nbsp;<%'=conta_registro%></td></tr>
	<%conta_registro = conta_registro + 1
	Wend%>	
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td>&nbsp;</td>
		<td align="center" colspan="6">&nbsp;Pág.:<%=zero(pagina+1)%></td>
		<td align="center"><b><%=zero(day(now))&"/"&zero(month(now))&"/"&year(now)%></b></td>
	</tr>
	<tr>
		<td id="no_print"><img src="../imagens/px.gif" alt="" width="5" height="1" border="0"></td>
		<td id="ok_print"><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="52" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="83" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="67" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
		<!--td><img src="../imagens/blackdot.gif" alt="" width="75" height="1" border="0"></td-->		
	</tr>
	<tr><td colspan="100%">&nbsp;</td></tr>


</table>
<%end if%>
</body>
</html>
