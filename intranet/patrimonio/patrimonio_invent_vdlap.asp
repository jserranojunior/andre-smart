<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
cd_user = session("cd_codigo")
pasta_loc = "patrimonio\"
arquivo_loc = "patrimonio_invent_vdlap.asp"

cd_pat_codigo = request("cd_pat_codigo")
strcd_unidade  =  request("cd_unidade")
mensagem = request("mensagem")

dt_hoje_dia = day(now)
dt_hoje_mes = month(now)
dt_hoje_ano = year(now)
dt_hoje = dt_hoje_mes&"/"&dt_hoje_dia&"/"&dt_hoje_ano


cd_ordem = request("cd_ordem")
	if int(cd_ordem) = 1 then
		cat_check = "selected"
		ordem = "cd_posicao"
		ordem_e = "cd_posicao"
	elseif int(cd_ordem) = 2 then
		cat_equipamento = "selected"
		ordem = "nm_equipamento_novo"
		ordem_e = "nm_equipamento_novo"
	elseif int(cd_ordem) = 3 then
		cat_patrimonio = "selected"
		ordem = "nm_tipo,no_patrimonio"
		ordem_e = "nm_tipo,no_patrimonio"
	elseif int(cd_ordem) = 4 then
		cat_marca = "selected"
		ordem = "nm_marca"
		ordem_e = "nm_marca"
	elseif int(cd_ordem) = 5 then
		cat_unidade = "selected"
		ordem = "nm_unidade"
		ordem_e = "nm_sigla_ori"
		elseif int(cd_ordem) = 6 then
		cat_situacao = "selected"
		ordem = "nm_estado"
		ordem_e = "nm_estado"
	end if

cd_itens = request("cd_itens")
	if cd_itens = "E" Then
		itens = " Where nm_tipo = 'E'"
		ck_item_E = "selected"
	elseif cd_itens = "I" Then
		itens = " Where nm_tipo = 'I'"
		ck_item_I = "selected"
	elseif cd_itens = "O" Then
		itens = " Where nm_tipo = 'O'"
		ck_item_O = "selected"
	else
		itens = ""
		ck_item = "selected"
	end if
	
cd_nao_constar = request("cd_nao_constar")
	if int(cd_nao_constar) = 2 then
		if itens <> "" Then
			var_constar = " AND "
		else
			var_constar = " Where "
		end if
		strcd_constar = var_constar&" nao_constar <> 2 "
	end if

cd_descarte = request("cd_descarte")


	if int(cd_descarte) = 1 then
		if itens <> "" Then
			var_descarte = " AND "
		else
			var_descarte = " Where "
		end if
		strcd_descarte = var_descarte&" cd_status > 1 "
		strtitulo_busca = "<br><br>DESCARTADOS"
		
	Else
		if itens <> "" Then
			var_descarte = " AND "
		else
			var_descarte = " Where "
		end if
		strcd_descarte = var_descarte&" cd_status = 1"
	end if
	
	
%>
<head>
	<title>Untitled</title>
</head>

<body><br class="no_print">

<table align="center" class="no_print" style="border:1 solid blue;">
<!--#include file="../includes/arquivo_loc.asp"-->

	<form action="patrimonio.asp" method="post">
	<input type="hidden" name="tipo" value="invent_vdlap">
	<tr style="border:1 solid orange; background-color:silver;">
		<td colspan="100%" align="center"><B>INVENTÁRIO DE EQUIPAMENTOS VDLAP</B></td>
	</tr>
	<tr>
		<td>&nbsp; Tipo:</td>
		<td><select name="cd_itens" class="inputs">				
				<option value="" <%=ck_item%>>Todos</option>
				<option value="E" <%=ck_item_E%>>Equipamentos</option>
				<option value="I" <%=ck_item_I%>>Instrumentos</option>
				<option value="O" <%=ck_item_O%>>Óticas</option>
			</select>
		</td>
	</tr>
	<tr>
	    <td>&nbsp; Ordem:</td>			
	    <td><select name="cd_ordem" class="inputs">
						<option value="1" <%=cat_check%>>Categoria</option>
						<option value="2" <%=cat_equipamento%>>Descrição</option>
						<option value="3" <%=cat_patrimonio%>>Patrimonio</option>
						<option value="4" <%=cat_marca%>>Marca</option>
						<option value="5" <%=cat_unidade%>>Unidade</option>
						<option value="6" <%=cat_situacao%>>Situação</option>
						
						</select>
						<!--a href="javascript:;" onclick="window.open('adm/unidade_cad.asp?janela=1', 'principal', 'width=500, height=400'); return false;">+</a-->
		</td>
	</tr>
	<tr>
		<td valign="top">&nbsp; Detalhes</td>
		<td valign="top"><input type="checkbox" name="cd_nao_constar" value="2" <%if cd_nao_constar = 2 then%>Checked<%end if%>> Esconder os "não constar" <br>(cabeçote, carrinho etc)</td>
	</tr>
	<tr>
		<td valign="top">&nbsp; Descarte</td>
		<td valign="top"><input type="checkbox" name="cd_descarte" value="1" <%if cd_descarte = 1 then%>Checked<%end if%>> Mostrar os descartes</td>
	</tr>
	<tr>		
		<td><input type="submit" name="ok" value="ok"></td>
		<td align="center"><img src="../imagens/ic_print.gif" alt="" width="24" height="24" border="0" onClick="window.print();"> &nbsp; <img src="../imagens/ic_print_view.gif" alt="" width="24" height="26" border="0" onclick="visualizarImpressao();"> &nbsp; &nbsp; &nbsp; &nbsp;</td>		
	</tr>
	<tr>
		<td valign="top" align="right"><b style="color:red;">Atenção:</b></td>
		<td valign="top"><b style="color:red;">Imprimir com orientação RETRATO<br> em folha A4.</b></td>
	</tr>
		</form>
	</tr>
</table><br class="no_print">
<%if cd_ordem <> "" Then%>

<table class="txt" border="0" width="600" align="center" style="border-collapse:collapse; border:0px solid red;">	
	<%num_lista = 1
	conta_registro = 1
				condicao = " Where dt_descarte is null"
				ordem = " "&ordem&" "
				
				'strsql = "SELECT * FROM VIEW_patrimonio_lista WHERE cd_unidade = "&strcd_unidade&" AND nm_tipo = 'E' AND nao_constar <> 2 order by cd_posicao, nm_equipamento"
				strsql = "SELECT * FROM VIEW_patrimonio_lista "&itens&" "&strcd_constar&" "&strcd_descarte&" Order by "&ordem&""
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
				cd_nf = rs_patrimonio("cd_nf")
				cd_unidade = rs_patrimonio("cd_unidade")
					strsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&" AND cd_status <> 0"
					Set rs_uni = dbconn.execute(strsql)
					while not rs_uni.EOF
						nm_unidade = rs_uni("nm_unidade")
					rs_uni.movenext
					wend
					
				nm_sigla = rs_patrimonio("nm_sigla")
				cd_rack = rs_patrimonio("cd_rack")
				
				cd_unidade_comodato = rs_patrimonio("cd_unidade_comodato")
				'nm_unidade_comodato = rs_patrimonio("nm_unidade_comodato")
				nm_sigla_comodato = rs_patrimonio("nm_sigla_comodato")
				
				
				num_hospital =  rs_patrimonio("num_hospital")
				nm_estado = rs_patrimonio("nm_estado")
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
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr><td colspan="100%">&nbsp;</td></tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td colspan="5" align="center" valign="middle" style="font-family:arial; font-size:12px; color:black;"><b>INVENTÁRIO DE EQUIPAMENTOS - VDLAP<%=strtitulo_busca%></b></td>
		<td colspan="2" align="right">
			<img src="../imagens/logotipo_vdlap_bw.gif" alt="" width="88" height="43" border="0" class="ok_print">
			<img src="../imagens/logotipo_vdlap.gif" alt="" width="88" height="43" border="0" class="no_print"></td>
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
		<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Situação</b></td>
	</tr>
	<%end if%>
	<%
	strsconderLista_emprestimo = 0
	'******************************************
	'***** Esconde a linha do item emprestado ***
	'********************************************
		strsql = "SELECT * FROM TBL_patrimonio_emprestimos where cd_situacao >= 4 and cd_patrimonio = '"&cd_pat_codigo&"' and dt_emprestimo <= '"&dt_hoje&"' and dt_retorno is NULL OR cd_situacao >= 4 and cd_patrimonio = '"&cd_pat_codigo&"' and dt_emprestimo <= '"&dt_hoje&"' and dt_retorno >= '"&dt_hoje&"'"
				Set rs_pat_emprestimo = dbconn.execute(strsql)
				while not rs_pat_emprestimo.EOF
					cd_emprestimo = rs_pat_emprestimo("cd_codigo")
				rs_pat_emprestimo.movenext
				wend
				
				
		if cd_emprestimo = "" Then
	'********************************************
	'num_lista = num_lista+1%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="center" valign="bottom" style="border:0px solid white; border-collapse:collapse; background-color:white;">&nbsp;</td>
		<td align="center" valign="top" style="<%=estilo_corpo%>"><%'=conta_registro%>&nbsp;<%=num_lista%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_tipo&" "&prefx_pat(no_patrimonio)%></td>		
		<td valign="top" align= "left"  style="<%=estilo_corpo%>">&nbsp;<a href="javascript:viod(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_pat_codigo%>&acao=alterar&jan=1&caminho=patrimonio','Alterações','width=700,height=600,scrollbars=1'); return false;"><%=nm_equipamento%></a> <%'="&nbsp;- "&cd_posicao%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_marca%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_pn%></td>
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_ns%></td>		
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_sigla_comodato%></td>
		<!--td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_sigla%></td-->
		<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_estado%></td>
		
	</tr>
	<%end if%>
	
	<%conta_registro = conta_registro + 1
	
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
	
	cd_emprestimo = ""
	nm_especialidade = ""
	num_lista = num_lista + 1
	rs_patrimonio.movenext
	loop
	
	conta_registro = conta_registro + 2%>
	<tr><td>&nbsp;</td></tr>
	<tr>
		<td>&nbsp;</td>
		<td  colspan="3" style="border:1px solid black; border-collapse:collapse; font-family:arial; font-size:11px;">&nbsp;<b>TOTAL DE ITENS APURADOS: <%=num_lista-1%></b></td>
	</tr>
	
	
	<%
	
	um_lista_e = 1
	'conta_registro = 1
	cabeca = 0
	if strsconderLista_emprestimo = 1 then
	strsql="SELECT * FROM View_patrimonio_emprestimos WHERE (cd_situacao >= 4) AND (dt_emprestimo <= '"&dt_hoje&"') AND (dt_retorno IS NULL) AND cd_cancela is NULL OR (cd_situacao >= 4) AND (dt_emprestimo <= '"&dt_hoje&"') AND (dt_retorno >= '"&dt_hoje&"') AND cd_cancela is NULL Order by "&ordem_e&""
	Set rs = dbconn.execute(strsql)
	while not rs.eof
		cod_emprestimos = rs("cd_codigo")
		cd_patrimonio = rs("cd_patrimonio")
			lista_patr_emp = lista_patr_emp &";"& cd_patrimonio
			
		nm_tipo = rs("nm_tipo")
		no_patrimonio = rs("no_patrimonio")
		nm_equipamento = rs("nm_equipamento_novo")
		nm_marca = rs("nm_marca")
		cd_pn = rs("cd_pn")
		cd_ns = rs("cd_ns")
		nm_sigla_ori = rs("nm_sigla_ori")
		nm_sigla_des = rs("nm_sigla_des")
	
		lista_patr_emp = mid(lista_patr_emp,2,len(lista_patr_emp))
		if cod_emprestimos <> "" then%>	
			
			<%if cabeca = 0 then
			conta_registro = conta_registro + 9%>
			<tr><td colspan="100%">&nbsp;</td></tr>
			<tr><td colspan="100%">&nbsp;</td></tr>
			<tr><td colspan="100%">&nbsp;</td></tr>
			<tr><td colspan="100%">&nbsp;</td></tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td colspan="5" align="center" valign="middle" style="font-family:arial; font-size:12px; color:black;"><b>INVENTÁRIO DE EQUIPAMENTOS - VDLAP <br>:: Empréstimos ::</b></td>
				<td colspan="2" align="right">&nbsp;
					<!--img src="../imagens/logotipo_vdlap_bw.gif" alt="" width="88" height="43" border="0" class="ok_print">
					<img src="../imagens/logotipo_vdlap.gif" alt="" width="88" height="43" border="0" class="no_print"--></td>
			</tr>
			<tr><td colspan="100%">&nbsp;</td></tr>
			<tr>
				<td align="center" valign="bottom">&nbsp;</td>
				<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Item</b></td>
				<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Patrimonio</b></td>
				<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Descrição</b></td>
				<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Marca</b></td>
				<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Modelo</b></td>
				<td align="center" valign="bottom" style="<%=estilo_titulo%>"><b>Nº Série</b></td>
				<td align="center" valign="bottom" style="<%=estilo_titulo%>" colspan="2"><b>Unidade</b></td>
			</tr>
			<%end if%>
					<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
						<td align="center" valign="bottom" style="border:0px solid white; border-collapse:collapse; background-color:white;">&nbsp;</td>
						<td align="center" valign="top" style="<%=estilo_corpo%>">&nbsp;<%=num_lista%></td>
						<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_tipo&prefx_pat(no_patrimonio)%></td>		
						<td valign="top" align= "left"  style="<%=estilo_corpo%>">&nbsp;<a href="javascript:viod(0);" onclick="window.open('patrimonio/patrimonio_cad.asp?tipo=cadastro&cd_pat_codigo=<%=cd_patrimonio%>&acao=alterar&jan=1&caminho=patrimonio','Alterações','width=700,height=600,scrollbars=1'); return false;"><%=nm_equipamento%></a></td>
						<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=nm_marca%></td>
						<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_pn%></td>
						<td valign="top" align="center" style="<%=estilo_corpo%>">&nbsp;<%=cd_ns%></td>
						<td valign="top" align="center" style="<%=estilo_corpo%>" colspan="2">&nbsp;<%=nm_sigla_ori%> -> <%=nm_sigla_des%></td>					
					</tr>
			<%conta_registro = conta_registro + 1
			num_lista = num_lista + 1
			num_lista_e = num_lista_e + 1
			cabeca = cabeca + 1%>
		<%end if%>
	<%rs.movenext
	wend
	end if
	%>
	<tr><td>&nbsp;</td></tr>
	
	<%if strsconderLista_emprestimo = 1 then %>
	<tr>
		<td>&nbsp;</td>
		<td  colspan="3" style="border:1px solid black; border-collapse:collapse; font-family:arial; font-size:11px;">&nbsp;<b>TOTAL DE ITENS EM EMPRÉSTIMO: <%=num_lista_e-1%></b></td>
	</tr>
	<%end if%>
	
	
	<%while not conta_registro > 55 and strsconderLista_emprestimo = 1%>
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
		<td class="no_print"><img src="../imagens/px.gif" alt="" width="5" height="1" border="0"></td>
		<td class="ok_print"><img src="../imagens/px.gif" alt="" width="0" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td>	
		<td><img src="../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="170" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="62" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="115" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<!--td><img src="../imagens/blackdot.gif" alt="" width="75" height="1" border="0"></td-->		
	</tr>
	<tr><td colspan="100%">&nbsp;</td></tr>


</table>
<%end if%>
</body>
</html>
