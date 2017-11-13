
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"--> 
<!--#include file="../includes/inc_area_restrita.asp"-->   
<!--#include file="../css/estilo.asp"-->  
<style type="text/css" media="print">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.txt_ {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.textopequeno{color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
#no_print{ 
	visibility:hidden; 
	display: none;}
#ok_print{
	visibility:visible;
	display:block;}
BODY { 
   margin: 0 0 0 0px; 
   text-align: center;}
table{
	background-color: #ffffff; 
    border: 0px solid #cccccc;
	width: 200px;
	font-family: verdana;
	font-size:9px;
	text-decoration:none;}
.divrolagem {
	overflow: auto; 
    height: auto;
    width: auto;}
-->
</style>
 
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/formValidator.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<script language="javascript">
<!--
  function mOvr(src,clrOver) {
    if (!src.contains(event.fromElement)) {
	  src.style.cursor = 'hand';
	  src.bgColor = clrOver;
	}
  }
  function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	  src.style.cursor = 'default';
	  src.bgColor = clrIn;
	}
  }
// -->	
</script>
<SCRIPT LANGUAGE="javascript">
    {       
       objVld=getValidator('form'); 
	   objVld.addRule('num_os',['required','texto'],'"O.S."');
	   objVld.addRule('dt_os',['required','texto'],'"Data da solicitação"');
	   objVld.addRule('nm_solicitante',['required','texto'],'"Solicitante"');
	   objVld.addRule('nm_unidade',['required','texto'],'"Unidade"');
   	   objVld.addRule('num_qtd',['required','texto'],'"Qtd"');
	   objVld.addRule('cd_equipamento',['required','texto'],'"Equip/Instr"');
	   objVld.addRule('cd_marca',['required','texto'],'"Marca"');
	   objVld.addRule('motivo',['required','texto'],'"Motivo"');
	   objVld.addRule('dt_entrada',['required','texto'],'"Data da provicência"');
	   objVld.addRule('nm_avaliacao',['required','texto'],'"Avaliação"');
	   objVld.addRule('cd_fornecedor',['required','texto'],'"Fornecedor"');
	   objVld.addRule('nm_unidade',['required','texto'],'"Unidade"');
   }
</SCRIPT>

<%filtro = request("filtro")
if filtro = "" Then
filtro = "pendentes"
end if%>
<BODY> 
<table align="center" border="0">
	<tr>
		<td>
			<table width="650" border="0" cellspacing="2" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef">
				
					
				<tr id="no_print">		
						<td class="txt_cinza" colspan="12">
						<b>Manutenção &raquo; - <font color="red">Listagem <%=filtro%></font></b></td>
						<td class="txt_cinza" align="right" valign="top">O.S.:&nbsp;</td>
						<td class="txt_cinza" colspan="2">
						<form action="manutencao.asp" name="busca" id="busca">
						<input type="hidden" name="tipo" value="andamento">
						<input type="text" name="cd_codigo" size="6" maxlength="6">
						<input type="hidden" name="campo" value="num_os">
						<input type="hidden" name="ordem" value="1">		
						<input type="submit" value="ok">
						</form>
						</td>
					</tr>
					<tr id="no_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manutencao.asp" name="busca" id="busca">
					<input type="hidden" name="tipo" value="nova">
					<tr id="no_print"><td colspan="100%">NOVA O.S.:<input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok"></td></tr>
					</form>
					<tr id="no_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr id="no_print">
						<td colspan="9" class="textopequeno" bgcolor="#f5f5f5">
							<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
						    / <a href="manutencao.asp?tipo=pendencias">Pendentes</a> 
							/ <a href="manutencao.asp?tipo=pendencias&filtro=encerradas&strtop=20">Encerradas</a>
							/ <a href="manutencao.asp?tipo=pendencias&filtro=todas&strtop=20">Todas</a>
							/ <a href="manutencao.asp?tipo=relatorio">Relatório</a>
						</td>
				</tr>
				<%if org = "os" then
					color_os = "green"
				elseif org = "equip" then
					color_equip = "green"
				elseif org = "marca" then
					color_marca = "green"
				elseif org = "unid" then
					color_unid = "green"
				elseif org = "fornec" then
					color_fornec = "green"
				elseif org = "aval" then
					color_aval = "green"
				elseif org = "sit" then
					color_sit = "green"
				elseif org = "prev" then
					color_prev = "green"
				end if%>
				<tr id="no_print">
						<td colspan="100%" class="textopequeno" bgcolor="#f5f5f5">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						    / <font color="red">Ordem de PENDENCIAS: </font>
							<a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=os"><font color="<%=color_os%>">O.S.</font></a> /
							<a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=equip"><font color="<%=color_equip%>">Equip/Instr.</font></a> /
							<a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=marca"><font color="<%=color_marca%>">Marca</font></a> /
							<a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=unid"><font color="<%=color_unid%>">Unidade</font></a> / 
							<a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=fornec"><font color="<%=color_fornec%>">Ass. Técnica</font></a> / 
							<a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=prev"><font color="<%=color_prev%>">Prev. Retorno</font></a> /  
							<a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=aval"><font color="<%=color_aval%>">Avaliacao</font></a> / 
							<a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=sit"><font color="<%=color_sit%>">Situação</font></a> /
						</td>
				</tr>
				<!--tr><td colspan="100%">Select: "up_listagem <%=selecao%>, <%=tabela%>, @condicoes='"&condicoes&"', @criterios='"&criterios&"'"</td></tr-->
					<tr><td colspan="10" class="textopequeno">&nbsp;</td></tr>
				<%'end if%>
				<tr id="ok_print"><td colspan="100%" align="center" class="txt_titulo"><b>PENDÊNCIAS EM: <%=day(now)%>/<%=month(now)%>/<%=year(now)%></b>&nbsp;</td></tr>
				<tr>
					<td><img src="imagens/px.gif" alt="" width="1" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="1" height="2" border="0"></td>
					<!--td></td>
					<td><img src="imagens/px.gif" alt="" width="1" height="2" border="0"></td-->
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="1" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="160" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="15" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="5" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="0" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="75" height="2" border="0"></td>
					<td></td>
					<!--td><img src="imagens/px.gif" alt="" width="5" height="0" border="0"></td>
					<td></td-->
					<td><img src="imagens/px.gif" alt="" width="5" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="60" height=2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="95" height="2" border="0"></td>
				</tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr>
				
				<!--tr id="ok_print"><td colspan="100%" align="center" class="txt_titulo"><b>RELATÓRIO </b><br>&nbsp;</td></tr>
				<tr id="ok_print"><td colspan="100%" align="right" class="textopequeno">Data da emissão: <%=dia%>/<%=mes%>/<%=ano%></td></tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr-->
				
					<tr bgcolor="darkgray">
						<td><b>Nº</b></td>
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><b><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=os"">O.S.</a></b></td>
						<!--td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><b><a href="manutencao.asp?tipo=pendencias&filtro=pendentes">PROVID.</a></b></td-->	
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><b><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=envio">DATA ENVIO</a></b></td>
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=equip"><b>EQUIPAMENTO/<br>INSTRUMENTAL.</b></a></td>	
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><b>PAT</b></td>	
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<!--td><img src="px.gif" alt="" width="1" height="1" border="0"><b>PATRIM.</b></td>	
						<td><img src="imagens/px.gif" width="1" height="1"></td-->
						<td><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=marca"><b>MARCA.</b></a></td>	
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=unid"><b>UNID.</b></a></td>		
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<!--td><img src="px.gif" alt="" width="1" height="1" border="0">&nbsp;&nbsp;&nbsp;</td/-->
						<td><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=fornec"><b>ASSIST. TÉCNICA</b></a></td>
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=prev"><b>PREV. RETORNO</b></a></td>
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=aval"><b>AVALIAÇÃO</b></a></td>
						<td><img src="../imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=sit"><b>SITUAÇÃO</b></a></td>
						<%if strsession = 2 Then%>
						<td><img src="../imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?tipo=pendencias&filtro=pendentes&org=sit"><b>diferença</b></a></td>
						<%end if%>
				</tr>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr>
				<%'1º
					dia = day(now)
					mes = month(now)
					ano = year(now)
		
					if dia < 10 Then
					dia = "0"&dia
					End if
					if mes < 10 Then
					mes = "0"& mes
					End if
		
					data = ano&mes&dia
					hoje = dia&"/"&mes&"/"&ano

		cd_codigo = request("cd_codigo")
		
		strtop = request("strtop")
		mov = request("mov")
		'filtro = request("filtro")
		regf = request("regf")
		regl = request("regl")
		
		org = request("org")
		
		str_os = request("str_os")
		marca = request("marca")
		
		fornec = request("fornec")
		unidade = request("unidade")
		
		equip = request("equip")
		ass_tecnica = request("ass_tecnica")
		
		
		num = request("num")
		solicitante = "'plan2007'"
		
		if filtro = "" Then
			filtro = "pendentes "
		end if
		
		'filtro = "AND cd_fornecedor = 16"
		
		'*************************
		'*** Ordem de exibição ***
		'*************************
		
	if filtro <> "pendentes" then
	ordenacao = 1
	'cond_icoes = " "
	Else
	ordenacao = 1
	end if
	
		'ordenacao = 1
		
		do while not ordenacao > 1
		'response.write(ordenacao&"<br>")
		
		'*** 0 - Lista encerradas ou todas *********
		if ordenacao = 0 then
		cond_icoes = " "
		'cor_aviso = "yellow"
		cor_aviso = "black"
		ordenacao = 4
		
		'*** 1 - Em andamento - no prazo - OK! *********
		Elseif ordenacao = 1 then
		'cond_icoes = " AND hoje <= prazo"
		'cond_icoes = " AND (CAST(dt_prev_retorno AS datetime) <= CAST(CONVERT(VARCHAR(10), GETDATE(), 101) AS DATETIME)) AND fecha = 0 "
		cond_icoes = " AND (CONVERT(VARCHAR(8), GETDATE(), 112) <= CONVERT(varchar(8), dt_prev_retorno, 112)) OR (fecha = 0) AND (dt_prev_retorno IS NULL)"
		cor_aviso = "black"
		
		'*** 2 - Prazo vencido até 50% ***********
		elseif ordenacao = 22 then
		'cond_icoes = " AND hoje > prazo AND hoje <= prazo* 1.5 "
		cond_icoes = "  AND (CONVERT(VARCHAR(8), GETDATE(), 112) > CONVERT(varchar(8), dt_prev_retorno, 112)) AND (DATEDIFF(d, dt_prev_retorno, GETDATE()) <= prazo * 1.5)"
		cor_aviso = "red"
		
		'*** 3 - Crítico - Roxo  OK ******************
		Elseif ordenacao = 3 then
		'cond_icoes = " AND hoje > prazo* 1.5"
		cond_icoes = " AND (CONVERT(VARCHAR(8), GETDATE(), 112) > CONVERT(varchar(8), dt_prev_retorno, 112)) AND (DATEDIFF(d, dt_prev_retorno, GETDATE()) > prazo * 1.5)"
		cor_aviso = "purple"
		
		Elseif ordenacao = 4 then
		'*** 4 - Retirar no fornecedor - Verde ***
		cond_icoes = " AND prazo is NULL "
		cor_aviso = "blue"
		'str_situacao = "'Pronto para retirar'"
		'cond_icoes = " AND nm_situacao = '"&str_situacao&"' "
		'*****************************************
		end if
		
	'end if
		
		criterios = " Order by num_os DESC "

		if filtro = "pendentes" AND org = ""  AND forn = "" OR filtro = ""  AND forn = "" Then
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&""' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&""
		condicoes = " Where fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&""
		'condicoes = " Where fecha = 0 "'OR fecha IS NULL"
		strtop = "1000"
		'response.write("1.1")
		
		elseif filtro = "pendentes" AND org = "fornec"  AND forn = "" Then
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&""
		condicoes = " Where fecha = 0 AND nm_solicitante <> '"&solicitante&"'"' "&cond_icoes&""
		strtop = "1000"
		criterios = " Order by nm_fornecedor ASC"
		'response.write("1.2")
		
		elseif filtro = "pendentes" AND org = "unid"  AND forn = "" Then
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&""
		condicoes = " Where fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&""
		strtop = "1000"
		criterios = " Order by nm_unidade ASC"
		'response.write("1.3")
		
		elseif filtro = "pendentes" AND org = "os"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" "'prev_retorno < data_hoje "
		strtop = "1000"
		criterios = " Order by num_os ASC"
		'response.write("1.4")

		elseif filtro = "pendentes" AND org = "equip"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" "
		strtop = "1000"
		criterios = " Order by nm_equipamento ASC"
		'response.write("1.5")
		
		elseif filtro = "pendentes" AND org = "marca"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&""
		strtop = "1000"
		criterios = " Order by nm_marca ASC"
		'response.write("1.6")
		
		elseif filtro = "pendentes" AND org = "aval"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" "
		strtop = "1000"
		criterios = " Order by nm_avaliacao ASC"
		'response.write("1.7")
		
		elseif filtro = "pendentes" AND org = "sit"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&""
		strtop = "1000"
		criterios = " Order by nm_situacao ASC"
		'response.write("1.8")
		
		elseif filtro = "pendentes" AND org = "prev"  AND forn = "" Then
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' AND dt_prev_retorno IS NOT NULL OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' AND dt_prev_retorno IS NOT NULL"
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" "
		strtop = "1000"
		criterios = " Order by dt_prev_retorno"
		'response.write("1.9")
		
		elseif filtro = "pendentes" AND org = "envio"  AND forn = "" Then
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' AND dt_prev_retorno IS NOT NULL OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' AND dt_prev_retorno IS NOT NULL"
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&cond_icoes&""
		strtop = "1000"
		criterios = " Order by dt_manut_orc"
		'response.write("2.0")
		
		elseif filtro = "relatorio" AND org = "" AND forn = "" Then
		nm_fornecedor = "'Erwin Guth'"
		mes_entrada = "'4'"
		ano_entrada = "'2007'" 
		
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		condicoes = " Where nm_fornecedor = '"&nm_fornecedor&"' AND MONTH(dt_entrada) > '"&mes_entrada&"' AND YEAR(dt_entrada) = '"&ano_entrada&"'"
		strtop = "1000"
		criterios = " Order by num_os desc"
		'response.write("10.0")

		Elseif filtro = "encerradas" Then
		condicoes = " Where fecha = 1 "
		'response.write("3.1")

		Elseif filtro = "todas" Then
		condicoes = " "
		'response.write("4.1")
		'response.write("Ordem padrão")

		Else
		condicoes = " "
		'response.write("4.1")
		'response.write("Ordem padrão")

		End if
		txt = "nada"
		selecao = " TOP "&strtop&" * "
		Tabela = " [dbo].[VIEW_os_lista] "
		%>

<%'2º
				'***** Contador Geral de registros********************
				xsql_contador ="up_contador_geral @tabela='"&tabela&"', @condicoes='"&condicoes&"'"
				Set rs_contador = dbconn.execute(xsql_contador)
				
				if Not rs_contador.EOF Then
				nreg = rs_contador("nreg")
				End if
				'***** Fim do contador *******************************
				
					If mov = "retr" Then
						if filtro = "pendentes" OR filtro = "" Then
						condicoes = " Where num_os > "&regf&" AND nm_solicitante <> '"&solicitante&"' OR fecha IS NULL OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "
						'condicoes = " Where num_os > "&regf&" AND fecha IS NULL OR fecha = 0 "
						Elseif filtro = "encerradas" Then
						condicoes = " Where num_os > "&regf&" AND fecha = 1 "
						Elseif filtro = "todas" Then
						condicoes = " Where num_os > "&regf&""
						End if
						txt = "retr"
					Elseif mov = "avan" Then
						if filtro = "pendentes" OR filtro = "" Then
						condicoes = " Where num_os < "&regl&" AND nm_solicitante <> '"&solicitante&"' OR fecha IS NULL OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "
						'condicoes = " Where num_os < "&regl&" AND fecha IS NULL OR fecha = 0 "
						Elseif filtro = "encerradas" Then
						condicoes = " Where num_os < "&regl&" AND fecha = 1 "
						Elseif filtro = "todas" Then
						condicoes = " Where num_os < "&regl&""
						End if
						txt = "avan"
					End if
				
				%>
				 				
<%'3º
				xsql = "up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
				Set rs = dbconn.execute(xsql)
				
				if num = 0 Then 
				num = 1
				End if
				item = 1
				
				
			Do while not rs.EOF
				
				num_os = rs("num_os")
								
				if num < 10 Then
				num = "0" & num
				End if
				
				
				'digito = rs("digito")
				dt_os = rs("dt_os")
				nm_equipamento = rs("nm_equipamento")
				cd_ns = rs("cd_ns")
				cd_patrimonio = rs("cd_patrimonio")
					if cd_patrimonio <> "" Then
					cd_patrimonio = Replace(cd_patrimonio,"VDLAP","")
					cd_patrimonio = Replace(cd_patrimonio,"vdlap","")
					end if
				nm_marca = rs("nm_marca")
				nm_unidade = rs("nm_unidade")
				
				dt_entrada = rs("dt_entrada")
					strdia = zero(day(dt_entrada))						'if 	strdia < 10 then						'strdia = "0"&strdia						'end if
					strmes = zero(month(dt_entrada))						'if 	strmes < 10 then						'strmes = "0"&strmes						'end if
					strano = year(dt_entrada)
				dt_entrada = strdia&"/"&strmes&"/"&strano
				nm_fornecedor = rs("nm_fornecedor")
				cd_fornecedor = rs("cd_fornecedor")
				cd_avaliacao = rs("cd_avaliacao")
				cd_liberacao = rs("cd_liberacao")
				cd_codigo = rs("cd_codigo")
				nm_situacao = rs("nm_situacao")
				
				dt_manut_orc = rs("dt_manut_orc")
					manut_dia = zero(day(dt_manut_orc))					'manut_dia = day(dt_manut_orc)					'		if 	manut_dia < 10 then											'manut_dia = "0"&manut_dia											'end if
					manut_mes = zero(month(dt_manut_orc))					'manut_mes = month(dt_manut_orc)											'if 	manut_mes < 10 then											'manut_mes = "0"&manut_mes											'end if
					manut_ano = year(dt_manut_orc)

				dt_manut_orc = manut_dia&"/"&manut_mes&"/"&manut_ano
				difer = rs("hoje")
				prazo = rs("prazo")
				
				dt_prev_retorno = rs("dt_prev_retorno")
					if year(dt_prev_retorno) = 1950 Then
						'dt_prev_retorno = ""
					end if
					
						if dt_prev_retorno <> "" Then
							dia_prev_retorno = day(dt_prev_retorno)
							mes_prev_retorno = month(dt_prev_retorno)
							ano_prev_retorno = year(dt_prev_retorno)
								dt_prev_retorno = zero(dia_prev_retorno)&"/"&zero(mes_prev_retorno)&"/"&ano_prev_retorno
						end if
						
						if dt_prev_retorno <> "" Then
						'dt_diferenca = DateDiff("d", dt_manut_orc, dt_prev_retorno) 
						dt_diferenca = DateDiff("d", dt_manut_orc, hoje) 
						end if
				
				'dt_prev_retorno = day(dt_manut_orc)+prazo
				
				'****************************************************************************************
				'*** Analiza uma data e soma x dias                                                   ***
				'****************************************************************************************
				if not dt_manut_orc = NULL Then
				data_entrada = dt_manut_orc'"01/12/2008"'Versão apenas para teste
				'data_entrada = "01/12/2000"'Versão apenas para teste
				
				
				'prazo = 40
				if prazo < 1 Then
				prazo = 1
				end if
				
				ultimo_dia = ultimodiames(month(data_entrada),year(data_entrada))'Verifica o ultimo dia do mes
				soma_dias = day(data_entrada)+prazo
				soma_mes = month(data_entrada)
				soma_ano = year(data_entrada)
				
					if soma_dias > ultimo_dia then
						soma_dias = soma_dias - ultimo_dia
						soma_mes = soma_mes + 1
							if soma_mes > 12 then
							soma_mes = soma_mes - 12
							soma_ano = soma_ano + 1
							end if
						soma_ano = soma_ano
					end if
				
				dt_prev_retorno = soma_dias&"/"&soma_mes&"/"&soma_ano
					soma_dias = ""
					soma_mes = ""
					soma_ano = ""
				end if
				'***********************************************
				'	 Verifica a data da previsão de retorno
				'************************************************
				'Transfoma as datas em números
					prev_retorno = ano_prev_retorno&zero(mes_prev_retorno)&zero(dia_prev_retorno)
					Data_hoje = ano&mes&dia
						
						
						if dt_diferenca > int(60) then
						maior = " - É"
						cor_aviso = "green"
						Else
						cor_aviso = "black"
						end if
						
						'if nm_situacao = "Pronto para retirar" Then
						'cor_aviso = "#339966"
						'Elseif prev_retorno < data_hoje AND prev_retorno <> ""  AND prev_retorno <> "19500101" AND fecha <> 1 then
						'cor_aviso = "red"
						'Elseif dt_diferenca >= int(30)  and fecha <> 1 then
						''maior = " - É"
						'cor_aviso = "#660099"
						'Else
						'cor_aviso = "black"
						'End if
				
				
				
				
				sequencia = 1
				fecha = rs("fecha")
				
					'****** Mostra os nomes das avaliacaoes ***
						strsql ="SELECT * FROM TBL_avaliacao WHERE cd_codigo='"&cd_avaliacao&"'"' up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
						Set rs_aval = dbconn.execute(strsql)
					Do while not rs_aval.EOF
						nm_avaliacao = rs_aval("nm_avaliacao")
					rs_aval.movenext
					loop
					'********************************************
				
				if nm_avaliacao = "manutencão" Then
					nm_avaliacao = "manutenção"
				Elseif nm_avaliaca0 = "substituicao" then
					nm_avaliacao = "substituição"
				End if
				
				if fecha = "1" then
				nm_situacao = "Encerrada"
				end if
				
				'***** Paginador *************************************
				
				paginas = nreg/strtop
				pags = Int(paginas)
				if paginas > pags then
				pags = pags + 1
				paginas = pags 'ags + pg
				End if
				
				If fecha = "1" Then
				nm_situacao = "Encerrada"
				End if
				%>
				<%if cor_aviso = "black" AND item = 1 Then%><!--tr><td colspan="100%" align="center" id="ok_print">Em andamento</td></tr--><%end if%>
				<%if cor_aviso = "red" AND item = 1 Then%><tr><td colspan="100%" align="center" id="ok_print">________________________________________________ <b>Prazo do fornecedor vencido</b> ________________________________________________</td></tr><%end if%>
				<%if cor_aviso = "purple" AND item = 1 Then%><tr><td colspan="100%" align="center" id="ok_print">_________________________________________________________ <b>Crítico</b> _________________________________________________________</td></tr><%end if%>
				<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:;" onmouseout="mOut(this,'');">
						<td valign="top" align="center" bgcolor="darkgray"><b><%=num%></b></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><a href="manutencao.asp?tipo=andamento&cd_codigo=<%=cd_codigo%>&campo=cd_codigo"><font color="<%=cor_aviso%>"><%=num_os%></font></a></td>
						<!--td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><a href="../manut_os_andamento.asp?cd_codigo=<%=cd_codigo%>&campo=cd_codigo"><font color="<%=cor_aviso%>"><%'=dt_os%><%=dt_entrada%></font></a></td-->
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><a href="manutencao.asp?tipo=andamento&cd_codigo=<%=cd_codigo%>&campo=cd_codigo"><font color="<%=cor_aviso%>"><%'=dt_entrada%><%=dt_manut_orc%></font></a></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%=Left(nm_equipamento,37)%></font></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<!--td valign="top"><font color="<%=cor_aviso%>"><%'=(cd_ns)%></font></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td-->
						<td valign="top" align="right"><font color="<%=cor_aviso%>"><%=cd_patrimonio%></font></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%=Left(nm_marca,10)%></font></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%=nm_unidade%></font></td>		
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<!--td valign="top">&nbsp;</td/-->
						<td valign="top"><a href="javascript:;" onclick="window.open('manutencao/fornecedor_cad.asp?cd_codigo=<%=cd_fornecedor%>&janela=pop_up', 'Fornecedor', 'width=517, height=290'); return false;"><font color="<%=cor_aviso%>"><%=Left(nm_fornecedor,12)%></font></a></td>		
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><a href="manutencao.asp?tipo=andamento&cd_codigo=<%=cd_codigo%>&campo=cd_codigo"><font color="<%=cor_aviso%>"><%=dt_prev_retorno%></font></a></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><a href="manutencao.asp?tipo=andamento&cd_codigo=<%=cd_codigo%>&campo=cd_codigo"><font color="<%=cor_aviso%>"><%=nm_avaliacao%></font></a></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td><%if nm_situacao="Em estudo" Then%><%neg_1="<b>"%><%neg_2="</b>"%><%else%><%neg_1=""%><%neg_2=""%><%end if%>
						<td valign="top"><a href="manutencao.asp?tipo=andamento&cd_codigo=<%=cd_codigo%>&campo=cd_codigo"><%=neg_1%><font color="<%=cor_aviso%>"><%=nm_situacao%></font><%=neg_2%></a></td>
						<%if strsession = "2" Then%>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%=difer%>-<%=prazo%></font></td>
						<td valign="top" id="no_print"><font color="<%=cor_aviso%>"><%=dt_prev_retorno%></font></td>
						<td valign="top" id="no_print"><%=prev_retorno&"-"%><%=data_hoje&"="%><%=dif_datas%>-<%=cor_aviso%></td>
						<!--td id="no_print"><%=dt_diferenca%><%'=maior%></td-->
						<%end if%>
				</tr>
				
				
				<%
				if regf > regl Then
				regf = regl
				'Elseif regf  > regl Then
				'regf = ""
				End if
				
				regl = num_os
				
				rs.movenext
				s = "s"
				
				'aqui!
				if filtro <> "pendentes" then
				'nreg = num
				s = ""
				end if
				
				num = num + 1
				item = item + 1
				'order = order + 1
				prev_retorno = ""
				nm_situacao = "" 
				cor_aviso = ""
				maior = ""
				
				loop
				
				'Reinicia contagem de O.S.
				if ordenacao <> order then
				'num = 1
				item = 1
				end if
				%>
				<%if ordenacao <> order then%>
				<tr><td colspan="100%"><!--Pula  linha --><br><br><b></b></td></tr>
				<%end if%>
				<%
ordenacao = ordenacao + 1%>
				<!--tr><td colspan="100%">SELECT <%=selecao%> FROM <%=tabela%><%=condicoes%><%=criterios%></td></tr-->
				<%
order = ordenacao
loop%>


				<tr class="textopequeno" id="no_print"><td colspan="7">&nbsp;</td></tr>
					<tr id="no_print">
						<td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
							<!--/ <a href="ordemservico_geracao.asp">Nova O.S. </a>
						    / <a href="ordemservico_listagem.asp">Listagem </a> 
							</ <a href="manutencao_pend.asp?sig=vazio&campo=dt_encerramento">Pendentes<%=fecha%></a-->
							
							< <%=nreg%> registros em <%=paginas%> página<%=s%> ></td>
						<td colspan="6" class="textopequeno" bgcolor="#f5f5f5"> 
						<%'If paginas > 10 Then
						'paginas = 10
						'continuacao = "Mais"
						'End if
						'pag = 1
				
						 'campo = request("campo")
						 'ordem = request("ordem")
						 pagina = request("pagina")
				
				
						'Mostra o número atual da página
						 if pagina = "" Then
						 pagina = 1
						 Else
						 pagina = pagina + 1
						 End if
				
						 'if ordem = "0" AND strmove = "1" OR ordem = "" AND strmove = "1" Then
						 'pagina = 1
						 'Elseif ordem = "1" AND strmove = "0" Then
						 'pagina = 1
						 'End if%>
						<%If pagina = 1 Then%>
						<b><font color="#b7b7b7">Início</font></b>
						<%else%>
						<b><a href="manutencao.asp?strtop=<%=strtop%>&filtro=<%=filtro%>&regf=<%=regf%>&regl=<%=regl%>&mov=retr&pagina=<%'=pagina%>&num=1"">Início</a></b>
						<%End if%>
						&nbsp;&nbsp;&nbsp;
						<%If pagina = paginas Then%>
						<b><font color="#b7b7b7">Próxima</font></b>
						<%else%>
						<b><a href="manutencao.asp?strtop=<%=strtop%>&filtro=<%=filtro%>&regf=<%=regf%>&regl=<%=regl%>&mov=avan&pagina=<%=pagina%>&num=<%=num%>">Próxima</a> </b>
						<%End if%>
						<td class="textopequeno" bgcolor="#f5f5f5" colspan="2">Pág: (<%=pagina%>)</td>
					</tr>
					<!--tr class="textopequeno" id="no_print"><td colspan="100%"><&nbsp;<%=regf%>/<%=regl%> - <%=condicoes%><%=criterios%> *** <%=txt%>></td></tr-->
					<%'if strsession = "2" Then%>
					<tr id="ok_print"><td colspan="100%">
					<%do while not num = "45"%> <br>
					<%num = num + 1
					loop%>
					
						<table border="1" cellpadding="0" cellspacing="1" bordercolor="#808080" bordercolordark="#808080" bordercolorlight="#808080">
							<tr><td>&nbsp;LEGENDA</td></tr>
							<tr>
								<td>&nbsp;<font color="black"><b>O</b></font> - Em andamento<br>
								&nbsp;<font color="red"><b>O</b></font> - Prazo do fornecedor vencido<br>
								&nbsp;<font color="purple"><b>O</b></font> - Crítico<br>
								<&nbsp;<font color="blue"><b>O</b></font> - Crítico>
								
								
								</td>
							</tr>
						</table-->
					</td></tr>
					<%'end if%>
				</table></td></tr></table>
				
				




</BODY>
</HTML>
