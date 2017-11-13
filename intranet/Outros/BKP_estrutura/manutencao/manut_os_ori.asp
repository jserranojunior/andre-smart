
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"--> 
<!--#include file="../includes/inc_area_restrita.asp"-->   
   
<style type="text/css" media="print">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.txt_ {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
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

<BODY> 
<!--Mostra qual é o arquivo//-->
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

<%
cd_codigo = request("cd_codigo")

strtop = request("strtop")
mov = request("mov")
tipo = request("tipo")
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

if tipo = "" Then
	tipo = "pendentes"
end if

if strsession = "2" Then
'filtro = "AND cd_fornecedor = 16"
end if
criterios = " Order by num_os DESC "

		if tipo = "pendentes" AND org = ""  AND forn = "" OR tipo = ""  AND forn = "" Then
		condicoes = " Where fecha = 0 "
'		condicoes = " Where fecha IS NULL OR fecha = 0"
		strtop = "1000"
		'response.write("1.1")
		
		elseif tipo = "pendentes" AND org = "fornec"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' "&filtro&" OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "&filtro&""
		strtop = "1000"
		criterios = " Order by nm_fornecedor ASC"
		'response.write("1.2")
		
		elseif tipo = "pendentes" AND org = "unid"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		strtop = "1000"
		criterios = " Order by nm_unidade ASC"
		'response.write("1.3")
		
		elseif tipo = "pendentes" AND org = "os"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		strtop = "1000"
		criterios = " Order by num_os ASC"
		'response.write("1.4")

		elseif tipo = "pendentes" AND org = "equip"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		strtop = "1000"
		criterios = " Order by nm_equipamento ASC"
		'response.write("1.5")
		
		elseif tipo = "pendentes" AND org = "marca"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		strtop = "1000"
		criterios = " Order by nm_marca ASC"
		'response.write("1.6")
		
		elseif tipo = "pendentes" AND org = "aval"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		strtop = "1000"
		criterios = " Order by nm_avaliacao ASC"
		response.write("1.7")
		
		elseif tipo = "pendentes" AND org = "sit"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		strtop = "1000"
		criterios = " Order by nm_situacao ASC"
		'response.write("1.8")
		
		elseif tipo = "pendentes" AND org = "prev"  AND forn = "" Then
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' AND dt_prev_retorno IS NOT NULL OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' AND dt_prev_retorno IS NOT NULL"
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		strtop = "1000"
		criterios = " Order by dt_prev_retorno"
		'response.write("1.9")
		
		elseif tipo = "pendentes" AND org = "envio"  AND forn = "" Then
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' AND dt_prev_retorno IS NOT NULL OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' AND dt_prev_retorno IS NOT NULL"
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		strtop = "1000"
		criterios = " Order by dt_manut_orc"
		'response.write("2.0")
		
		elseif tipo = "relatorio" AND org = "" AND forn = "" Then
		nm_fornecedor = "'Erwin Guth'"
		mes_entrada = "'4'"
		ano_entrada = "'2007'" 
		
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		condicoes = " Where nm_fornecedor = '"&nm_fornecedor&"' AND MONTH(dt_entrada) > '"&mes_entrada&"' AND YEAR(dt_entrada) = '"&ano_entrada&"'"
		strtop = "1000"
		criterios = " Order by num_os desc"
		'response.write("10.0")

		Elseif tipo = "encerradas" Then
		condicoes = " Where fecha = 1 "
		'response.write("3.1")

		Elseif tipo = "todas" Then
		condicoes = " "
		'response.write("4.1")
		response.write("Ordem padrão")

		Else
		condicoes = " "
		'response.write("4.1")
		response.write("Ordem padrão")

		End if
		txt = "nada"
		selecao = " TOP "&strtop&" * "
		Tabela = " [dbo].[VIEW_os_lista] "

%>
<table align="center" border="0">
	<tr>
		<td>
			<table width="650" border="0" cellspacing="2" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef">
				<%
				'	dia = day("31/05/2007")
				'	mes = month("31/05/2007")
				'	ano = year("31/05/2007")
				
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
				
				'***** Contador Geral *******************************
				xsql_contador ="up_contador_geral @tabela='"&tabela&"', @condicoes='"&condicoes&"'"
				Set rs_contador = dbconn.execute(xsql_contador)
				
				if Not rs_contador.EOF Then
				nreg = rs_contador("nreg")
				End if
				'***** Fim do contador *******************************
				
					If mov = "retr" Then
						if tipo = "pendentes" OR tipo = "" Then
						condicoes = " Where num_os > "&regf&" AND nm_solicitante <> '"&solicitante&"' OR fecha IS NULL OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "
						'condicoes = " Where num_os > "&regf&" AND fecha IS NULL OR fecha = 0 "
						Elseif tipo = "encerradas" Then
						condicoes = " Where num_os > "&regf&" AND fecha = 1 "
						Elseif tipo = "todas" Then
						condicoes = " Where num_os > "&regf&""
						End if
						txt = "retr"
					Elseif mov = "avan" Then
						if tipo = "pendentes" OR tipo = "" Then
						condicoes = " Where num_os < "&regl&" AND nm_solicitante <> '"&solicitante&"' OR fecha IS NULL OR fecha = 0 AND nm_solicitante <> '"&solicitante&"' "
						'condicoes = " Where num_os < "&regl&" AND fecha IS NULL OR fecha = 0 "
						Elseif tipo = "encerradas" Then
						condicoes = " Where num_os < "&regl&" AND fecha = 1 "
						Elseif tipo = "todas" Then
						condicoes = " Where num_os < "&regl&""
						End if
						txt = "avan"
					End if
				
				%>
					
				<tr id="no_print">		
						<td class="txt_cinza" colspan="12">
						<b>Manutenção &raquo; - <font color="red">Listagem <%=tipo%></font></b></td>
						<td class="txt_cinza" align="right" valign="top">O.S.:&nbsp;</td>
						<td class="txt_cinza" colspan="2">
						<form action="manut_os_andamento.asp" name="busca" id="busca">
						<input type="text" name="cd_codigo" size="6" maxlength="6">
						<input type="hidden" name="campo" value="num_os">
						<input type="hidden" name="ordem" value="1">		
						<input type="submit" value="ok">
						</form>
						</td>
					</tr>
					<tr id="no_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manut_os_nova.asp" name="busca" id="busca">
					<tr id="no_print"><td colspan="100%">NOVA O.S.:<input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok"></td></tr>
					</form>
					<tr id="no_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<tr id="no_print">
						<td colspan="9" class="textopequeno" bgcolor="#f5f5f5">
							<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
						    / <a href="manutencao.asp?strtop=20&tipo=pendentes">Pendentes</a> 
							/ <a href="manutencao.asp?strtop=20&tipo=encerradas">Encerradas</a>
							/ <a href="manutencao.asp?strtop=20&tipo=todas">Todas</a>
							/ <a href="relatorio_manut.asp">Relatório</a>
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
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=os&ver=<%=ver%>"><font color="<%=color_os%>">O.S.</font></a> /
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=equip&ver=<%=ver%>"><font color="<%=color_equip%>">Equip/Instr.</font></a> /
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=marca&ver=<%=ver%>"><font color="<%=color_marca%>">Marca</font></a> /
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=unid&ver=<%=ver%>"><font color="<%=color_unid%>">Unidade</font></a> / 
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=fornec&ver=<%=ver%>"><font color="<%=color_fornec%>">Ass. Técnica</font></a> / 
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=prev&ver=<%=ver%>"><font color="<%=color_prev%>">Prev. Retorno</font></a> /  
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=aval&ver=<%=ver%>"><font color="<%=color_aval%>">Avaliacao</font></a> / 
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=sit&ver=<%=ver%>"><font color="<%=color_sit%>">Situação</font></a> /
						</td>
				</tr>
				<!--tr><td colspan="100%">Select: "up_listagem <%=selecao%>, <%=tabela%>, @condicoes='"&condicoes&"', @criterios='"&criterios&"'"</td></tr-->
					<tr><td colspan="10" class="textopequeno">&nbsp;</td></tr>
				<%'end if%>
				<tr id="ok_print"><td colspan="100%" align="center" class="txt_titulo"><b>PENDÊNCIAS EM: <%=dia%>/<%=mes%>/<%=ano%></b>&nbsp;</td></tr>
				<tr>
					<td><img src="imagens/px.gif" alt="" width="1" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="1" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="1" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="1" height="2" border="0"></td>
					<td></td>
					<td><img src="imagens/px.gif" alt="" width="135" height="2" border="0"></td>
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
					<td><img src="imagens/px.gif" alt="" width="60" height=2" border="0"></td>
					<td></td>
					<td><img src="imagens/blackdot.gif" alt="" width="40" height=2" border="0"></td>
				</tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr>
				
				<!--tr id="ok_print"><td colspan="100%" align="center" class="txt_titulo"><b>RELATÓRIO </b><br>&nbsp;</td></tr>
				<tr id="ok_print"><td colspan="100%" align="right" class="textopequeno">Data da emissão: <%=dia%>/<%=mes%>/<%=ano%></td></tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr-->
				
					<tr bgcolor="darkgray">
						<td><b>Nº</b></td>
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><b><a href="manutencao.asp?strtop=20&tipo=pendentes&org=os&ver=<%=ver%>">O.S.</a></b></td>
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><b><a href="manutencao.asp?strtop=20&tipo=pendentes&ver=<%=ver%>">PROVID.</a></b></td>	
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><b><a href="manutencao.asp?strtop=20&tipo=pendentes&org=envio&ver=<%=ver%>">DATA ENVIO</a></b></td>
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?strtop=20&tipo=pendentes&org=equipv"><b>EQUIPAMENTO/<br>INSTRUMENTAL.</b></a></td>	
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><b>PAT</b></td>	
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<!--td><img src="px.gif" alt="" width="1" height="1" border="0"><b>PATRIM.</b></td>	
						<td><img src="imagens/px.gif" width="1" height="1"></td-->
						<td><a href="manutencao.asp?strtop=20&tipo=pendentes&org=marca&ver=<%=ver%>"><b>MARCA.</b></a></td>	
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?strtop=20&tipo=pendentes&org=unid&ver=<%=ver%>"><b>UNID.</b></a></td>		
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<!--td><img src="px.gif" alt="" width="1" height="1" border="0">&nbsp;&nbsp;&nbsp;</td/-->
						<td><a href="manutencao.asp?strtop=20&tipo=pendentes&org=fornec&ver=<%=ver%>"><b>ASSIST. TÉCNICA</b></a></td>
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?strtop=20&tipo=pendentes&org=prev&ver=<%=ver%>"><b>PREV. RETORNO</b></a></td>
						<td><img src="imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?strtop=20&tipo=pendentes&org=aval&ver=<%=ver%>"><b>AVALIAÇÃO</b></a></td>
						<td><img src="../imagens/px.gif" width="3" height="1"></td>
						<td><a href="manutencao.asp?strtop=20&tipo=pendentes&org=sit&ver=<%=ver%>"><b>SITUAÇÃO</b></a></td>
				</tr>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr>
								
				<%
				'if ordem = "0" Then
				'	ordem = "ASC"
				'Else
				'	ordem = "DESC"
				'End if
				
				'If campo= "" AND cd_codigo = "" Then
				'campo = "num_os"
				'formulario = "0"
				'Elseif campo <> "" Then
				'formulario = "0"
				'End If
				
				'if strtop = "" Then
				'strtop = "20"
				'Elseif strtop = "todos" then 
				'strtop = "1000"
				'End if
				
				'compare = "="
				'strmove = request("strmove")
				'strreg = request("strreg")
				
				'if strmove = "1" Then
				'sinal = ">"
				'Elseif strmove = "0" then
				'sinal = "<"
				'End if
				
				'If strreg = "" And situacao = "" Then
				'stroutros = " fecha IS NULL" 'para listar todos os registros
				'Elseif strsituacao = "" Then
				'Elseif strreg = "" And situacao = "" Then
				'stroutros = " fecha IS NULL AND num_os "&sinal&" "&strreg&"" 'para listar top 20
				'Elseif strsituacao = "1" Then
				'stroutros = " fecha = 1 AND num_os "&sinal&" "&strreg&"" 
				'End if
				
				'If strreg = "" AND strsituacao = "" Then
				'stroutros = " fecha IS NULL" 'lista as pendencias em 1 página
				'sit = "1"
				'Elseif strreg <> "" AND strsituacao = "" Then
				' stroutros = " fecha IS NULL AND num_os "&sinal&" "&strreg&"" 'lista as pendencias em páginas com 20 registros
				'sit = "2"
				'Elseif strreg = "" AND strsituacao = "1" Then
				'stroutros = " fecha = 1 AND num_os "&sinal&" "&strreg&"" 'Lista as O.S. encerradas em páginas de 20 registros
				'sit = "3"
				'End if
				
				
				'xsql ="up_os_lista @cd_codigo='"&cd_codigo&"',@campo='"&campo&"', @ordem='"&ordem&"', @top='"&strtop&"',@compare='"&compare&"', @outros='"&stroutros&"'"
				xsql = "up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
				Set rs = dbconn.execute(xsql)
				
				if num = 0 Then 
				num = 1
				End if
				
				
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
				cd_avaliacao = rs("cd_avaliacao")
				cd_liberacao = rs("cd_liberacao")
				cd_codigo = rs("cd_codigo")
				nm_situacao = rs("nm_situacao")
				
				dt_manut_orc = rs("dt_manut_orc")
					manut_dia = zero(day(dt_manut_orc))					'manut_dia = day(dt_manut_orc)					'		if 	manut_dia < 10 then											'manut_dia = "0"&manut_dia											'end if
					manut_mes = zero(month(dt_manut_orc))					'manut_mes = month(dt_manut_orc)											'if 	manut_mes < 10 then											'manut_mes = "0"&manut_mes											'end if
					manut_ano = year(dt_manut_orc)

				dt_manut_orc = manut_dia&"/"&manut_mes&"/"&manut_ano
				
				dt_prev_retorno = rs("dt_prev_retorno")
					if year(dt_prev_retorno) = 1950 Then
						'dt_prev_retorno = ""
					end if
						if dt_prev_retorno <> "" Then
							dia_prev_retorno = day(dt_prev_retorno)
							mes_prev_retorno = month(dt_prev_retorno)
							ano_prev_retorno = year(dt_prev_retorno)
								dt_prev_retorno = dia_prev_retorno&"/"&mes_prev_retorno&"/"&ano_prev_retorno
						end if
						
						if dt_prev_retorno <> "" Then
						'dt_diferenca = DateDiff("d", dt_manut_orc, dt_prev_retorno) 
						dt_diferenca = DateDiff("d", dt_manut_orc, hoje) 
						end if
						difer = rs("hoje")
						prazo = rs("prazo")
				'***********************************************
				'	 Verifica a data da previsão de retorno
				'************************************************
				'Transfoma as datas em números
					prev_retorno = ano_prev_retorno&zero(mes_prev_retorno)&zero(dia_prev_retorno)
					Data_hoje = ano&mes&dia
						
						
						'if dt_diferenca > int(60) then
						'maior = " - É"
						'cor_aviso = "green"
						'Else
						'cor_aviso = "black"
						'end if
						
						if nm_situacao = "Pronto para retirar" Then
						cor_aviso = "#339966"
						Elseif prev_retorno < data_hoje AND prev_retorno <> ""  AND prev_retorno <> "19500101" AND fecha <> 1 then
						cor_aviso = "red"
						Elseif dt_diferenca >= int(30)  and fecha <> 1 then
						''maior = " - É"
						cor_aviso = "#660099"
						Else
						cor_aviso = "black"
						End if
				
				
				
				
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
				<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('manut_os_andamento.asp?cd_codigo=<%=cd_codigo%>&campo=cd_codigo');" onmouseout="mOut(this,'');">
						<td valign="top" align="center" bgcolor="darkgray"><b><%=num%></b></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%=num_os%></font></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%'=dt_os%><%=dt_entrada%></font></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%'=dt_entrada%><%=dt_manut_orc%></font></td>
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
						<td valign="top"><font color="<%=cor_aviso%>"><%=Left(nm_fornecedor,12)%></font></td>		
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%=dt_prev_retorno%></font></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%=nm_avaliacao%></font></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><font color="<%=cor_aviso%>"><%=nm_situacao%></font></td>
						<%if strsession = "2" Then%>
						<td valign="top" id="no_print"><font color="<%=cor_aviso%>"><%=dt_prev_retorno%></font></td>
						<td valign="top" id="no_print"><%=difer%>-<%=prazo%><%'=prev_retorno&"-"%><%'=data_hoje&"="%><%'=dif_datas&"-"%><%'=cor_aviso%></td>
						<!--td id="no_print"><%'=dt_diferenca%><%'=maior%></td-->
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
				
				num = num + 1
				prev_retorno = ""
				nm_situacao = "" 
				cor_aviso = ""
				maior = ""
				
				loop
				
				%>
				<tr class="textopequeno" id="no_print"><td colspan="7">&nbsp;</td></tr>
					<tr id="no_print">
						<td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
							<!--/ <a href="ordemservico_geracao.asp">Nova O.S. </a>
						    / <a href="ordemservico_listagem.asp">Listagem </a> 
							</ <a href="manutencao_pend.asp?sig=vazio&campo=dt_encerramento">Pendentes<%=fecha%></a-->
							< <%=nreg%> registros em <%=paginas%> páginas ></td>
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
						<b><a href="manutencao.asp?strtop=<%=strtop%>&tipo=<%=tipo%>&regf=<%=regf%>&regl=<%=regl%>&mov=retr&pagina=<%'=pagina%>&num=1"">Início</a></b>
						<%End if%>
						&nbsp;&nbsp;&nbsp;
						<%If pagina = paginas Then%>
						<b><font color="#b7b7b7">Próxima</font></b>
						<%else%>
						<b><a href="manutencao.asp?strtop=<%=strtop%>&tipo=<%=tipo%>&regf=<%=regf%>&regl=<%=regl%>&mov=avan&pagina=<%=pagina%>&num=<%=num%>">Próxima</a> </b>
						<%End if%>
						<td class="textopequeno" bgcolor="#f5f5f5" colspan="2">Pág: (<%=pagina%>)</td>
					</tr>
					<tr class="textopequeno" id="no_print"><td colspan="10"><!--&nbsp;<%=regf%>/<%=regl%> - <%=condicoes%><%=criterios%> *** <%=txt%>--></td></tr>
					<%'if strsession = "2" Then%>
					<tr id="ok_print"><td colspan="100%">
					<%do while not num = "45"%> <br>
					<%num = num + 1
					loop%>
					
						<table border="1" cellpadding="0" cellspacing="1" bordercolor="#808080" bordercolordark="#808080" bordercolorlight="#808080">
							<tr><td>&nbsp;LEGENDA</td></tr>
							<tr>
								<td>&nbsp;<font color="black"><b>O</b></font> - Em andamento<br>
								&nbsp;<font color="red"><b>O</b></font> - Prazo de prev. de entrega vencido<br>
								&nbsp;<font color="#660099"><b>O</b></font> - Crítico<br>
								&nbsp;<font color="#339966"><b>O</b></font> - Retirar fornecedor
								
								</td>
							</tr>
						</table>
					</td></tr>
					<%'if strsession = 2 Then%>
					<tr id="no_print"><td colspan="100%"><a href="manutencao.asp?ver=">Voltar</a></a></td></tr>
					<%'end if%>
				</table></td></tr></table>
</BODY>
</HTML>
