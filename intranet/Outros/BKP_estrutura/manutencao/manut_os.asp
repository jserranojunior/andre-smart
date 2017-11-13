<%
'print = request("print")
'if print <> "" then
'border="0"
%>
<!--#include file="../includes/util.asp"-->

<%'end if%> 
<!--#include file="../includes/inc_open_connection.asp"--> 
<!--#include file="../includes/inc_area_restrita.asp"-->   
   
    <style type="text/css">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #000000; }  

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
<%'if print = "1" Then%>
	<!-->
	<script type="text/javascript">
	function printpage() {
	self.print();
	self.close()
	}
	</script>
	-->
<BODY onLoad="printpage()"-->
<%'Elseif print <> "1" Then%>
<!-->
	<script type="text/javascript">
	function printpage() {
	alert("impressão!")
	}
	</script>
	<BODY>
<%'end if%>
<!-- -->
<!--BODY onLoad="printpage()"--> 
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

criterios = " Order by num_os DESC "

		if tipo = "pendentes" AND org = ""  AND forn = "" OR tipo = ""  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
'		condicoes = " Where fecha IS NULL OR fecha = 0"
		strtop = "1000"
		'response.write("1.1")
		
		elseif tipo = "pendentes" AND org = "fornec"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
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
		'response.write("1.7")
		
		elseif tipo = "pendentes" AND org = "sit"  AND forn = "" Then
		condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		strtop = "1000"
		criterios = " Order by nm_situacao ASC"
		'response.write("1.8")
		
		elseif tipo = "relatorio" AND org = "" AND forn = "" Then
		nm_fornecedor = "'Erwin Guth'"
		mes_entrada = "'4'"
		ano_entrada = "'2007'" 
		
		'condicoes = " Where fecha IS NULL AND nm_solicitante <> '"&solicitante&"' OR fecha = 0 AND nm_solicitante <> '"&solicitante&"'"
		condicoes = " Where nm_fornecedor = '"&nm_fornecedor&"' AND MONTH(dt_entrada) > '"&mes_entrada&"' AND YEAR(dt_entrada) = '"&ano_entrada&"'"
		strtop = "1000"
		criterios = " Order by num_os ASC"
		'response.write("2.0")

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
<table align="center">
	<tr>
		<td>
			<table width="650" border="<%=border%>" cellspacing="2" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef">
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
					
				<%
				If print <> "" Then%>
					<!--Não mostra o cabaçalho com links-->
				<%Else
				%>
				<tr id="no_print"><td colspan="100%" align="center" class="txt_titulo"><b>PENDÊNCIAS EM: <%=dia%>/<%=mes%>/<%=ano%></b><br>&nbsp;</td></tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr>
				
				<tr>		
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
					<tr id="no_print">		
						<td class="txt_cinza" colspan="12">
						</td>
						<td class="txt_cinza" align="right" valign="top"><input type="button" onclick="window.open('manutencao/manut_os.asp?tipo=<%=tipo%>&org=<%=org%>&print=1','print','width=550, height=1, toolbar=0')"; id="btn_imprime" Value="Imprimir" />
																		<input type="button" onclick="window.open('manutencao/manut_os.asp?tipo=<%=tipo%>&org=<%=org%>&print=0 ','print','width=800, height=500, toolbar=1, scrollbar=1')"; id="btn_imprime" Value="Ver na tela" /></td>
						<td class="txt_cinza" colspan="2">
						</td>
					</tr>
					<tr id="no_print">
						<td colspan="9" class="textopequeno" bgcolor="#f5f5f5">
							/ <a href="manut_os_nova.asp">Nova O.S. </a>
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
				end if%>
				<tr id="no_print">
						<td colspan="15" class="textopequeno" bgcolor="#f5f5f5">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						    / <font color="red">Ordem de PENDENCIAS: </font>
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=os"><font color="<%=color_os%>">O.S.</font></a> /
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=equip"><font color="<%=color_equip%>">Equip/Instr.</a> /
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=marca"><font color="<%=color_marca%>">Marca</a> /
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=unid"><font color="<%=color_unid%>">Unidade</a> / 
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=fornec"><font color="<%=color_fornec%>">Ass. Técnica</a> / 
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=aval"><font color="<%=color_aval%>">Avaliacao</a> / 
							<a href="manutencao.asp?strtop=20&tipo=pendentes&org=sit"><font color="<%=color_sit%>">Situação</a> / 
							
							
							
						</td>
				</tr>
				
					<tr><td colspan="10" class="textopequeno">&nbsp;</td></tr>
				<%end if%>
				
				
				<%if print <> "" AND tipo = "pendentes" Then%>
				
				<tr><td colspan="100%" align="center" class="txt_titulo"><b>PENDÊNCIAS EM: <%=dia%>/<%=mes%>/<%=ano%></b><br>&nbsp;</td></tr>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr>
				<%Elseif print = "1" AND tipo = "relatorio" Then%>
				<tr><td colspan="100%" align="center" class="txt_titulo"><b>RELATÓRIO </b><br>&nbsp;</td></tr>
				<tr><td colspan="100%" align="right" class="textopequeno">Data da emissão: <%=dia%>/<%=mes%>/<%=ano%></td></tr>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr>
				<%end if%>
					<tr bgcolor="darkgray" class="print">
						<td><b>Nº</b></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="1" height="1" border="0"><b>O.S.</b></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="1" height="1" border="0"><b>PROVID.</b></td>	
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="1" height="1" border="0"><b>EQUIPAMENTO/<br>INSTRUMENTAL.</b></td>	
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="1" height="1" border="0"><b>N/S</b></td>	
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<!--td><img src="px.gif" alt="" width="1" height="1" border="0"><b>PATRIM.</b></td>	
						<td><img src="imagens/px.gif" width="1" height="1"></td-->
						<td><img src="px.gif" alt="" width="1" height="1" border="0"><b>MARCA.</b></td>	
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="1" height="1" border="0"><b>UNID.</b></td>		
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<!--td><img src="px.gif" alt="" width="1" height="1" border="0">&nbsp;&nbsp;&nbsp;</td/-->
						<td><img src="px.gif" alt="" width="1" height="1" border="0"><b><b>ASSIST. TÉCNICA</b></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="1" height="1" border="0"><b>DATA ENVIO</b></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="1" height="1" border="0"><b>AVALIAÇÃO</b></td>
						<td><img src="../imagens/px.gif" width="1" height="1"></td>
						<td>&nbsp;<b>SITUAÇÃO</b></td>
				</tr>
				<%if print = "1" Then%>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="2"></td></tr>
				<%end if%>
				
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
				'cd_patrimonio = rs("cd_patrimonio")
				nm_marca = rs("nm_marca")
				nm_unidade = rs("nm_unidade")
				dt_entrada = rs("dt_entrada")
				nm_fornecedor = rs("nm_fornecedor")
				cd_avaliacao = rs("cd_avaliacao")
				cd_liberacao = rs("cd_liberacao")
				cd_codigo = rs("cd_codigo")
				nm_situacao = rs("nm_situacao")
				
				dt_manut_orc = rs("dt_manut_orc")
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
				
				if nm_avaliacao = "manutencao" Then
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
				<tr class="textopequeno" onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('manut_os_andamento.asp?cd_codigo=<%=cd_codigo%>&campo=cd_codigo');" onmouseout="mOut(this,'');">
						<td valign="top" align="center" bgcolor="darkgray"><img src="px.gif" alt="" width="1" height="1" border="0"><b><%=num%></b></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%=num_os%></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%'=dt_os%><%=dt_entrada%></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%=Left(nm_equipamento,37)%></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%=(cd_ns)%></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<!--td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%=Left(nm_equipamento,37)%></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td-->
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%=Left(nm_marca,10)%></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%=nm_unidade%></td>		
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<!--td valign="top">&nbsp;</td/-->
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%=Left(nm_fornecedor,12)%></td>		
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%'=dt_entrada%><%=dt_manut_orc%></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%=nm_avaliacao%></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td valign="top"><img src="px.gif" alt="" width="1" height="1" border="0"><%=nm_situacao%></td>
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
				nm_situacao = "" 
				loop
				
				if print = "1" Then
				else
				%>
				<tr class="textopequeno"><td colspan="7">&nbsp;</td></tr>
					<tr><td><img src="px.gif" alt="" width="15" height="1" border="0"></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="30" height="1" border="0"></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="65" height="1" border="0"></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="201" height="1" border="0"></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="56" height="1" border="0"></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="45" height="1" border="0"></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="80" height="1" border="0"></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						
						<td><img src="px.gif" alt="" width="65" height="1" border="0"></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="75" height="1" border="0"></td>
						<td><img src="imagens/px.gif" width="1" height="1"></td>
						<td><img src="px.gif" alt="" width="120" height="1" border="0"></td>
					</tr>
				
					<tr>
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
						 'End if
				
				
						 %>
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
					<tr class="textopequeno"><td colspan="10"><!--&nbsp;<%=regf%>/<%=regl%> - <%=condicoes%><%=criterios%> *** <%=txt%>--></td></tr>
				<%end if%>
				</table>
				<%'End If%>
				
					
			</table>						
		</td>
	</tr>
</table>




</BODY>
</HTML>



