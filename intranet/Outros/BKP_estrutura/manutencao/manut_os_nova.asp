


<script language="JavaScript" type="text/javascript" src="../js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../js/formValidator.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../js/mascarainput.js"></SCRIPT>
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
<script language="JavaScript">
function validar_os(shipinfo){
	if (shipinfo.dt_os.value.length==""){window.alert ("A data de abertura da O.S. não foi preenchida.");shipinfo.dt_os.focus();return (false);
	} else {
	if (shipinfo.nm_solicitante.value.length==""){window.alert ("O solicitante não foi informado.");shipinfo.nm_solicitante.focus();return (false);
	} else {
	if (shipinfo.nm_unidade.value.length==""){window.alert ("Selecione uma unidade.");shipinfo.nm_unidade.focus();return (false);
	} else {
	if (shipinfo.num_qtd.value.length==""){window.alert ("Informe a quantidade.");shipinfo.num_qtd.focus();return (false);
	} else {
	if (shipinfo.cd_especialidade.value.length==""){window.alert ("Selecione uma especialidade.");shipinfo.cd_especialidade.focus();return (false);
	} else {
	if (shipinfo.cd_equipamento.value.length==""){window.alert ("Selecione um equipamento.");shipinfo.cd_equipamento.focus();return (false);
	} else {
	if (shipinfo.cd_marca.value.length==""){window.alert ("Selecione uma marca.");shipinfo.cd_marca.focus();return (false);
	} else {
	if (shipinfo.motivo.value.length==""){window.alert ("Qual motivo da solicitação?.");shipinfo.motivo.focus();return (false);
	} else {
	if (shipinfo.dt_entrada.value.length==""){window.alert ("informe a data da providencia.");shipinfo.dt_entrada.focus();return (false);
	} else {
	if (shipinfo.cd_avaliacao.value.length==""){window.alert ("informe o andamento da providencia.");shipinfo.cd_avaliacao.focus();return (false);
	} else {
	if (shipinfo.cd_fornecedor.value.length==""){window.alert ("informe o fornecedor.");shipinfo.cd_fornecedor.focus();return (false);
	} else {
	}}}}}}}}}}}return (true);}
</script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "num_os"; // nome do primeiro campo do site
netscape = "";
ver = navigator.appVersion; len = ver.length;
for(iln = 0; iln < len; iln++) if (ver.charAt(iln) == "(") break;
	netscape = (ver.charAt(iln+1).toUpperCase() != "C");
	function keyDown(DnEvents) {
		// ve quando e o netscape ou IE
		k = (netscape) ? DnEvents.which : window.event.keyCode;
		if (k == 13) { // preciona tecla enter
		if (nextfield == 'done') {
			//alert("viu como funciona?");
			eval('document.form.' + nextfield + '.focus()');
			return false;
			//return true; // envia quando termina os campos
		} else {
			// se existem mais campos vai para o proximo
			eval('document.form.' + nextfield + '.focus()');
			return false;
		}
	}
}
document.onkeydown = keyDown; // work together to analyze keystrokes
if (netscape) document.captureEvents(Event.KEYDOWN|Event.KEYUP);
// End -->
function foco(){
document.form.cd_patrimonio.focus(); }
</script>
<%esconder = "1"%>
<BODY> 



<br>
<table align="center" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<table width="600" border="1" cellspacing="1" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno">
				<tr>		
					<td class="txt_cinza" colspan="6">
					<b>Manutenção &raquo; - <font color="red">Listagem</font></b><br><br></td>
					<!--td class="txt_cinza" align="right" valign="top">Busca O.S.: &nbsp;</td>
					<td class="txt_cinza" colspan="2">
					<form action="ordemservico_andamento.asp" name="busca" id="busca">
					<input type="text" name="cd_codigo" size="6" maxlength="6">
					<input type="hidden" name="campo" value="num_os">
					<input type="hidden" name="ordem" value="1">		
					<input type="submit" value="ok">
					</form>
					</td-->
				</tr>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manut_os_nova.asp" name="busca" id="busca">
				<tr><td colspan="100%">NOVA O.S.:<input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok"></td></tr>
					</form>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					
				<tr>
				    <td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
						<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
					    / <a href="manutencao.asp?strtop=20&tipo=pendentes">Pendentes</a> 
						/ <a href="ordemservico_listagem.asp?strtop=20&tipo=encerradas">Encerradas</a>
						/ <a href="ordemservico_listagem.asp?strtop=20&tipo=todas">Todas</a>
					</td>
				</tr>					
				<%
				
				cd_codigo = request("cd_codigo")
				num_os = request("num_os")
				campo = request("campo")
				ordem = request("ordem")
				list = request("list")
				lista = "1"
			
				if ordem = "" Then
				ordem = "1"
				Elseif ordem = "1" Then
				ordem = "0"
				End if
			
				if ordem = "0" Or ordem = "" Then
				ordem = "DESC"
				Else
				ordem = "ASC"
				End if
			
				If campo= "" AND cd_codigo = "" Then
				campo = "num_os"
				formulario = "0"
				End If
			
				compare = "="
				strtop = "100 percent"
				outros = ""
				acao = "1"
				tipo = "cadastra_os"
				
				
'*************************************
'*** Verifica a existência da O.S. ***
'*************************************
if num_os <> "" Then
	xsql = "SELECT * FROM TBL_OrdemServico where num_os="&num_os&""
	Set rs_verifica = dbconn.execute(xsql)
		do while not rs_verifica.EOF
		num_os_verifica = rs_verifica("num_os")
		cd_codigo_verifica = rs_verifica("cd_codigo")
		rs_verifica.movenext
		loop
			if num_os_verifica = int(num_os) then 'Redireciona à O.S. existente
				response.redirect("manut_os_andamento.asp?cd_codigo="&cd_codigo_verifica&"&campo=cd_codigo")
			End if
end if
'*************************************



				
				if cd_codigo <> "" Then
				xsql ="up_os_lista @cd_codigo='"&cd_codigo&"',@campo='"&campo&"', @ordem='"&ordem&"', @top='"&strtop&"',@compare='"&compare&"', @outros='"&stroutros&"'"
				Set rs = dbconn.execute(xsql)
				
				num_os = rs("num_os")
				dt_os = rs("dt_os")
				nm_unidade = rs("nm_unidade")
				nm_especialidade = rs("nm_especialidade")
				nm_equipamento = rs("nm_equipamento")
				cd_ns = rs("cd_ns")
				cd_patrimonio = rs("cd_patrimonio")
			
				dt_entrada = rs("dt_entrada")
				cd_avaliacao = rs("cd_avaliacao")
				'nm_avaliacao = rs("nm_avaliacao")
				nm_fornecedor = rs("nm_fornecedor")
				cd_fornecedor = rs("cd_fornecedor")
				cd_liberacao = rs("cd_liberacao")
			
				cd_codigo = rs("cd_codigo")
			
				num_qtd = rs("num_qtd")
				nm_marca = rs("nm_marca")
				motivo = rs("motivo")
				nm_solicitante = rs("nm_solicitante")
				observacoes = rs("observacoes")
				
				acao = "2"
				tipo = "modifica_os"
				sequencia = "1"
				fecha = rs("fecha")
				End if%>
			<!--Início da etapa 1.1-->
				<tr>			
					<form method="post" action="manutencao/manut_os_acao.asp" name="form"  id="form" onsubmit="return validar_os(document.form);">
						<input type="hidden" name="acao" value="<%=acao%>">
						<input type="hidden" name="tipo" value="<%=tipo%>">
						<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
						<input type="hidden" name="sequencia" value="1">
						<input type="hidden" name="cd_liberacao" value="<%=session("nm_nome")%>">
				<tr bgcolor="#b4b4b4">
					<td  colspan="100%" align="center" class="textopadrao">SOLICITAÇÃO</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">O.S.</td>
					<td colspan="4"><input type="text" name="num_os" size="5" maxlength="4" class="inputs" value="<%=num_os%>" onFocus="nextfield ='dt_os';">&nbsp;
					<!--input type="text" name="digito" size="1" maxlength="3" class="inputs" value=""--></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Data</td>
					<td colspan="4"><input type="text" name="dt_os" size="11" maxlength="50" class="inputs" value="<%=dt_os%>" onFocus="nextfield ='nm_solicitante';"></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Solicitante</td>
					<td colspan="4"><input type="text" name="nm_solicitante" size="40" maxlength="40" class="inputs" value="<%=nm_solicitante%>" onFocus="nextfield ='nm_unidade';"></td>							
				
				<!--tr>
					<td align="center" bgcolor="#f5f5f5">Unidade</td>
					<td bgcolor="#f5f5f5" valign="top">
						<%
'						salto = 1
'						Do while not rs_uni.EOF
'						cd_uni = rs_uni("cd_codigo")
'						nm_sigla = rs_uni("nm_sigla")
'						nome_unidade = rs_uni("nm_unidade")
'						
'						if nm_unidade = nm_sigla Then
'						ck_uni = "checked"
'						end if%>
						<input type="radio" name="nm_unidade" value="<%'=nm_sigla%>" <%'=ck_uni%>><%'=nm_sigla%> - <%'=nome_unidade%> <%''=divisao%><%''=salto%>&nbsp;<br><%'if int(divisao) = salto Then%></td><td bgcolor="#f5f5f5" valign="top"><%'end if%>
						<%'rs_uni.movenext
'						ck_uni = ""
'						salto = salto + 1
'						loop
						%>
					</td>
					<td colspan="2"bgcolor="#f5f5f5" valign="bottom" align="center"><!--a href="#" onclick="window.open('adm/equipamento_cad.asp', 'Janela','width=430,height=500')">+<p></p--></td>
					
				</tr-->
				<tr>
					<td align="center" bgcolor="#f5f5f5">Unidade</td>
						<%' ***** Conta quantas unidades existem ***
						selecao = " Count(cd_codigo) as conta"
						tabela = " TBL_unidades"
						condicoes = " WHERE cd_codigo <> 6 AND cd_status > 0 "
						'criterios = ""' group by cd_codigo"', nm_unidade, nm_sigla Order by nm_unidade"
						criterios = " group by nm_sigla, nm_unidade, cd_codigo Order by nm_sigla"
						
						strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
						Set rs_uni = dbconn.execute(strsql)
						conta = rs_uni("conta")
						
						divisao = conta/2
						
						'****** Mostra as unidades ***
						selecao = " * "
						tabela = " TBL_unidades"
						condicoes = " WHERE cd_codigo <> 6 AND cd_status > 0 "
						criterios = " Order by nm_unidade"

						strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
										  	
						Set rs_uni = dbconn.execute(strsql)%>
					<td bgcolor="#f5f5f5" valign="top">
					<select name="nm_unidade" class="inputs" onFocus="nextfield ='num_qtd';">
					<option value=""></option>
						<%
						Do while not rs_uni.EOF
						cd_uni = rs_uni("cd_codigo")
						nm_sigla = rs_uni("nm_sigla")
						nome_unidade = rs_uni("nm_unidade")
						
						if nm_unidade = nm_sigla Then
						ck_uni = "selected"
						end if%>
						
					<option value="<%=nm_sigla%>" <%=ck_uni%>><%=nm_sigla%> - <%=nome_unidade%></option>
						<%rs_uni.movenext
						ck_uni = ""
						'salto = salto + 1
						loop
						%>
					</select>
			

					</td>
					<td colspan="2"bgcolor="#f5f5f5" valign="bottom" align="center"><!--a href="#" onclick="window.open('adm/equipamento_cad.asp', 'Janela','width=430,height=500')">+<p></p--></td>
				</tr>
				<tr bgcolor="#f5f5f5">	
					<td  align="center">Qtd	</td>
					<td><input type="text" name="num_qtd" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='cd_especialidade';"></td>
					<td align="center">Especialidade</td>
								<%
											selecao = " top 100 percent * "
											tabela = " TBL_especialidades"
											condicoes = ""
											criterios = " Order by nm_especialidade"
											
											strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
										  	
							  	Set rs_espec = dbconn.execute(strsql)%>
					<td colspan="2">	<select name="cd_especialidade" onFocus="nextfield ='cd_equipamento';">
										<option class="inputs" value="">Selecione</option>
										<%Do while Not rs_espec.EOF%><%especialidade=rs_espec("nm_especialidade")%><%if nm_especialidade=especialidade Then%><%ck_espec="selected"%><%else ck_espec=""%><%end if%>
										<option class="inputs" value="<%=rs_espec("cd_codigo")%>" <%=ck_espec%>><%=rs_espec("nm_especialidade")%></option>	<%rs_espec.movenext
									  	Loop%>
					</td>
											
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center">Equip./instr.</td>
						<%
											selecao = " top 100 percent * "
											tabela = " TBL_equipamento"
											condicoes = ""
											criterios = " Order by nm_equipamento"
											
											strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
										  	Set rs_equip = dbconn.execute(strsql)%>
											<td colspan="2"><select name="cd_equipamento" onFocus="nextfield ='cd_marca';">
											<option value="" class="inputs">Selecione</option>
											<%Do while Not rs_equip.EOF%><%equip=rs_equip("nm_equipamento")%><%if nm_equipamento=equip Then%><%ck_equip="selected"%><%else ck_equip=""%><%end if%>
											<option class="inputs" value="<%=rs_equip("cd_codigo")%>" <%=ck_equip%>><%=rs_equip("nm_equipamento")%></option>	<%rs_equip.movenext
										  	Loop%>
					<td  align="center">Marca</td>
						<%
											'selecao = " top 100 percent * "
											tabela = " TBL_marca"
											condicoes = ""
											criterios = " Order by nm_marca"
											
											strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
										  	
					  	Set rs_marca = dbconn.execute(strsql)%>
					<td><select name="cd_marca" class="inputs" onFocus="nextfield ='cd_ns';">
					<option value="">Selecione</option>
						<%Do while Not rs_marca.EOF %><%marca=rs_marca("nm_marca")%><%if nm_marca=marca Then%><%ck_marca="selected"%><%else ck_marca=""%><%end if%>
						<option value="<%=rs_marca("cd_codigo")%>" <%=ck_marca%>><%=rs_marca("nm_marca")%></option><%rs_marca.movenext
					  	Loop%>
					</td>	
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">N/S</td>
					<td><input type="text" name="cd_ns" maxlength="10" size="15"  value="<%=cd_ns%>" class="inputs" onFocus="nextfield ='cd_patrimonio';"></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Patrimônio</td>
					<%'selecao = " top 100 percent * "
					tabela = " View_patrimonio_lista"
					condicoes = " Where dt_descarte Is Null  "
					criterios = " Order by nm_equipamento"
					strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
					Set rs_pat = dbconn.execute(strsql)%>
					<td><input type="text" name="cd_patrimonio" maxlength="10" size="15" value="<%=cd_patrimonio%>" class="inputs" onFocus="nextfield ='motivo';">
						
						<!--select name="cd_patrimonio" class="inputs"><option value="">Selecione</option>
						<%'Do while Not rs_pat.EOF %>
						<%'patrimonio=rs_pat("nm_patrimonio")%>
						<%'if nm_pat=patrimonio Then%>
						<%'ck_pat="selected"%>
						<%'else ck_pat=""%>
						<%'end if%>
						<option value="<%'=rs_pat("cd_codigo")%>" <%'=ck_pat%>><%'=rs_pat("cd_patrimonio")%></option><%'rs_pat.movenext
					  	'Loop%>-->
					</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Motivo</td>
					<td colspan="4">
						<textarea cols="70" rows="2" name="motivo" class="inputs" onFocus="nextfield ='dt_entrada';"><%=motivo%></textarea>
					</td>
				</tr>
			<!--Fim da etapa 1.1-->			
			<!--Início da etapa 1.2-->				
				<tr><td colspan="100%"><!--Separador de etapas-->&nbsp;</td></tr>
								
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">PROVIDÊNCIA</td></tr>
				<tr bgcolor="#f5f5f5">
				  	<td  align="center">Data</td>
					<td><input type="text" name="dt_entrada" size="11" maxlength="50" class="inputs" value="<%=dt_entrada%>" onFocus="nextfield ='cd_avaliacao';"></td>
					<td  align="center">Avaliação</td>
					<td colspan="2"><select name="cd_avaliacao" class="inputs" onFocus="nextfield ='cd_fornecedor';">
						<option value="">Selecione</option>
						<%
						'****** Mostra as opções ***
						selecao = " * "
						tabela = " TBL_avaliacao"
						condicoes = ""
						criterios = " Order by nm_avaliacao"

						strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
						Set rs_aval = dbconn.execute(strsql)%>
						
						<%
						Do while not rs_aval.EOF
						cd_aval = rs_aval("cd_codigo")
						nm_avaliacao = rs_aval("nm_avaliacao")
						
						'if cd_unidade = cd_aval Then
						if cd_avaliacao = cd_aval Then
						ck_aval = "selected"
						end if%>
						<option value="<%=cd_aval%>" <%=ck_aval%>><%=nm_avaliacao%></option>
						<%rs_aval.movenext
						ck_aval = ""
						loop
						%>
					 
					</td>
				</tr>
				<tr bgcolor="#f5f5f5"><td  align="center">Encaminhar para</td>
						<%strsql ="TBL_Fornecedor"
					  	Set rs_forn = dbconn.execute(strsql)%>
					<td colspan="4"><select name="cd_fornecedor" class="inputs" onFocus="nextfield ='observacoes';">
						<option value="">
						<%Do while Not rs_forn.EOF%><%fornecedor=rs_forn("nm_fornecedor")%><%if nm_fornecedor=fornecedor Then%><%ck_forn="selected"%><%else ck_forn=""%><%end if%>
						<option value="<%=rs_forn("cd_codigo")%>" <%=ck_forn%>><%=rs_forn("nm_fornecedor")%></option><%rs_forn.movenext
					 	Loop%></td>
				</tr>
				<tr bgcolor="#f5f5f5"><td  align="center">Liberado por</td>
					<td colspan="4"><%=session("nm_nome")%></td>
				</tr>	
				<tr bgcolor="#f5f5f5">
					<td align="center">observações<br></td>
					<td colspan="4">
						<textarea cols="40" rows="2" name="observacoes" class="inputs" onFocus="nextfield ='dt_manut_orc';"><%=observacoes%></textarea></td>
				</tr>
				<tr><td colspan="100%"><!--Separador de etapas-->&nbsp;</td></tr>	
				</tr>
				<!--Os novos campos começam aqui-->
				<%if esconder = "1" Then
				Else%>
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">MANUTENÇÃO / ORÇAMENTO</td></tr>
				<tr bgcolor="#f5f5f5">
					<td>Data manut. / orc.</td>
					<td><input type="text" name="dt_manut_orc" size="11" maxlength="50" class="inputs" value="<%=dt_entrada%>"  onFocus="nextfield ='nm_contato';"></td> 
					<td  align="center">Assist. Técnica / Fornecedor</td>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					
					<td>Contato</td><td><input type="text" name="nm_contato" size="15" class="inputs"  onFocus="nextfield ='dt_prev_resposta';"></td>
					<td  align="center">Previsão resposta</td>
					<td><input type="text" name="dt_prev_resposta" size="11" maxlength="20" class="inputs" value="<%=dt_prev_resposta%>"  onFocus="nextfield ='dt_resposta_orc';"></td>
				</tr>
				<tr><td colspan="100%"><!--Separador de etapas-->&nbsp;</td></tr>
				<tr>
					<td  align="center">Data da resposta do orçamento</td>
					<td colspan="2"><input type="text" name="dt_resposta_orc" size="11" maxlength="20" class="inputs" value="<%=dt_previsao%>"  onFocus="nextfield ='nm_orcamento';"></td>
					<td  align="center">Orçamento</td>
					<%If nm_orcamento = "0" Then%>
					<%ck_mo10 = "selected"%>
					<%elseif nm_orcamento = "1" Then%>
					<%ck_mo11 = "selected"%>
					<%elseif nm_orcamento = "2" Then%>
					<%ck_mo12 = "selected"%>
					<%End if%>
					
					<td><select name="nm_orcamento" class="inputs"  onFocus="nextfield ='num_os_assist';">
							<option value="1">Aguardando</option>
							<option value="2">Aprovado</option>
							<option value="0">Reprovado</option>																		
					</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center">O.S. Assist. Téc.</td>
					<td colspan="2"><input type="text" name="num_os_assist" size="11" maxlength="10" class="inputs" value="" onFocus="nextfield ='dt_prev_retorno';"></td>
					
					<td  align="center">Previsão retorno/entrega</td>
					<td><input type="text" name="dt_prev_retorno" size="11" maxlength="20" class="inputs" value="" onFocus="nextfield ='cd_status';"></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center">Situação</td>
					<%If cd_status = "1" Then%>
					<%ck_os11 = "selected"%>
					<%elseif cd_status = "2" Then%>
					<%ck_os12 = "selected"%>
					<%elseIf cd_status = "0" Then%>
					<%ck_os10 = "selected"%>
					<%End if%>
					<td colspan="4"><select name="cd_status" class="inputs" onFocus="nextfield ='comentarios';">
																<option value="1" <%=ck_os11%>>Aguardando resposta</option>
																<option value="2">Aguardando aprovação</option>
																<option value="0">Aguardando entrega</option>
																<option value="0">Aprovado</option>
																<option value="0">Reprovado</option>
																
																<option value="0">Aguardando</option>
																<option value="0">Em análise</option>
																<option value="0">Manutenção concluída</option>
																<option value="0">Disponível para retirar</option>
																<option value="0">Aguardando a entrega</option>
																<option value="0">Descarte</option>
																<option value="0">Sem conserto</option>
																<option value="0">Manutenção reprovada</option>
					</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Comentários<br></td>
					<td colspan="4">
						<textarea cols="40" rows="2" name="comentarios" class="inputs"></textarea onFocus="nextfield ='dt_retorno';"></td>
				</tr>
				<tr>
					<td  align="center">Data de entrega/retorno</td>
					<td colspan="2"><input type="text" name="dt_retorno" size="11" maxlength="20" class="inputs" value="" onFocus="nextfield ='dt_retorno_un';"></td>
				</tr>
				<tr>
					<td  align="center">Data de envio para Hospital</td>
					<td colspan="2"><input type="text" name="dt_retorno_un" size="11" maxlength="20" class="inputs" value=""></td>
				</tr>
				
				
				<%end if%>
				<tr>
					<td colspan="100%" align="center">
					<br><%if acao = "1" Then%><input  name="ok" type="submit" value="Cadastrar O.S."><%else%>
					<input type="submit" value="Modificar O.S.">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ordemservico_andamento.asp?cd_codigo=<%=cd_codigo%>&campo=<%=campo%>">Voltar</a><%End if%>
					
					</td></tr>
				<!--Fim da etapa 1.2-->
			
						</form>
							
			</table>

		</td>
	</tr>
</table>	
						
				

<SCRIPT LANGUAGE="javascript">
 {
    MaskInput(document.form.dt_os, "99/99/9999"); 
    MaskInput(document.form.dt_entrada, "99/99/9999");
    MaskInput(document.form.dt_encaminhado, "99/99/9999");
 }
</SCRIPT>



</BODY>
</HTML>



